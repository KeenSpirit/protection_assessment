import sys
sys.path.append(r"\\Ecasd01\WksMgmt\PowerFactory\ScriptsDEV\PowerFactoryTyping")
import powerfactorytyping as pft
from importlib import reload
from fault_study import lines_results, analysis, floating_terminals as ft

reload(analysis)
reload(ft)
reload(lines_results)


def fault_study(app, feeder, site_name_map) -> tuple[list, list, list]:
    """

    :param app:
    :param feeder:
    :param site_name_map:
    :return:
    """

    # For each of the feeder devices, identify all downstream nodes
    devices_terminals, devices_loads, device_lines = get_downstream_objects(app, site_name_map)
    # Update all devices with the lists of downstream devices and upstreams devices
    us_devices, ds_devices = us_ds_device(devices_terminals)

    ds_capacity = get_ds_capacity(devices_loads)
    section_loads = get_device_sections(devices_loads)
    device_max_load, device_max_trs = get_section_max_tr(section_loads)
    devices_sections = get_device_sections(devices_terminals)
    floating_terms = ft.get_floating_terminals(feeder, devices_sections)

    app.PrintPlain(f'Getting {feeder.loc_name} maximum fault levels...')
    bound = 'Max'
    f_type = 'Ground'
    analysis.short_circuit(app, bound, f_type)
    # Max transformer data
    max_tr_pg_fls = terminal_fls(device_max_trs, f_type)
    sect_tr_pg_max = sect_fl_bound(max_tr_pg_fls, bound)
    # Terminal data
    pg_max_first_pass = terminal_fls(devices_sections, f_type)
    pg_max_all = append_floating_terms(app, pg_max_first_pass, floating_terms, bound, f_type)
    sect_pg_max = sect_fl_bound(pg_max_all, bound)

    f_type = '3 Phase'
    analysis.short_circuit(app, bound, f_type)
    # Max transformer data
    max_tr_p_fls = terminal_fls(device_max_trs, f_type)
    sect_tr_phase_max = sect_fl_bound(max_tr_p_fls, bound)
    # Terminal data
    phase_max_first_pass = terminal_fls(devices_sections, f_type)
    phase_max_all = append_floating_terms(app, phase_max_first_pass, floating_terms, bound, f_type)
    sect_phase_max = sect_fl_bound(phase_max_all, bound)

    app.PrintPlain(f'Getting {feeder.loc_name} minimum fault levels...')
    bound = 'Min'
    f_type = 'Ground'
    analysis.short_circuit(app, bound, f_type)
    pg_min_first_pass = terminal_fls(devices_sections, f_type)
    pg_min_all = append_floating_terms(app, pg_min_first_pass, floating_terms, bound, f_type)
    sect_pg_min = sect_fl_bound(pg_min_all, bound)

    f_type = '2 Phase'
    analysis.short_circuit(app, bound, f_type)
    phase_min_first_pass = terminal_fls(devices_sections, f_type)
    phase_min_all = append_floating_terms(app, phase_min_first_pass, floating_terms, bound, f_type)
    sect_phase_min = sect_fl_bound(phase_min_all, bound)

    # Package study results
    study_results = [sect_phase_max, sect_pg_max, sect_phase_min, sect_pg_min, sect_tr_phase_max, sect_tr_pg_max,
                     device_max_load, us_devices, ds_devices, ds_capacity, device_lines]

    # Package detailed fl data
    detailed_fls = [pg_max_all, phase_max_all, pg_min_all, phase_min_all, section_loads]

    # Obtain line results for conductor damage studies
    line_fls = lines_results.regional_lines(app, device_lines, phase_max_all, pg_max_all, phase_min_all, pg_min_all)

    return study_results, detailed_fls, line_fls


def get_downstream_objects(app, site_name_map) \
        -> tuple[dict[pft.ElmTerm:pft.ElmTerm], dict[pft.ElmTerm:pft.ElmTr2], dict[pft.ElmTerm:pft.ElmLne]]:
    """

    :param app:
    :param devices:
    :return:
    """

    all_grids = app.GetCalcRelevantObjects('*.ElmXnet')
    grids = [grid for grid in all_grids if grid.outserv == 0]

    devices_terminals = {}
    devices_loads = {}
    devices_lines = {}
    for dictionary in site_name_map.values():
        for key, value in dictionary.items():
            cubicle = key
            termination = value
        devices_terminals[termination] = [termination]
        devices_loads[termination] = []
        devices_lines[termination] = []
        # Do a topological search of the device downstream ojects
        down_devices = cubicle.GetAll(1, 0)
        # If the external grid is in the downstream list, you're searching in the wrong direction
        if any(item in grids for item in down_devices):
            ds_objs_list = cubicle.GetAll(0, 0)
        else:
            ds_objs_list = down_devices
        for down_object in ds_objs_list:
            if down_object.GetClassName() == "ElmTerm" and down_object.uknom > 1:
                devices_terminals[termination].append(down_object)
            if down_object.GetClassName() == "ElmTr2":
                devices_loads[termination].append(down_object)
            if down_object.GetClassName() == "ElmLne":
                devices_lines[termination].append(down_object)

    return devices_terminals, devices_loads, devices_lines


def us_ds_device(devices_terminals: dict[pft.ElmTerm:pft.ElmTerm]) \
        -> tuple[dict[pft.ElmTerm:pft.ElmTerm], dict[pft.ElmTerm:pft.ElmTerm]]:
    """
    Update all devices with the lists of downstream devices and upstreams devices
    :param devices_terminals:
    :param all_devices:
    :return:
    """
    us_devices = {device: [] for device in devices_terminals.keys()}
    ds_devices = {device: [] for device in devices_terminals.keys()}
    for device, terms in devices_terminals.items():
        # get a dictionary of device-terms that include the device in its list of terminals
        d_t_dic = {}
        for other_device, other_terms in devices_terminals.items():
            if other_device == device:
                continue
            if device in other_terms:
                d_t_dic[other_device] = other_terms
        # From this dictionary, the device with the shortest list of terminals is the backup device
        if d_t_dic:
            min_val = min([len(value) for key, value in d_t_dic.items()])
            bu_device = False
            for other_device, other_terms in d_t_dic.items():
                if len(other_terms) == min_val:
                    bu_device = other_device
            if bu_device:
                if bu_device not in us_devices[device]:
                    us_devices[device].append(bu_device)
                if device not in ds_devices[device]:
                    ds_devices[bu_device].append(device)

    return us_devices, ds_devices


def get_ds_capacity(devices_loads: dict[pft.ElmTerm:pft.ElmTr2]) -> dict[pft.ElmTerm:float]:
    """
    Calculate the capacity of all distribution transformers downstream of each device.
    """

    ds_capacity = {}
    for device, loads in devices_loads.items():
        load_kva = {load: round(load.Snom_a * 1000) for load in loads}
        ds_capacity[device] = sum(load_kva.values())
    return ds_capacity


def get_device_sections(devices_terms: dict[pft.ElmTerm:list[object]]) -> dict[pft.ElmTerm:list[object]]:
    """
    For a dictionary of device: [terminals],determine the device sections
    :param devices_terminals:
    :return:
    """

    # Sort the keys by the length of their lists in descending order
    sorted_keys = sorted(devices_terms, key=lambda k: len(devices_terms[k]), reverse=True)

    # Iterate over the sorted keys
    for i, key1 in enumerate(sorted_keys):
        for key2 in sorted_keys[i + 1:]:
            set1 = set(devices_terms[key1])
            set2 = set(devices_terms[key2])

            # Find common elements except the key of the shorter list (key2)
            common_elements = set1 & set2 - {key2}

            # Remove common elements from the longer list (key1's list)
            devices_terms[key1] = [elem for elem in devices_terms[key1] if elem not in common_elements]

    return devices_terms


def get_section_max_tr(section_loads: dict[pft.ElmTerm:pft.ElmTr2]) -> (
        tuple)[dict[pft.ElmTerm:float], dict[pft.ElmTerm:pft.ElmTerm]]:
    """

    :param app:
    :param section_loads:
    :return:
    device_max_load = {'device': str, ...}
    device_max_trs = {device.switch: [term1, term2...], ..}
    """

    device_max_load = {}
    device_max_trs = {}
    for device, loads in section_loads.items():
        load_values = {load: round(load.Snom_a * 1000) for load in loads}
        if load_values.values():
            max_load_value = max(load_values.values())
            device_max_load[device] = max_load_value
        else:
            device_max_load[device] = 0
        max_loads = [load for load in load_values if load_values[load] == max_load_value]
        # If there are multiple max loads, need to return all of them, so we can find the max load with highest fl.
        max_load_terms = [load.bushv.cterm for load in max_loads]
        device_max_trs[device] = max_load_terms

    return device_max_load, device_max_trs


def terminal_fls(devices_sections: dict[pft.ElmTerm:pft.ElmTerm], f_type: str) -> dict[pft.ElmTerm:dict[pft.ElmTerm:float]]:
    """

    :param devices_sections:
    :param f_type: 'Ground', '2 Phase', '3 Phase'
    :return:
    """

    if f_type == "Ground":
        attribute = 'm:Ikss:A'
    elif f_type == '2 Phase':
        attribute = 'm:Ikss:B'
    else:
        # f_type == '3 Phase'
        attribute = 'm:Ikss'

    results_all = {}
    for device, terminals in devices_sections.items():
        results_all[device] = {}
        for terminal in terminals:
            if terminal.HasAttribute(attribute):
                results_all[device][terminal] = round(terminal.GetAttribute(attribute), 3) * 1000
            else:
                results_all[device][terminal] = 0

    return results_all


def sect_fl_bound(results_all: dict[pft.ElmTerm:dict[pft.ElmTerm:float]], bound: str) -> dict[pft.ElmTerm:float]:
    """

    :param results_all:
    :param bound: 'Min', 'Max'.
    :return:
    """

    sect_bound = {}
    for device, terms in results_all.items():
        sect_bound[device] = {}
        non_zero_terms = {term: fl for term, fl in terms.items() if fl != 0}
        if not non_zero_terms:
            sect_bound[device] = 'no terminations'
            continue
        if bound == 'Min':
            min_term = min(non_zero_terms, key=non_zero_terms.get)
            sect_bound[device] = {min_term: non_zero_terms[min_term]}
        elif bound == 'Max':
            max_term = max(non_zero_terms, key=non_zero_terms.get)
            sect_bound[device] = {max_term: non_zero_terms[max_term]}

    return sect_bound


def append_floating_terms(app, results_all: dict[pft.ElmTerm:dict[pft.ElmTerm:float]],
                          floating_terms: dict[pft.ElmTerm:dict[pft.ElmLne:float]], bound: str, f_type: str) \
        -> dict[pft.ElmTerm:dict[pft.ElmTerm:float]]:
    """

    :param app:
    :param results_all:
    :param floating_terms:
    :param bound: 'Max', 'Min'
    :param f_type: 'Ground', '2 Phase', '3 Phase'
    :return:
    """

    for device, lines in floating_terms.items():
        for line, term in lines.items():
            if line.bus1.cterm == term:
                ppro = 1
            else:
                ppro = 99
            analysis.short_circuit(app, bound, f_type, location=line, ppro=ppro)
            line_current = analysis.get_line_current(line)
            results_all[device].update({term: line_current})

    return results_all