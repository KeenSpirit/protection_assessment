import sys
sys.path.append(r"\\Ecasd01\WksMgmt\PowerFactory\ScriptsDEV\PowerFactoryTyping")
import powerfactorytyping as pft
from devices import devices as ds
from fault_study import lines_results, analysis, fault_impedance, floating_terminals as ft
from logging_config.configure_logging import log_arguments

from importlib import reload
reload(ft)
reload(fault_impedance)

def fault_study(app, region, feeder, bu_devices, devices):
    """

    :param app:
    :param feeder:
    :param sites:
    :return:
    """

    app.PrintPlain("Performing fault level study...")

    get_downstream_objects(app, devices)
    us_ds_device(devices, bu_devices)
    get_ds_capacity(devices)
    get_device_sections(devices)

    analysis.short_circuit(app, bound='Max', f_type='Ground')
    terminal_fls(devices, bound='Max', f_type='Ground')
    analysis.short_circuit(app, bound='Max', f_type='3 Phase')
    terminal_fls(devices, bound='Max', f_type='3 Phase')
    analysis.short_circuit(app, bound='Max', f_type='2 Phase')
    terminal_fls(devices, bound='Max', f_type='2 Phase')
    analysis.short_circuit(app, bound='Min', f_type='Ground')
    terminal_fls(devices, bound='Min', f_type='Ground')
    analysis.short_circuit(app, bound='Min', f_type='2 Phase')
    terminal_fls(devices, bound='Min', f_type='2 Phase')
    analysis.short_circuit(app, bound='Min', f_type='Ground Z10')
    terminal_fls(devices, bound='Min', f_type='Ground Z10')
    analysis.short_circuit(app, bound='Min', f_type='Ground Z50')
    terminal_fls(devices, bound='Min', f_type='Ground Z50')
    fault_impedance.update_node_construction(devices)

    floating_terms = ft.get_floating_terminals(feeder, devices)
    append_floating_terms(app, devices, floating_terms)
    update_device_data(app, region, devices)
    update_line_data(devices)

    return devices


def obtain_region(app):

    project = app.GetActiveProject()
    derived_proj = project.der_baseproject
    der_proj_name = derived_proj.GetFullName()

    regional_model = 'Regional Models'
    seq_model = 'SEQ'

    if regional_model in der_proj_name:
        # This is a regional model
        region=regional_model
    elif seq_model in der_proj_name:
        # This is a SEQ model
        region = seq_model
    else:
        msg = (
            "The appropriate region for the model could not be found. "
            "Please contact the script administrator to resolve this issue."
        )
        raise RuntimeError(msg)
    return region


def get_downstream_objects(app, devices):
    """

    :param devices:
    :return:
    """
    all_grids = app.GetCalcRelevantObjects('*.ElmXnet')
    grids = [grid for grid in all_grids if grid.outserv == 0]

    region = obtain_region(app)
    for device in devices:
        terminals = [device.term]
        loads = []
        lines = []
        down_devices = device.cubicle.GetAll(1, 0)
        # If the external grid is in the downstream list, you're searching in the wrong direction
        if any(item in grids for item in down_devices):
            down_objs = device.cubicle.GetAll(0, 0)
        else:
            down_objs = down_devices
        for obj in down_objs:
            if obj.GetClassName() == "ElmTerm" and obj.uknom > 1:
                terminals.append(obj)
            if obj.GetClassName() == "ElmLod" and region == 'SEQ':
                loads.append(obj)
            if obj.GetClassName() == "ElmTr2" and region == 'Regional Models':
                load_type = obj.typ_id
                if "Regulators" not in load_type.GetFullName():
                    loads.append(obj)
            if obj.GetClassName() == "ElmLne":
                lines.append(obj)
        device.sect_terms = terminals
        device.sect_loads = loads
        device.sect_lines = lines


def us_ds_device(devices, bu_devices):
    """

    :param devices:
    :return:
    """

    for device in devices:
        us_devices = []
        for other_device in devices:
            if other_device == device:
                continue
            if device.term in other_device.sect_terms:
                us_devices.append(other_device)
        if us_devices:
            bu_device = min(us_devices, key=lambda item: len(item.sect_terms))
            device.us_devices.append(bu_device)
            bu_device.ds_devices.append(device)
        if not device.us_devices:
            connected_elements = device.cubicle.GetAll(1, 0) + device.cubicle.GetAll(0, 0)
            for grid, grid_devices in bu_devices.items():
                if grid in connected_elements:
                    device.us_devices.extend(grid_devices)


def get_ds_capacity(devices):
    """
    Calculate the capacity of all distribution transformers downstream of each device.
    """
    def _get_load(obj):
        return obj.Strat if obj.GetClassName() == "ElmLod" else obj.Snom_a * 1000

    for device in devices:
        device.ds_capacity = round(sum([_get_load(obj) for obj in device.sect_loads]))


def get_device_sections(devices):
    """

    :param devices:
    :return:
    """

    def _sections(devices_objs):
        # Sort the keys by the length of their lists in descending order
        sorted_keys = sorted(devices_objs, key=lambda k: len(devices_objs[k]), reverse=True)
        # Iterate over the sorted keys
        for i, key1 in enumerate(sorted_keys):
            for key2 in sorted_keys[i + 1:]:
                set1 = set(devices_objs[key1])
                set2 = set(devices_objs[key2])
                # Find common elements except the key of the shorter list (key2)
                common_elements = set1 & set2 - {key2}
                # Remove common elements from the longer list (key1's list)
                devices_objs[key1] = [elem for elem in devices_objs[key1] if elem not in common_elements]
        return devices_objs

    devices_terms = {device.term:device.sect_terms for device in devices}
    devices_loads = {device.term:device.sect_loads for device in devices}
    devices_lines = {device.term:device.sect_lines for device in devices}

    for device in devices:
        section_terms = _sections(devices_terms)[device.term]
        dataclass_terms = [ds.Termination(
            obj, None, ph_attr_lookup(obj.phtech), round(obj.uknom,2), None, None, None, None, None, None
        ) for obj in section_terms]
        device.sect_terms = dataclass_terms
        section_loads = _sections(devices_loads)[device.term]
        dataclass_loads = [dataclass_load(obj) for obj in section_loads]
        device.sect_loads = dataclass_loads

        section_lines = _sections(devices_lines)[device.term]
        dataclass_lines = []
        for elmlne in section_lines:
            line_type, line_therm_rating = lines_results.get_conductor(elmlne)
            dataclass_lines.append(
                ds.Line(
                    elmlne, None, None, None, None, line_type, line_therm_rating, None, None, None, None)
            )
        device.sect_lines = dataclass_lines



def terminal_fls(devices, bound, f_type):
    """

    :param devices:
    :param bound:
    :param f_type:
    :return:
    """

    def _check_att(obj, attribute):
        if obj.HasAttribute(attribute):
            terminal_fl = round(obj.GetAttribute(attribute) * 1000)
        else:
            terminal_fl = 0
        return terminal_fl

    for device in devices:
        for terminal in device.sect_terms:
            obj = terminal.object
            Ia = _check_att(obj, 'm:Ikss:A')
            Ib = _check_att(obj, 'm:Ikss:B')
            Ic = _check_att(obj, 'm:Ikss:C')

            if bound == 'Max':
                if f_type == 'Ground':
                    terminal.max_fl_pg = max(Ia, Ib, Ic)
                elif terminal.max_fl_ph:
                    terminal.max_fl_ph = max(terminal.max_fl_ph, max(Ia, Ib, Ic))
                else:
                    terminal.max_fl_ph = max(Ia, Ib, Ic)
            elif f_type == 'Ground':
                terminal.min_fl_pg = max(Ia, Ib, Ic)
            elif f_type == 'Ground Z10':
                terminal.min_fl_pg10 = max(Ia, Ib, Ic)
            elif f_type == 'Ground Z50':
                terminal.min_fl_pg50 = max(Ia, Ib, Ic)
            else:
                terminal.min_fl_ph = max(Ia, Ib, Ic)


def append_floating_terms(app, devices, floating_terms):
    """

    :param app:
    :param devices:
    :param floating_terms:
    :return:
    """

    for dev, lines in floating_terms.items():
        for line, term in lines.items():
            if line.bus1.cterm == term:
                ppro = 1
            else:
                ppro = 99
            termination = ds.Termination(
                term, None, ph_attr_lookup(term.phtech), round(term.uknom,2), None, None, None, None, None, None
            )
            analysis.short_circuit(app, bound='Max', f_type='3 Phase', location=line, ppro=ppro)
            current = analysis.get_line_current(line)
            termination.max_fl_ph = current
            analysis.short_circuit(app, bound='Max', f_type='Ground', location=line, ppro=ppro)
            current = analysis.get_line_current(line)
            termination.max_fl_pg = current
            analysis.short_circuit(app, bound='Min', f_type='2 Phase', location=line, ppro=ppro)
            current = analysis.get_line_current(line)
            termination.min_fl_ph = current
            analysis.short_circuit(app, bound='Min', f_type='Ground', location=line, ppro=ppro)
            current = analysis.get_line_current(line)
            termination.min_fl_pg = current
            analysis.short_circuit(app, bound='Min', f_type='Ground Z10', location=line, ppro=ppro)
            current = analysis.get_line_current(line)
            termination.min_fl_pg10 = current
            analysis.short_circuit(app, bound='Min', f_type='Ground Z50', location=line, ppro=ppro)
            current = analysis.get_line_current(line)
            termination.min_fl_pg50 = current

            sect_terms = [device.sect_terms for device in devices if device.object == dev]
            sect_terms.append(termination)


def update_device_data(app, region, devices):
    """

    :param devices:
    :return:
    """

    def _safe_max(sequence):
        try:
            return max(sequence)
        except ValueError:
            return 0

    def _safe_min(sequence):
        try:
            return min(sequence)
        except ValueError:
            return 0

    for device in devices:
        # Update transformer data
        try:
            device.max_tr_size = _safe_max([load.load_kva for load in device.sect_loads])
        except:
            app.PrintPlain(f"device.object{device.object}")
            app.PrintPlain(f"device.sect_loads{device.sect_loads}")
        max_ds_trs = [load.term for load in device.sect_loads if load.load_kva == device.max_tr_size]
        max_ds_trs_class = [term for term in device.sect_terms if term.object in max_ds_trs]
        try:
            device.tr_max_ph = _safe_max([term.max_fl_ph for term in device.sect_terms if term in max_ds_trs_class])
        except:
            app.PrintPlain(f"device.object{device.object}")
            app.PrintPlain(f"max_ds_trs{max_ds_trs}")
            app.PrintPlain(f"device.sect_terms{device.sect_terms}")
            app.PrintPlain(f"max_ds_trs_class{max_ds_trs_class}")
        device.tr_max_pg = _safe_max([term.max_fl_pg for term in device.sect_terms if term in max_ds_trs_class])
        device.max_ds_tr = \
            [term.object.cpSubstat.loc_name for term in max_ds_trs_class if term.max_fl_pg == device.tr_max_pg][0]
        # Update device fl data
        device.max_fl_ph = _safe_max([term.max_fl_ph for term in device.sect_terms])
        device.max_fl_pg = _safe_max([term.max_fl_pg for term in device.sect_terms])
        device.min_fl_ph = _safe_min([term.min_fl_ph for term in device.sect_terms if term.min_fl_ph > 0])
        device.min_fl_pg = (
            _safe_min([fault_impedance.term_pg_fl(region, term) for term in device.sect_terms if term.min_fl_pg > 0]))


def update_line_data(devices):
    """

    :param devices:
    :return:
    """

    for device in devices:
        lines = device.sect_lines
        for line in lines:
            elmlne = line.object
            lne_cubs = [elmlne.bus1, elmlne.bus2]
            lne_terms = [cub.cterm for cub in lne_cubs if cub is not None]
            sect_term_obs = [term.object for term in device.sect_terms]
            if any(terms in sect_term_obs for terms in lne_terms):
                line.max_fl_ph = max([term.max_fl_ph for term in device.sect_terms if term.object in lne_terms])
                line.max_fl_pg = max([term.max_fl_pg for term in device.sect_terms if term.object in lne_terms])
                line.min_fl_ph = min([term.min_fl_ph for term in device.sect_terms if term.object in lne_terms])
                line.min_fl_pg = min([term.min_fl_pg for term in device.sect_terms if term.object in lne_terms])
            else:
                line.max_fl_ph = 0
                line.max_fl_pg = 0
                line.min_fl_ph = 0
                line.min_fl_pg = 0


def dataclass_load(load):
    if load.GetClassName() == "ElmLod":
        return ds.Load(load, load.bus1.cterm, load.Strat)
    if load.GetClassName() == "ElmTr2":
        return ds.Load(load, load.bushv.cterm, round(load.Snom_a * 1000))


def ph_attr_lookup(attr):
    """
    Convert the terminal phase technology attribute phtech to a meaningful value
    :param attr:
    :return:
    """
    phases = {1:[6, 7, 8], 2:[2, 3, 4, 5], 3:[0, 1]}
    for phase, attr_list in phases.items():
        if attr in attr_list:
            return phase


def print_devices(app, devices):
    """ Function used for debugging purposes only"""

    for device in devices:
        app.PrintPlain(f"object: {device.object}")
        app.PrintPlain(f"cubicle: {device.cubicle}")
        app.PrintPlain(f"term: {device.term}")
        app.PrintPlain(f"ds_capacity: {device.ds_capacity}")
        app.PrintPlain(f"max_fl_ph: {device.max_fl_ph}")
        app.PrintPlain(f"max_fl_pg: {device.max_fl_pg}")
        app.PrintPlain(f"min_fl_ph: {device.min_fl_ph}")
        app.PrintPlain(f"min_fl_pg: {device.min_fl_pg}")
        app.PrintPlain(f"max_ds_tr: {device.max_ds_tr}")
        app.PrintPlain(f"max_tr_size: {device.max_tr_size}")
        app.PrintPlain(f"tr_max_ph: {device.tr_max_ph}")
        app.PrintPlain(f"tr_max_pg: {device.tr_max_pg}")
        app.PrintPlain(f"sect_terms: {device.sect_terms}")
        app.PrintPlain(f"sect_loads: {device.sect_loads}")
        app.PrintPlain(f"sect_lines: {device.sect_lines}")
        app.PrintPlain(f"us_devices: {device.us_devices}")
        app.PrintPlain(f"ds_devices: {device.ds_devices}")

