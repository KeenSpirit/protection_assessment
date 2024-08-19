import sys
sys.path.append(r"\\Ecasd01\WksMgmt\PowerFactory\ScriptsDEV\PowerFactoryTyping")
import powerfactorytyping as pft
from dataclasses import dataclass
from fault_study import lines_results, analysis
from importlib import reload
from logging_config.configure_logging import log_arguments


def fault_study(app, feeder, sites) -> tuple[list, list, list]:
    """

    :param app:
    :param feeder:
    :param sites:
    :return:
    """

    devices = [
        Device(site,site.fold_id,site.fold_id.cterm,None,None,None,None,None,None,None,None,None,[],[],[],[],[]) for site in sites
    ]

    get_downstream_objects(devices)
    us_ds_device(devices)
    get_ds_capacity(devices)
    get_device_sections(devices)

    analysis.short_circuit(app, bound='Max', f_type='Ground')
    terminal_fls(devices, bound='Max', f_type='Ground')
    analysis.short_circuit(app, bound='Max', f_type='3 Phase')
    terminal_fls(devices, bound='Max', f_type='3 Phase')
    analysis.short_circuit(app, bound='Min', f_type='Ground')
    terminal_fls(devices, bound='Min', f_type='Ground')
    analysis.short_circuit(app, bound='Min', f_type='2 Phase')
    terminal_fls(devices, bound='Min', f_type='2 Phase')

    floating_terms = get_floating_terminals(feeder, devices)
    append_floating_terms(app, devices, floating_terms)
    update_device_data(devices)
    update_line_data(devices)


@log_arguments
def get_downstream_objects(devices):
    """

    :param app:
    :param devices:
    :return:
    """

    for device in devices:
        terminals = [device.term]
        loads = []
        lines = []
        down_objs = device.cubicle.GetAll()
        for obj in down_objs:
            if obj.GetClassName() == "ElmTerm" and obj.uknom > 1:
                terminals.append(obj)
            if obj.GetClassName() == "ElmTr2":
                loads.append(obj)
            if obj.GetClassName() == "ElmLne":
                lines.append(obj)
        device.sect_terms = terminals
        device.sect_loads = loads
        device.sect_lines = lines


def us_ds_device(devices):
    """
    Update all devices with the lists of downstream devices and upstreams devices
    :param devices_terminals:
    :param all_devices:
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


def get_ds_capacity(devices):
    """
    Calculate the capacity of all distribution transformers downstream of each device.
    """
    for device in devices:
        device.ds_capacity = sum([load.Snom_a * 1000 for load in device.sect_loads])


def get_device_sections(devices) -> dict[pft.ElmTerm:list[object]]:
    """
    For a dictionary of device: [terminals],determine the device sections
    :param devices_terminals:
    :return:
    """

    devices_terms = {device.term:device.sect_terms for device in devices}
    devices_loads = {device.term:device.sect_loads for device in devices}
    devices_lines = {device.term:device.sect_lines for device in devices}

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

    for device in devices:
        section_terms = _sections(devices_terms)[device.term]
        dataclass_terms = [Termination(obj, None, None, None, None) for obj in section_terms]
        device.sect_terms = dataclass_terms

        section_loads = _sections(devices_loads)[device.term]
        dataclass_loads = [Load(obj, obj.bushv.cterm, obj.Snom_a * 1000) for obj in section_loads]
        device.sect_loads = dataclass_loads

        section_lines = _sections(devices_lines)[device.term]
        dataclass_lines = []
        for elmlne in section_lines:
            line_type, line_therm_rating = lines_results.get_conductor(elmlne)
            dataclass_lines.append(
                Line(elmlne, None, None, None, None, line_type, line_therm_rating, None, None, None, None)
            )
        device.sect_lines = dataclass_lines


def find_end_points(feeder: object) -> list[object]:
    """
    Returns a list of sections with only one connection (i.e. end points).
    :param feeder: The feeder being investigated.
    :return: A list of line sections that have only one connection (i.e. end points).
    """

    floating_lines = []

    # Get all the sections that make up the selected feeder.
    feeder_lines = feeder.GetObjs('ElmLne')

    for ElmLne in feeder_lines:
        if ElmLne.bus1:
            bus1 = [x.obj_id for x in ElmLne.bus1.cterm.GetConnectedCubicles()
                    if x is not ElmLne.bus1
                    if x.obj_id.GetClassName() == 'ElmLne']
        else:
            bus1 = []

        if ElmLne.bus2:
            if ElmLne.bus2.HasAttribute('cterm'):
                bus2 = [x.obj_id for x in ElmLne.bus2.cterm.GetConnectedCubicles()
                        if x is not ElmLne.bus2
                        if x.obj_id.GetClassName() == 'ElmLne']
            else:
                bus2 = []
        else:
            bus2 = []

        if len(bus1) == 1 or len(bus2) == 1 \
                or (len(bus1) > 1 and ElmLne not in bus1) \
                or (len(bus2) > 1 and ElmLne not in bus2):
            floating_lines.append(ElmLne)

    return floating_lines


def get_floating_terminals(feeder, devices) -> dict[object:dict[object:object]]:
    """
    Outputs all floating terminal objects with their associated line objects for all devices
    :param feeder:
    :param devices_section:
    :return:
    """

    floating_terms= {}
    floating_lines = find_end_points(feeder)
    for device in devices:
        terms = [term.object for term in device.sect_terms]
        floating_terms[device.object ] = {}
        for line in floating_lines:
            t1, t2 = line.GetConnectedElements()
            t3 = line.GetConnectedElements(1,1,0)
            if len(t3) == 1 and t3[0] == t2 and t2 in terms and t1 not in terms:
                floating_terms[device.object ][line] = Termination(t1, None, None, None, None)
            elif len(t3) == 1 and t3[0] == t1 and t1 in terms and t2 not in terms:
                floating_terms[device.object ][line] = Termination(t2, None, None, None, None)

    return floating_terms


def terminal_fls(devices, bound, f_type) -> dict[pft.ElmTerm:dict[pft.ElmTerm:float]]:
    """

    :param devices_sections:
    :param f_type: 'Ground', 'Phase'
    :return:
    """

    def _check_att(obj, attribute):
        if obj.HasAttribute(attribute):
            terminal_fl = round(obj.GetAttribute(attribute), 3) * 1000
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
                else:
                    terminal.max_fl_ph = max(Ia, Ib, Ic)
            elif f_type == 'Ground':
                terminal.min_fl_pg = max(Ia, Ib, Ic)
            else:
                terminal.min_fl_ph = max(Ia, Ib, Ic)


def append_floating_terms(app, devices, floating_terms):
    """

    :param app:
    :param results_all:
    :param floating_terms:
    :param bound: 'Max', 'Min'
    :param f_type: 'Ground', '2 Phase', '3 Phase'
    :return:
    """

    for dev, lines in floating_terms.items():
        for line, term in lines.items():
            if line.bus1.cterm == term:
                ppro = 1
            else:
                ppro = 99
            termination = Termination(term, None, None, None, None)
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

            sect_terms = [device.sect_terms for device in devices if device.object == dev]
            sect_terms.append(termination)


def update_device_data(devices):

    for device in devices:
        # Update transformer data
        device.max_tr_size = max([load.load_kva for load in device.sect_loads])
        max_ds_trs = [load.term for load in device.sect_loads if load.load_kva == device.max_tr_size]
        max_ds_trs_class = [termination for termination in device.sect_terms if termination.object in max_ds_trs]
        device.tr_max_ph = max([termination.max_fl_ph for termination in device.sect_terms if termination in max_ds_trs_class])
        device.tr_max_pg = max([termination.max_fl_pg for termination in device.sect_terms if termination in max_ds_trs_class])
        device.max_ds_tr = [termination.object for termination in max_ds_trs_class if termination.max_fl_pg == device.tr_max_pg][0]
        # Update device fl data
        device.max_fl_ph = max([termination.max_fl_ph for termination in device.sect_terms])
        device.max_fl_pg = max([termination.max_fl_pg for termination in device.sect_terms])
        device.min_fl_ph = min([termination.max_fl_ph for termination in device.sect_terms])
        device.min_fl_pg = min([termination.max_fl_pg for termination in device.sect_terms])


def update_line_data(devices):


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
                line.min_fl_ph = min([term.min_fl_pg for term in device.sect_terms if term.object in lne_terms])
            else:
                line.max_fl_ph = 0
                line.max_fl_pg = 0
                line.min_fl_ph = 0
                line.min_fl_ph = 0


def print_devices(app, devices):

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


@dataclass
class Device:
    object: object
    cubicle: object
    term: object
    ds_capacity: float
    max_fl_ph: float
    max_fl_pg: float
    min_fl_ph: float
    min_fl_pg: float
    max_ds_tr: str
    max_tr_size: int
    tr_max_ph: float
    tr_max_pg: float
    sect_terms: list
    sect_loads: list
    sect_lines: list
    us_devices: list
    ds_devices: list


@dataclass
class Line:
    object: object
    max_fl_ph: float
    max_fl_pg: float
    min_fl_ph: float
    min_fl_pg: float
    line_type: str
    thermal_rating: float
    ph_clear_time: float
    ph_fl: float
    pg_clear_time: float
    pg_fl: float


@dataclass
class Termination:
    object: object
    max_fl_ph: float
    max_fl_pg: float
    min_fl_ph: float
    min_fl_pg: float


@dataclass
class Load:
    object: object
    term: object
    load_kva: float

