"""
Performs comprehensive fault level analysis for PowerFactory distribution networks.

Main workflow:
1. Build device hierarchy and downstream objects
2. Run short-circuit stuides (max/min, 3ph/2ph, ground)
3. Extract fault currents for terminals and lines
4. Handle floating terminals at network endpoints
5. Update device protection settings

"""

import sys

from typing import List, Dict, Union
from pf_config import pft
import pf_protection_helper as helper
from fault_study import analysis, fault_impedance, floating_terminals as ft
import script_classes as dd
from importlib import reload

reload(analysis)
reload(ft)
reload(dd)
reload(fault_impedance)


def fault_study(app: pft.Application, external_grid: Dict, region: str, feeder: dd.Feeder, study_selections: List[str]):
    """

    :param app:
    :param external_grid:
    :param region:
    :param feeder:
    :return:
    """

    app.PrintPlain(f"Performing fault level study for {feeder.obj.loc_name}...")
    get_downstream_objects(app, feeder.devices)
    us_ds_device(feeder.devices, feeder.bu_devices)
    get_ds_capacity(feeder.devices)
    get_device_sections(feeder.devices)

    study_configs = [
        ('Max', 'Ground'), ('Max', '3-Phase'), ('Max', '2-Phase'),
        ('Min', 'Ground'), ('Min', '3-Phase'), ('Min', '2-Phase'),
        ('Min', 'Ground Z10'), ('Min', 'Ground Z50'),
    ]
    sn_study_configs = [
        ('Min', '2-Phase'), ('Min', 'Ground'), ('Min', 'Ground Z10'),
        ('Min', 'Ground Z50'),
    ]

    if "Fault Level Study (all relays configured in model)" in study_selections:
        consider_prot = 'All'
    else:
        consider_prot = 'None'
    for bound, fault_type in study_configs:
        analysis.short_circuit(app, bound, fault_type, consider_prot)
        terminal_fls(feeder.devices, bound=bound, f_type=fault_type)
    if grid_equivalance_check(external_grid):
        copy_min_fls(feeder.devices)
    else:
        reset_min_source_imp(external_grid, sys_norm_min=True)
        for bound, fault_type in sn_study_configs:
            analysis.short_circuit(app, bound, fault_type, consider_prot)
            terminal_fls(feeder.devices, bound='SN_Min', f_type=fault_type)
        reset_min_source_imp(external_grid, sys_norm_min=False)

    fault_impedance.update_node_construction(feeder.devices)

    floating_terms = ft.get_floating_terminals(feeder.obj, feeder.devices)
    append_floating_terms(app, external_grid, feeder.devices, floating_terms, consider_prot)
    update_device_data(region, feeder.devices)
    update_line_data(app, region, feeder.devices)
    for device in feeder.devices:
        for terminal in device.sect_terms:
            dd.populate_fault_currents(terminal)


def get_downstream_objects(app: pft.Application, devices: List):
    """
    Populates device objects with their downstream network components.

    Traverses the network topology from each device to identify:
    - Terminals (elmterm) with voltage > 1kV
    - Distribution loads (elmlod) in SEQ region
    - Transformers (ElmTr2) in Regional models (excluding regulators)
    - Line segments (elmline)

    Args:
        app: PowerFactory application instance
        devices: List of Device dataclass objects to populate

    Side Effects:
        Updates section_terms, section_loads, and sect_lines for each device
    """

    all_grids = app.GetCalcRelevantObjects('*.ElmXnet')
    grids = [grid for grid in all_grids if grid.outserv == 0]

    region = helper.obtain_region(app)
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
            if obj.GetClassName() == dd.ElementType.TERM.value and obj.uknom > 1:
                terminals.append(obj)
            if obj.GetClassName() == dd.ElementType.LOAD.value and region == 'SEQ':
                loads.append(obj)
            if obj.GetClassName() == dd.ElementType.TFMR.value and region == 'Regional Models':
                load_type = obj.typ_id
                if "Regulators" not in load_type.GetFullName():
                    loads.append(obj)
            if obj.GetClassName() == dd.ElementType.LINE.value:
                lines.append(obj)
        device.sect_terms = terminals
        device.sect_loads = loads
        device.sect_lines = lines


def us_ds_device(devices: List[dd.Device], bu_devices: Dict):
    """

    :param devices:
    :param bu_devices:
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
            if bu_devices:
                for grid, grid_devices in bu_devices.items():
                    if grid in connected_elements:
                        device.us_devices.extend(grid_devices)


def get_ds_capacity(devices: List[dd.Device]):
    """
    Calculate the capacity of all distribution transformers downstream of each device.
    """
    def _get_load(obj: pft.ElmLod) -> float:
        return obj.Strat if obj.GetClassName() == dd.ElementType.LOAD.value else obj.Snom_a * 1000

    for device in devices:
        device.ds_capacity = round(sum([_get_load(obj) for obj in device.sect_loads]))


def get_device_sections(devices: List[dd.Device]):
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
        dataclass_terms = [dd.initialise_term_dataclass(elmterm) for elmterm in section_terms]

        device.sect_terms = dataclass_terms
        section_loads = _sections(devices_loads)[device.term]
        dataclass_loads = [dd.initialise_load_dataclass(elmlod) for elmlod in section_loads]
        device.sect_loads = dataclass_loads

        section_lines = _sections(devices_lines)[device.term]
        dataclass_lines = []
        for elmlne in section_lines:
            dataclass_lines.append(dd.initialise_line_dataclass(elmlne))
        device.sect_lines = dataclass_lines


def terminal_fls(devices: List[dd.Device], bound: str, f_type: str):
    """

    :param devices:
    :param bound:
    :param f_type:
    :return:
    """

    for device in devices:
        for terminal in device.sect_terms:
            elmterm = terminal.obj
            if bound == 'Max':
                if f_type == 'Ground':
                    terminal.max_fl_pg = analysis.get_terminal_current(elmterm)
                elif f_type == '3-Phase':
                    terminal.max_fl_3ph = analysis.get_terminal_current(elmterm)
                else:
                    terminal.max_fl_2ph = analysis.get_terminal_current(elmterm)
            elif bound == 'Min':
                if f_type == 'Ground':
                    terminal.min_fl_pg = analysis.get_terminal_current(elmterm)
                elif f_type == 'Ground Z10':
                    terminal.min_fl_pg10 = analysis.get_terminal_current(elmterm)
                elif f_type == 'Ground Z50':
                    terminal.min_fl_pg50 = analysis.get_terminal_current(elmterm)
                elif f_type == '3-Phase':
                    terminal.min_fl_3ph = analysis.get_terminal_current(elmterm)
                else:
                    terminal.min_fl_2ph = analysis.get_terminal_current(elmterm)
            else:
                if f_type == 'Ground':
                    terminal.min_sn_fl_pg = analysis.get_terminal_current(elmterm)
                elif f_type == 'Ground Z10':
                    terminal.min_sn_fl_pg10 = analysis.get_terminal_current(elmterm)
                elif f_type == 'Ground Z50':
                    terminal.min_sn_fl_pg50 = analysis.get_terminal_current(elmterm)
                elif f_type == '2-Phase':
                    terminal.min_sn_fl_2ph = analysis.get_terminal_current(elmterm)


def append_floating_terms(app: pft.Application, external_grid: Dict, devices: List[dd.Device], floating_terms: Dict, consider_prot: str):
    """

    :param app:
    :param external_grid:
    :param devices:
    :param floating_terms:
    :param consider_prot:
    :return:
    """

    for dev, lines in floating_terms.items():
        for line, elmterm in lines.items():
            if line.bus1.cterm == elmterm:
                ppro = 1
            else:
                ppro = 99
            termination = dd.initialise_term_dataclass(elmterm)
            study_configs = [
                ('Max', '3-Phase','max_fl_3ph'), ('Max', '2-Phase', 'max_fl_2ph'), ('Max', 'Ground', 'max_fl_pg'),
                ('Min', '3-Phase', 'min_fl_3ph'), ('Min', '2-Phase', 'min_fl_2ph'), ('Min', 'Ground', 'min_fl_pg'),
                ('Min', 'Ground Z10', 'min_fl_pg10'), ('Min', 'Ground Z50', 'min_fl_pg50'),
            ]
            for bound, fault_type, attribute in study_configs:
                analysis.short_circuit(app, bound, fault_type, consider_prot, location=line, relative=ppro)
                current = analysis.get_line_current(line)
                setattr(termination, attribute, current)

            if grid_equivalance_check:
                termination.min_sn_fl_pg = termination.min_fl_pg
                termination.min_sn_fl_pg10 = termination.min_fl_pg10
                termination.min_sn_fl_pg50 = termination.min_fl_pg50
                termination.min_sn_fl_2ph = termination.min_fl_2ph
            else:
                sn_study_configs = [
                    ('Min', '2-Phase', 'min_sn_fl_2ph'), ('Min', 'Ground', 'min_sn_fl_pg'),
                    ('Min', 'Ground Z10', 'min_sn_fl_pg10'), ('Min', 'Ground Z50', 'min_sn_fl_pg50'),
                ]
                reset_min_source_imp(external_grid, sys_norm_min=True)
                for bound, fault_type, attribute in sn_study_configs:
                    analysis.short_circuit(app, bound, fault_type, consider_prot, location=line, relative=ppro)
                    current = analysis.get_line_current(line)
                    setattr(termination, attribute, current)
                reset_min_source_imp(external_grid, sys_norm_min=False)

            sect_terms = [device.sect_terms for device in devices if device.term == dev][0]
            sect_terms.append(termination)


def grid_equivalance_check(new_grid_data: Dict) -> bool:
    identical_grids = True
    for grid, attributes in new_grid_data.items():
        for i in range (5, 10):
            if attributes[i] != attributes[i+5]:
                identical_grids = False
                break
    return identical_grids


def reset_min_source_imp(new_grid_data: Dict, sys_norm_min: bool=False):
    """"""

    for grid, attributes in new_grid_data.items():
        if sys_norm_min and attributes[10] <= 0:
            grid.SetAttribute('outserv', 1)
        elif sys_norm_min:
            grid.SetAttribute('ikssmin', attributes[10])
            grid.SetAttribute('rntxnmin', attributes[11])
            grid.SetAttribute('z2tz1min', attributes[12])
            grid.SetAttribute('x0tx1min', attributes[13])
            grid.SetAttribute('r0tx0min', attributes[14])
        else:
            grid.SetAttribute('outserv', 0)
            grid.SetAttribute('ikssmin', attributes[5])
            grid.SetAttribute('rntxnmin', attributes[6])
            grid.SetAttribute('z2tz1min', attributes[7])
            grid.SetAttribute('x0tx1min', attributes[8])
            grid.SetAttribute('r0tx0min', attributes[9])


def copy_min_fls(devices: List[dd.Device]):
    """

    :param devices:
    :param bound:
    :param f_type:
    :return:
    """
    for device in devices:
        for terminal in device.sect_terms:
            terminal.min_sn_fl_pg = terminal.min_fl_pg
            terminal.min_sn_fl_pg10 = terminal.min_fl_pg10
            terminal.min_sn_fl_pg50 = terminal.min_fl_pg50
            terminal.min_sn_fl_2ph = terminal.min_fl_2ph


def update_device_data(region: str, devices: List[dd.Device]):
    """

    :param devices:
    :return:
    """

    def _safe_max(sequence: List) -> Union[int, float]:
        try:
            return max(sequence)
        except ValueError:
            return 0

    def _safe_min(sequence: List) -> Union[int, float]:
        try:
            return min(sequence)
        except ValueError:
            return 0

    for device in devices:
        # Update transformer data
        max_tr_size = _safe_max([load.load_kva for load in device.sect_loads])
        max_ds_trs = [load for load in device.sect_loads if load.load_kva == max_tr_size]
        # The load must have a termination in the list of section terminations
        sect_terms = [term.obj for term in device.sect_terms]
        max_ds_trs = [load for load in max_ds_trs if load.term in sect_terms]
        # Initiate max values
        try:
            max_ds_tr = max_ds_trs[0]
        except IndexError:
            max_ds_tr = dd.initialise_load_dataclass(None)
        max_fl_pg = 0
        # Search for max tr
        for tr in max_ds_trs:
            term_dataclass = [t for t in device.sect_terms if t.obj == tr.term][0]
            if term_dataclass.max_fl_pg >= max_fl_pg:
                max_fl_pg = term_dataclass.max_fl_pg
                tr.term = term_dataclass.obj
                tr.max_pg = term_dataclass.max_fl_pg
                tr.max_ph = max(term_dataclass.max_fl_3ph, term_dataclass.max_fl_3ph)
                max_ds_tr = tr
        device.max_ds_tr = max_ds_tr

        # Update device fl data
        device.max_fl_3ph = _safe_max([term.max_fl_3ph for term in device.sect_terms])
        device.max_fl_2ph = _safe_max([term.max_fl_2ph for term in device.sect_terms])
        device.max_fl_pg = _safe_max([term.max_fl_pg for term in device.sect_terms])
        device.min_fl_3ph = _safe_min([term.min_fl_3ph for term in device.sect_terms if term.min_fl_3ph > 0])
        device.min_fl_2ph = _safe_min([term.min_fl_2ph for term in device.sect_terms if term.min_fl_2ph > 0])
        device.min_fl_pg = (
            _safe_min([fault_impedance.term_pg_fl(region, term) for term in device.sect_terms if term.min_fl_pg > 0]))
        device.min_sn_fl_2ph = _safe_min([term.min_sn_fl_2ph for term in device.sect_terms if term.min_sn_fl_2ph > 0])
        device.min_sn_fl_pg = (
            _safe_min([fault_impedance.term_sn_pg_fl(region, term) for term in device.sect_terms if term.min_sn_fl_pg > 0]))
        device.sect_terms = sorted(device.sect_terms, key=lambda term: term.min_fl_pg, reverse=True)


def update_line_data(app: pft.Application, region: str, devices: List[dd.Device]):
    """
    Get max and min fault current seen by the line for faults occurring within the device protection section
    (i.e. not just faults that occur on the line).
    These values are used in the conductor damage assessment.
    :param app:
    :param region:
    :param devices:
    :return:
    """

    all_grids = app.GetCalcRelevantObjects('*.ElmXnet')
    grids = [grid for grid in all_grids if grid.outserv == 0]

    for device in devices:
        lines = device.sect_lines
        sect_term_obs = [term.obj for term in device.sect_terms]
        for line in lines:
            elmlne = line.obj
            lne_cubs = [cub for cub in [elmlne.bus1, elmlne.bus2] if cub is not None]
            lne_term_obs = [cub.cterm for cub in lne_cubs]

            if any(terms in sect_term_obs for terms in lne_term_obs):
                line_terms = [term for term in device.sect_terms if term.obj in lne_term_obs]
                line.max_fl_3ph = max([term.max_fl_3ph for term in line_terms])
                line.max_fl_2ph = max([term.max_fl_2ph for term in line_terms])
                line.max_fl_pg = max([term.max_fl_pg for term in line_terms])

                # Get line min (i.e. min fl for faults downstream of the line in the protection section)
                lne_max_term_obj = [term.obj for term in line_terms if term.max_fl_pg == line.max_fl_pg][0]
                max_lne_cub = [cub for cub in lne_cubs if cub.cterm == lne_max_term_obj][0]

                down_elements = max_lne_cub.GetAll(1, 0)
                # If the external grid is in the downstream list, you're searching in the wrong direction
                if any(item in grids for item in down_elements):
                    down_objs = max_lne_cub.GetAll(0, 0)
                else:
                    down_objs = down_elements

                line_ds_terms = [obj for obj in down_objs if obj.GetClassName() == dd.ElementType.TERM.value]
                if line_ds_terms:
                    try:
                        line.min_fl_3ph = min([term.min_fl_3ph for term in device.sect_terms if term.obj in line_ds_terms])
                        line.min_fl_2ph = min([term.min_fl_2ph for term in device.sect_terms if term.obj in line_ds_terms])
                        line.min_sn_fl_2ph = min([term.min_sn_fl_2ph for term in device.sect_terms if term.obj in line_ds_terms])
                    except (AttributeError, ValueError):
                        app.PrintPlain(line.obj)
                        app.PrintPlain(f"max_lne_cub: {max_lne_cub}")
                        app.PrintPlain(f"line_ds_terms: {line_ds_terms}")
                    line.min_fl_pg = min([fault_impedance.term_pg_fl(region, term) for term in device.sect_terms if term.obj in line_ds_terms])
                    line.min_sn_fl_pg = min([fault_impedance.term_sn_pg_fl(region, term) for term in device.sect_terms if term.obj in line_ds_terms])
                else:
                    line.min_fl_3ph = min([term.min_fl_3ph for term in line_terms])
                    line.min_fl_2ph = min([term.min_fl_2ph for term in line_terms])
                    line.min_sn_fl_2ph = min([term.min_sn_fl_2ph for term in line_terms])
                    line.min_fl_pg = min([fault_impedance.term_pg_fl(region, term) for term in line_terms])
                    line.min_sn_fl_pg = min([fault_impedance.term_sn_pg_fl(region, term) for term in line_terms])
            else:
                line.max_fl_3ph = 0
                line.max_fl_2ph = 0
                line.max_fl_pg = 0
                line.min_fl_3ph = 0
                line.min_fl_2ph = 0
                line.min_fl_pg = 0
                line.min_sn_fl_2ph = 0
                line.min_sn_fl_pg = 0
        device.sect_lines = sorted(device.sect_lines, key=lambda line: line.min_fl_pg, reverse=True)
