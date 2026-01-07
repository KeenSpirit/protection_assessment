"""
Data structure bridge for legacy SEQ fault level study output.

This module converts the modern dataclass-based study results into the
legacy dictionary format required by the original SEQ Excel output
module. It maintains backward compatibility with existing SEQ output
workflows.

The bridge extracts fault level data from Device and Feeder dataclasses
and reorganizes it into nested dictionaries keyed by feeder and device
names, matching the structure expected by save_results.output_results().

Functions:
    bridge_results: Main entry point for legacy output generation
    substation_name: Extract substation name from active grids
"""

import math
from typing import Any, Dict, List

from pf_config import pft
from legacy_script import save_results as sr
from importlib import reload

reload(sr)


def bridge_results(
    app: pft.Application,
    external_grid: Dict,
    feeders: List
) -> None:
    """
    Convert dataclass results to legacy format and generate output.

    Transforms the modern Feeder and Device dataclass structures into
    the nested dictionary format required by the legacy SEQ Excel
    output module.

    Args:
        app: PowerFactory application instance.
        external_grid: Dictionary of external grid parameters.
        feeders: List of Feeder dataclasses with populated devices.

    Side Effects:
        Calls save_results module to generate Excel output.

    Data Structure Conversion:
        Device.max_fl_3ph -> results_max_3p[feeder][device][terminal]
        Device.sect_terms -> results_all_max_3p[feeder][device][terminal]
        Device.sect_lines -> results_lines_max_3p[feeder][device][line]

    Note:
        This bridge exists only for backward compatibility with SEQ
        Protection department legacy workflows. New implementations
        should use the modern save_result module directly.
    """
    sub_name = substation_name(app)

    # Initialize feeder-level dictionaries
    feeders_devices_inrush: Dict[str, Dict] = {}
    results_max_3p: Dict[str, Dict] = {}
    results_max_2p: Dict[str, Dict] = {}
    results_max_pg: Dict[str, Dict] = {}
    results_min_2p: Dict[str, Dict] = {}
    results_min_3p: Dict[str, Dict] = {}
    results_min_pg: Dict[str, Dict] = {}
    result_sys_norm_min_2p: Dict[str, Dict] = {}
    result_sys_norm_min_pg: Dict[str, Dict] = {}
    feeders_sections_trmax_size: Dict[str, Dict] = {}
    results_max_tr_3p: Dict[str, Dict] = {}
    results_max_tr_pg: Dict[str, Dict] = {}
    results_all_max_3p: Dict[str, Dict] = {}
    results_all_max_2p: Dict[str, Dict] = {}
    results_all_max_pg: Dict[str, Dict] = {}
    results_all_min_3p: Dict[str, Dict] = {}
    results_all_min_2p: Dict[str, Dict] = {}
    results_all_min_pg: Dict[str, Dict] = {}
    result_all_sys_norm_min_2p: Dict[str, Dict] = {}
    result_all_sys_norm_min_pg: Dict[str, Dict] = {}
    feeders_devices_load: Dict[str, Dict] = {}
    results_lines_max_3p: Dict[str, Dict] = {}
    results_lines_max_2p: Dict[str, Dict] = {}
    results_lines_max_pg: Dict[str, Dict] = {}
    results_lines_min_3p: Dict[str, Dict] = {}
    results_lines_min_2p: Dict[str, Dict] = {}
    results_lines_min_pg: Dict[str, Dict] = {}
    result_lines_sys_norm_min_2p: Dict[str, Dict] = {}
    result_lines_sys_norm_min_pg: Dict[str, Dict] = {}
    result_lines_type: Dict[str, Dict] = {}
    result_lines_therm_rating: Dict[str, Dict] = {}
    fdrs_open_switches: Dict[str, Dict] = {}

    # Process each feeder
    for fdr in feeders:
        feeder = fdr.obj.loc_name
        devices = fdr.devices

        # Initialize device-level dictionaries for this feeder
        feeder_data = _process_feeder_devices(devices)

        # Assign to feeder-level dictionaries
        feeders_devices_inrush[feeder] = feeder_data['inrush']
        results_max_3p[feeder] = feeder_data['max_3p']
        results_max_2p[feeder] = feeder_data['max_2p']
        results_max_pg[feeder] = feeder_data['max_pg']
        results_min_3p[feeder] = feeder_data['min_3p']
        results_min_2p[feeder] = feeder_data['min_2p']
        results_min_pg[feeder] = feeder_data['min_pg']
        result_sys_norm_min_2p[feeder] = feeder_data['sys_norm_min_2p']
        result_sys_norm_min_pg[feeder] = feeder_data['sys_norm_min_pg']
        feeders_sections_trmax_size[feeder] = feeder_data['trmax_size']
        results_max_tr_3p[feeder] = feeder_data['max_tr_3p']
        results_max_tr_pg[feeder] = feeder_data['max_tr_pg']
        results_all_max_3p[feeder] = feeder_data['all_max_3p']
        results_all_max_2p[feeder] = feeder_data['all_max_2p']
        results_all_max_pg[feeder] = feeder_data['all_max_pg']
        results_all_min_3p[feeder] = feeder_data['all_min_3p']
        results_all_min_2p[feeder] = feeder_data['all_min_2p']
        results_all_min_pg[feeder] = feeder_data['all_min_pg']
        result_all_sys_norm_min_2p[feeder] = feeder_data['all_sys_norm_min_2p']
        result_all_sys_norm_min_pg[feeder] = feeder_data['all_sys_norm_min_pg']
        feeders_devices_load[feeder] = feeder_data['load']
        results_lines_max_3p[feeder] = feeder_data['lines_max_3p']
        results_lines_max_2p[feeder] = feeder_data['lines_max_2p']
        results_lines_max_pg[feeder] = feeder_data['lines_max_pg']
        results_lines_min_3p[feeder] = feeder_data['lines_min_3p']
        results_lines_min_2p[feeder] = feeder_data['lines_min_2p']
        results_lines_min_pg[feeder] = feeder_data['lines_min_pg']
        result_lines_sys_norm_min_2p[feeder] = feeder_data['lines_sys_norm_min_2p']
        result_lines_sys_norm_min_pg[feeder] = feeder_data['lines_sys_norm_min_pg']
        result_lines_type[feeder] = feeder_data['lines_type']
        result_lines_therm_rating[feeder] = feeder_data['lines_therm_rating']
        fdrs_open_switches[feeder] = fdr.open_points

    # Generate legacy Excel output
    output = sr.output_results(
        app, sub_name, external_grid, feeders_devices_inrush,
        results_max_3p, results_max_2p, results_max_pg,
        results_min_2p, results_min_3p, results_min_pg,
        result_sys_norm_min_2p, result_sys_norm_min_pg,
        feeders_sections_trmax_size, results_max_tr_3p, results_max_tr_pg,
        results_all_max_3p, results_all_max_2p, results_all_max_pg,
        results_all_min_3p, results_all_min_2p, results_all_min_pg,
        result_all_sys_norm_min_2p, result_all_sys_norm_min_pg,
        feeders_devices_load,
        results_lines_max_3p, results_lines_max_2p, results_lines_max_pg,
        results_lines_min_3p, results_lines_min_2p, results_lines_min_pg,
        result_lines_sys_norm_min_2p, result_lines_sys_norm_min_pg,
        result_lines_type, result_lines_therm_rating, fdrs_open_switches
    )

    sr.save_results(app, sub_name, output)


def _process_feeder_devices(devices: List) -> Dict[str, Dict]:
    """
    Process all devices for a feeder and extract legacy data structures.

    Args:
        devices: List of Device dataclasses.

    Returns:
        Dictionary of device-level data dictionaries.
    """
    # Device-level dictionaries
    devices_inrush: Dict[str, Any] = {}
    devices_max_3p: Dict[str, Dict] = {}
    devices_max_2p: Dict[str, Dict] = {}
    devices_max_pg: Dict[str, Dict] = {}
    devices_min_2p: Dict[str, Dict] = {}
    devices_min_3p: Dict[str, Dict] = {}
    devices_min_pg: Dict[str, Dict] = {}
    devices_sys_norm_min_2p: Dict[str, Dict] = {}
    devices_sys_norm_min_pg: Dict[str, Dict] = {}
    sections_trmax_size: Dict[str, Dict] = {}
    devices_max_tr_3p: Dict[str, Dict] = {}
    devices_max_tr_pg: Dict[str, Dict] = {}
    devices_all_max_3p: Dict[str, Dict] = {}
    devices_all_max_2p: Dict[str, Dict] = {}
    devices_all_max_pg: Dict[str, Dict] = {}
    devices_all_min_3p: Dict[str, Dict] = {}
    devices_all_min_2p: Dict[str, Dict] = {}
    devices_all_min_pg: Dict[str, Dict] = {}
    devices_all_sys_norm_min_2p: Dict[str, Dict] = {}
    devices_all_sys_norm_min_pg: Dict[str, Dict] = {}
    devices_load: Dict[str, Dict] = {}
    devices_lines_max_3p: Dict[str, Dict] = {}
    devices_lines_max_2p: Dict[str, Dict] = {}
    devices_lines_max_pg: Dict[str, Dict] = {}
    devices_lines_min_3p: Dict[str, Dict] = {}
    devices_lines_min_2p: Dict[str, Dict] = {}
    devices_lines_min_pg: Dict[str, Dict] = {}
    devices_lines_sys_norm_min_2p: Dict[str, Dict] = {}
    devices_lines_sys_norm_min_pg: Dict[str, Dict] = {}
    devices_lines_type: Dict[str, Dict] = {}
    devices_lines_therm_rating: Dict[str, Dict] = {}

    for device in devices:
        dev_name = device.obj.loc_name

        # Device-level fault currents
        devices_inrush[dev_name] = 0
        devices_max_3p[dev_name] = {dev_name: device.max_fl_3ph}
        devices_max_2p[dev_name] = {dev_name: device.max_fl_2ph}
        devices_max_pg[dev_name] = {dev_name: device.max_fl_pg}
        devices_min_3p[dev_name] = {dev_name: device.min_fl_3ph}
        devices_min_2p[dev_name] = {dev_name: device.min_fl_2ph}
        devices_min_pg[dev_name] = {dev_name: device.min_fl_pg}
        devices_sys_norm_min_2p[dev_name] = {dev_name: device.min_sn_fl_2ph}
        devices_sys_norm_min_pg[dev_name] = {dev_name: device.min_sn_fl_pg}

        # Transformer data
        max_ds_tr = device.max_ds_tr
        sections_trmax_size[dev_name] = {
            dev_name: getattr(max_ds_tr, 'load_kva', 0)
        }
        devices_max_tr_3p[dev_name] = {
            dev_name: getattr(max_ds_tr, 'max_ph', 0)
        }
        devices_max_tr_pg[dev_name] = {
            dev_name: getattr(max_ds_tr, 'max_pg', 0)
        }

        # Terminal-level data
        devices_all_max_3p[dev_name] = {
            term.obj.loc_name: getattr(term, 'max_fl_3ph', 0)
            for term in device.sect_terms
        }
        devices_all_max_2p[dev_name] = {
            term.obj.loc_name: getattr(term, 'max_fl_2ph', 0)
            for term in device.sect_terms
        }
        devices_all_max_pg[dev_name] = {
            term.obj.loc_name: getattr(term, 'max_fl_pg', 0)
            for term in device.sect_terms
        }
        devices_all_min_3p[dev_name] = {
            term.obj.loc_name: getattr(term, 'min_fl_3ph', 0)
            for term in device.sect_terms
        }
        devices_all_min_2p[dev_name] = {
            term.obj.loc_name: getattr(term, 'min_fl_2ph', 0)
            for term in device.sect_terms
        }
        devices_all_min_pg[dev_name] = {
            term.obj.loc_name: getattr(term, 'min_fl_pg', 0)
            for term in device.sect_terms
        }
        devices_all_sys_norm_min_2p[dev_name] = {
            term.obj.loc_name: getattr(term, 'min_sn_fl_2ph', 0)
            for term in device.sect_terms
        }
        devices_all_sys_norm_min_pg[dev_name] = {
            term.obj.loc_name: getattr(term, 'min_sn_fl_pg', 0)
            for term in device.sect_terms
        }
        devices_load[dev_name] = {
            load.loc_name: getattr(load, 'Strat', 0)
            for load in device.sect_loads
        }

        # Line-level data
        devices_lines_max_3p[dev_name] = {
            line.obj: getattr(line, 'max_fl_3ph', 0)
            for line in device.sect_lines
        }
        devices_lines_max_2p[dev_name] = {
            line.obj: getattr(line, 'max_fl_2ph', 0)
            for line in device.sect_lines
        }
        devices_lines_max_pg[dev_name] = {
            line.obj: getattr(line, 'max_fl_pg', 0)
            for line in device.sect_lines
        }
        devices_lines_min_3p[dev_name] = {
            line.obj: getattr(line, 'min_fl_3ph', 0)
            for line in device.sect_lines
        }
        devices_lines_min_2p[dev_name] = {
            line.obj: getattr(line, 'min_fl_2ph', 0)
            for line in device.sect_lines
        }
        devices_lines_min_pg[dev_name] = {
            line.obj: getattr(line, 'min_fl_pg', 0)
            for line in device.sect_lines
        }
        devices_lines_sys_norm_min_2p[dev_name] = {
            line.obj: getattr(line, 'min_sn_fl_2ph', 0)
            for line in device.sect_lines
        }
        devices_lines_sys_norm_min_pg[dev_name] = {
            line.obj: getattr(line, 'min_sn_fl_pg', 0)
            for line in device.sect_lines
        }
        devices_lines_type[dev_name] = {
            line.obj: getattr(line, 'line_type', 0)
            for line in device.sect_lines
        }
        devices_lines_therm_rating[dev_name] = {
            line.obj: getattr(line, 'thermal_rating', 0)
            for line in device.sect_lines
        }

    return {
        'inrush': devices_inrush,
        'max_3p': devices_max_3p,
        'max_2p': devices_max_2p,
        'max_pg': devices_max_pg,
        'min_3p': devices_min_3p,
        'min_2p': devices_min_2p,
        'min_pg': devices_min_pg,
        'sys_norm_min_2p': devices_sys_norm_min_2p,
        'sys_norm_min_pg': devices_sys_norm_min_pg,
        'trmax_size': sections_trmax_size,
        'max_tr_3p': devices_max_tr_3p,
        'max_tr_pg': devices_max_tr_pg,
        'all_max_3p': devices_all_max_3p,
        'all_max_2p': devices_all_max_2p,
        'all_max_pg': devices_all_max_pg,
        'all_min_3p': devices_all_min_3p,
        'all_min_2p': devices_all_min_2p,
        'all_min_pg': devices_all_min_pg,
        'all_sys_norm_min_2p': devices_all_sys_norm_min_2p,
        'all_sys_norm_min_pg': devices_all_sys_norm_min_pg,
        'load': devices_load,
        'lines_max_3p': devices_lines_max_3p,
        'lines_max_2p': devices_lines_max_2p,
        'lines_max_pg': devices_lines_max_pg,
        'lines_min_3p': devices_lines_min_3p,
        'lines_min_2p': devices_lines_min_2p,
        'lines_min_pg': devices_lines_min_pg,
        'lines_sys_norm_min_2p': devices_lines_sys_norm_min_2p,
        'lines_sys_norm_min_pg': devices_lines_sys_norm_min_pg,
        'lines_type': devices_lines_type,
        'lines_therm_rating': devices_lines_therm_rating,
    }


def substation_name(app: pft.Application) -> str:
    """
    Extract substation name from active calculation-relevant grids.

    Builds a composite name from all active substations, excluding
    system folders like 'New Elements', 'Summary Grid', and
    'Boundary Subs'.

    Args:
        app: PowerFactory application instance.

    Returns:
        Underscore-separated string of substation names with '11kV'
        suffix (e.g., 'ABY 11kV_BDB 11kV').

    Example:
        >>> name = substation_name(app)
        >>> print(name)
        'Abermain 11kV'
    """
    sub_names = app.GetCalcRelevantObjects('*.ElmNet')

    excluded = {"New Elements", "Summary Grid", "Boundary Subs"}
    subs = [
        sub.loc_name + " 11kV"
        for sub in sub_names
        if sub.loc_name not in excluded
    ]

    return '_'.join(subs)