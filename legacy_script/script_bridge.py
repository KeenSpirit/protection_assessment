import math
import sys
sys.path.append(r"\\Ecasd01\WksMgmt\PowerFactory\ScriptsDEV\PowerFactoryTyping")
import powerfactorytyping as pft
from legacy_script import save_results as sr
from importlib import reload

reload(sr)

def bridge_results(app, external_grid, feeders):

    sub_name = substation_name(app)

    feeders_devices_inrush = {}
    results_max_3p = {}
    results_max_2p = {}
    results_max_pg = {}
    results_min_2p = {}
    results_min_3p = {}
    results_min_pg = {}
    result_sys_norm_min_2p = {}
    result_sys_norm_min_pg = {}
    feeders_sections_trmax_size = {}
    results_max_tr_3p = {}
    results_max_tr_pg = {}
    results_all_max_3p = {}
    results_all_max_2p = {}
    results_all_max_pg = {}
    results_all_min_3p = {}
    results_all_min_2p = {}
    results_all_min_pg = {}
    result_all_sys_norm_min_2p = {}
    result_all_sys_norm_min_pg = {}
    feeders_devices_load = {}
    results_lines_max_3p = {}
    results_lines_max_2p = {}
    results_lines_max_pg = {}
    results_lines_min_3p = {}
    results_lines_min_2p = {}
    results_lines_min_pg = {}
    result_lines_sys_norm_min_2p = {}
    result_lines_sys_norm_min_pg = {}
    result_lines_type = {}
    result_lines_therm_rating = {}
    fdrs_open_switches = {}

    for fdr in feeders:
        feeder = fdr.obj.loc_name
        devices = fdr.devices
        devices_inrush = {}
        devices_max_3p = {}
        devices_max_2p = {}
        devices_max_pg = {}
        devices_min_2p = {}
        devices_min_3p = {}
        devices_min_pg = {}
        devices_sys_norm_min_2p = {}
        devices_sys_norm_min_pg = {}
        sections_trmax_size = {}
        devices_max_tr_3p = {}
        devices_max_tr_pg = {}
        devices_all_max_3p = {}
        devices_all_max_2p = {}
        devices_all_max_pg = {}
        devices_all_min_3p = {}
        devices_all_min_2p = {}
        devices_all_min_pg = {}
        devices_all_sys_norm_min_2p = {}
        devices_all_sys_norm_min_pg = {}
        devices_load = {}
        devices_lines_max_3p = {}
        devices_lines_max_2p = {}
        devices_lines_max_pg = {}
        devices_lines_min_3p = {}
        devices_lines_min_2p = {}
        devices_lines_min_pg = {}
        devices_lines_sys_norm_min_2p = {}
        devices_lines_sys_norm_min_pg = {}
        devices_lines_type = {}
        devices_lines_therm_rating = {}
        for device in devices:
            devices_inrush[device.obj.loc_name] = device.ds_capacity * 12 / (11 * math.sqrt(3))
            devices_max_3p[device.obj.loc_name] = {[term.obj.loc_name for term in device.sect_terms if term.max_fl_3ph == device.max_fl_3ph][0]: device.max_fl_3ph}
            devices_max_2p[device.obj.loc_name] = {[term.obj.loc_name for term in device.sect_terms if term.max_fl_2ph == device.max_fl_2ph][0]: device.max_fl_2ph}
            devices_max_pg[device.obj.loc_name] = {[term.obj.loc_name for term in device.sect_terms if term.max_fl_pg == device.max_fl_pg][0]: device.max_fl_pg}
            devices_min_3p[device.obj.loc_name] = {[term.obj.loc_name for term in device.sect_terms if term.min_fl_3ph == device.min_fl_3ph][0]: device.min_fl_3ph}
            devices_min_2p[device.obj.loc_name] = {[term.obj.loc_name for term in device.sect_terms if term.min_fl_2ph == device.min_fl_2ph][0]: device.min_fl_2ph}
            devices_min_pg[device.obj.loc_name] = {[term.obj.loc_name for term in device.sect_terms if term.min_fl_pg == device.min_fl_pg][0]: device.min_fl_pg}
            devices_sys_norm_min_2p[device.obj.loc_name] = {[term.obj.loc_name for term in device.sect_terms if term.min_sn_fl_2ph == device.min_sn_fl_2ph][0]: device.min_sn_fl_2ph}
            devices_sys_norm_min_pg[device.obj.loc_name] = {[term.obj.loc_name for term in device.sect_terms if term.min_sn_fl_pg == device.min_sn_fl_pg][0]: device.min_sn_fl_pg}
            sections_trmax_size[device.obj.loc_name] = {[load.obj.loc_name for load in device.sect_loads if load.obj == device.max_ds_tr.obj][0]: device.max_ds_tr.load_kva}
            devices_max_tr_3p[device.obj.loc_name] = {[load.obj.loc_name for load in device.sect_loads if load.obj == device.max_ds_tr.obj][0]: device.max_ds_tr.max_ph}
            devices_max_tr_pg[device.obj.loc_name] = {[load.obj.loc_name for load in device.sect_loads if load.obj == device.max_ds_tr.obj][0]: device.max_ds_tr.max_pg}
            devices_all_max_3p[device.obj.loc_name] = {term.obj.loc_name: getattr(term, 'max_fl_3ph', 0) for term in device.sect_terms}
            devices_all_max_2p[device.obj.loc_name] = {term.obj.loc_name: getattr(term, 'max_fl_2ph', 0) for term in device.sect_terms}
            devices_all_max_pg[device.obj.loc_name] = {term.obj.loc_name: getattr(term, 'max_fl_pg', 0) for term in device.sect_terms}
            devices_all_min_3p[device.obj.loc_name] = {term.obj.loc_name: getattr(term, 'min_fl_3ph', 0) for term in device.sect_terms}
            devices_all_min_2p[device.obj.loc_name] = {term.obj.loc_name: getattr(term, 'min_fl_2ph', 0) for term in device.sect_terms}
            devices_all_min_pg[device.obj.loc_name] = {term.obj.loc_name: getattr(term, 'min_fl_pg', 0) for term in device.sect_terms}
            devices_all_sys_norm_min_2p[device.obj.loc_name] = {term.obj.loc_name: getattr(term, 'min_sn_fl_2ph', 0) for term in device.sect_terms}
            devices_all_sys_norm_min_pg[device.obj.loc_name] = {term.obj.loc_name: getattr(term, 'min_sn_fl_pg', 0) for term in device.sect_terms}

            loads = {term.obj.loc_name: "" for term in device.sect_terms}
            sect_loads = [load.term.loc_name for load in device.sect_loads]
            for load, value in loads.items():
                if load in sect_loads:
                    value = [draw.load_kva for draw in device.sect_loads if draw.term.loc_name == load][0]
                    loads[load] = value

            devices_load[device.obj.loc_name] = loads
            devices_lines_max_3p[device.obj.loc_name] = {line.obj: getattr(line, 'max_fl_3ph', 0) for line in device.sect_lines}
            devices_lines_max_2p[device.obj.loc_name] = {line.obj: getattr(line, 'max_fl_2ph', 0) for line in device.sect_lines}
            devices_lines_max_pg[device.obj.loc_name] = {line.obj: getattr(line, 'max_fl_pg', 0) for line in device.sect_lines}
            devices_lines_min_3p[device.obj.loc_name] = {line.obj: getattr(line, 'min_fl_3ph', 0) for line in device.sect_lines}
            devices_lines_min_2p[device.obj.loc_name] = {line.obj: getattr(line, 'min_fl_2ph', 0) for line in device.sect_lines}
            devices_lines_min_pg[device.obj.loc_name] = {line.obj: getattr(line, 'min_fl_pg', 0) for line in device.sect_lines}
            devices_lines_sys_norm_min_2p[device.obj.loc_name] = {line.obj: getattr(line, 'min_sn_fl_2ph', 0) for line in device.sect_lines}
            devices_lines_sys_norm_min_pg[device.obj.loc_name] = {line.obj: getattr(line, 'min_sn_fl_pg', 0) for line in device.sect_lines}
            devices_lines_type[device.obj.loc_name] = {line.obj: getattr(line, 'line_type', 0) for line in device.sect_lines}
            devices_lines_therm_rating[device.obj.loc_name] = {line.obj: getattr(line, 'thermal_rating', 0) for line in device.sect_lines}
        feeders_devices_inrush[feeder] = devices_inrush
        results_max_3p[feeder] = devices_max_3p
        results_max_2p[feeder] = devices_max_2p
        results_max_pg[feeder] = devices_max_pg
        results_min_3p[feeder] = devices_min_3p
        results_min_2p[feeder] = devices_min_2p
        results_min_pg[feeder] = devices_min_pg
        result_sys_norm_min_2p[feeder] = devices_sys_norm_min_2p
        result_sys_norm_min_pg[feeder] = devices_sys_norm_min_pg
        feeders_sections_trmax_size[feeder] = sections_trmax_size
        results_max_tr_3p[feeder] = devices_max_tr_3p
        results_max_tr_pg[feeder] = devices_max_tr_pg
        results_all_max_3p[feeder] = devices_all_max_3p
        results_all_max_2p[feeder] = devices_all_max_2p
        results_all_max_pg[feeder] = devices_all_max_pg
        results_all_min_3p[feeder] = devices_all_min_3p
        results_all_min_2p[feeder] = devices_all_min_2p
        results_all_min_pg[feeder] = devices_all_min_pg
        result_all_sys_norm_min_2p[feeder] = devices_all_sys_norm_min_2p
        result_all_sys_norm_min_pg[feeder] = devices_all_sys_norm_min_pg
        feeders_devices_load[feeder] = devices_load
        results_lines_max_3p[feeder] = devices_lines_max_3p
        results_lines_max_2p[feeder] = devices_lines_max_2p
        results_lines_max_pg[feeder] = devices_lines_max_pg
        results_lines_min_3p[feeder] = devices_lines_min_3p
        results_lines_min_2p[feeder] = devices_lines_min_2p
        results_lines_min_pg[feeder] = devices_lines_min_pg
        result_lines_sys_norm_min_2p[feeder] = devices_lines_sys_norm_min_2p
        result_lines_sys_norm_min_pg[feeder] = devices_lines_sys_norm_min_pg
        result_lines_type[feeder] = devices_lines_type
        result_lines_therm_rating[feeder] = devices_lines_therm_rating
        fdrs_open_switches[feeder] = fdr.open_points

    output = sr.output_results(app, sub_name, external_grid, feeders_devices_inrush,
                            results_max_3p, results_max_2p, results_max_pg, results_min_2p, results_min_3p,
                            results_min_pg,
                            result_sys_norm_min_2p, result_sys_norm_min_pg, feeders_sections_trmax_size,
                            results_max_tr_3p,
                            results_max_tr_pg, results_all_max_3p, results_all_max_2p, results_all_max_pg,
                            results_all_min_3p,
                            results_all_min_2p, results_all_min_pg, result_all_sys_norm_min_2p,
                            result_all_sys_norm_min_pg,
                            feeders_devices_load, results_lines_max_3p, results_lines_max_2p, results_lines_max_pg,
                            results_lines_min_3p, results_lines_min_2p, results_lines_min_pg,
                            result_lines_sys_norm_min_2p,
                            result_lines_sys_norm_min_pg, result_lines_type, result_lines_therm_rating,
                            fdrs_open_switches)
    sr.save_results(app, sub_name, output)
    # values = [sub_name, external_grid, feeders_devices_inrush,
    #         results_max_3p, results_max_2p, results_max_pg, results_min_2p, results_min_3p, results_min_pg,
    #         result_sys_norm_min_2p, result_sys_norm_min_pg, feeders_sections_trmax_size, results_max_tr_3p,
    #         results_max_tr_pg, results_all_max_3p, results_all_max_2p, results_all_max_pg, results_all_min_3p,
    #         results_all_min_2p, results_all_min_pg, result_all_sys_norm_min_2p, result_all_sys_norm_min_pg,
    #         feeders_devices_load, results_lines_max_3p, results_lines_max_2p, results_lines_max_pg,
    #         results_lines_min_3p, results_lines_min_2p, results_lines_min_pg, result_lines_sys_norm_min_2p,
    #         result_lines_sys_norm_min_pg, result_lines_type, result_lines_therm_rating, fdrs_open_switches
    #           ]
    # for value in values:
    #     app.PrintPlain(value)


def substation_name(app):
    """Get the substation name"""

    sub_names = app.GetCalcRelevantObjects('*.ElmNet')
    subs = []
    for sub in sub_names:
        if sub.loc_name != "New Elements" and sub.loc_name != "Summary Grid" and sub.loc_name != "Boundary Subs":
            subs.append(sub.loc_name + " 11kV")
    sub_name = '_'.join(subs)
    return sub_name
