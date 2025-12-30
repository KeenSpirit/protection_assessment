import time
import powerfactory as pf
import sys
from pf_config import pft
from typing import Union, Dict, List
import pf_protection_helper as helper
import model_checks
from fault_study import fault_level_study as fs, fault_level_study as tm
from fdr_open_points import get_open_points as gop
from find_substation import find_sub
import script_classes as dd
from user_inputs import get_inputs as gi, study_selection as ss
from legacy_script import script_bridge as sb
from oc_plots import plot_relays as pr, get_rmu_fuses as grf
from devices import fuses as ds, relays
from cond_damage import conductor_damage as cd
from save_results import save_result as sr
import logging.config
from config_logging import configure_logging as cl
from test_package import test_module
from colour_maps import colour_maps as cm

from importlib import reload
reload(model_checks)
reload(gi)
reload(gop)
reload(fs)
reload(pr)
reload(ds)
reload(relays)
reload(cd)
reload(tm)
reload(sr)
reload(ss)
reload(grf)
reload(test_module)
reload(sb)
reload(find_sub)
reload(cm)


def main(app: pft.Application):
    """

    :param app:
    :return:
    """

    # with temporary_variation(app):
    study_selections = ss.get_study_selections(app)
    if "Find PowerFactory Project" in study_selections:
        find_sub.get_project(app)
        return

    prjt = app.GetActiveProject()
    if prjt is None:
        app.PrintWarn("No Active Project, Ending Script")
        return

    if "Find Feeder Open Points" in study_selections:
        gop.main(app)
        return

    # Undertake model checks
    all_relays = relays.get_all_relays(app)
    model_checks.relay_checks(app, all_relays)

    region = helper.obtain_region(app)
    feeders_devices, bu_devices, user_selection, external_grid = gi.get_input(app, region, study_selections)
    feeders = cvrt_fdr_to_dataclass(app, feeders_devices, bu_devices)


    # Turn on all grids momentarily to load all floating terminals in to line connected elements
    user_selected_study_case = app.GetActiveStudyCase()
    switch_study_case(app, user_selected_study_case, all_grids=True)
    switch_study_case(app, user_selected_study_case, all_grids=False)

    for feeder in feeders:
        gop.get_open_points(app, feeder)
        fs.fault_study(app, external_grid, region, feeder, study_selections)
        if user_selection:
            selected_devices = [device for device in feeder.devices if device.obj in user_selection[feeder.obj.loc_name]]
            if 'Conductor Damage Assessment' in study_selections:
                cd.cond_damage(app, feeder.obj.loc_name, selected_devices)
            if 'Protection Relay Coordination Plot' in study_selections:
                pr.plot_all_relays(app, feeder, selected_devices)
    if "Fault Level Study (legacy)" in study_selections:
        sb.bridge_results(app, external_grid, feeders)
    else:
        cm.colour_map(app, region, feeders, study_selections)
        sr.save_dataframe(app, region, study_selections, external_grid, feeders)


def switch_study_case(app: pft.Application, user_selected_study_case: pft.IntCase, all_grids: bool=False):

    if user_selected_study_case is None:
        app.PrintError('Please select a study case and re-run the script.')
        exit(1)
    if all_grids:
        int_case = app.GetProjectFolder("study").GetContents("All Active Grids Study Case")[0]
    else:
        int_case = user_selected_study_case
    int_case.Activate()


def cvrt_fdr_to_dataclass(app: pft.Application, feeders_devices: Dict, bu_devices: Dict) -> List[dd.Feeder]:
    """
    Convert a dictionary pf protection elements in to a list of feeder dataclasses
    :param app:
    :param feeders_devices:
    :param bu_devices:
    :return:
    """

    if bu_devices:
        for grid, grid_devices in bu_devices.items():
            bu_devices[grid] = [dd.initialise_dev_dataclass(device) for device in grid_devices]
    feeders = []
    for fdr, devs in feeders_devices.items():
        feeder = app.GetCalcRelevantObjects(fdr + ".ElmFeeder")[0]
        feeder = dd.initialise_fdr_dataclass(feeder)
        devices = [dd.initialise_dev_dataclass(dev) for dev in devs]
        feeder.devices = devices
        feeder.bu_devices = bu_devices
        feeders.append(feeder)
    return feeders


if __name__ == '__main__':
    start = time.time()

    # Configure logging
    logging.basicConfig(
        filename=cl.getpath() / 'prot_assess_log.txt',
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    app = pf.GetApplication()

    with helper.app_manager(app, gui=True) as app:
        main(app)

    end = time.time()
    run_time = round(end - start, 6)
    app.PrintPlain(f"Script run time: {run_time} seconds")
