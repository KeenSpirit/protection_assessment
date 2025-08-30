import time
import powerfactory as pf
from pf_protection_helper import *
import model_checks
from fault_study import fault_level_study as fs, fault_level_study as tm
from user_inputs import get_inputs as gi
from user_inputs import study_selection as ss
import oc_plots
from importlib import reload
reload(oc_plots)
from oc_plots import plot_relays as pr, get_rmu_fuses as grf
from devices import fuses as ds
from devices import dataclass_definitions as dd
from cond_damage import conductor_damage as cd
from save_results import save_result as sr
import logging.config
from logging_config import configure_logging as cl
from test_package import test_module

# from importlib import reload
reload(model_checks)
reload(gi)
reload(fs)
reload(pr)
reload(ds)
reload(cd)
reload(tm)
reload(sr)
reload(ss)
reload(grf)
reload(test_module)


def main(app):
    """

    :param app:
    :return:
    """

    # Undertake model checks
    all_relays = get_all_relays(app)
    model_checks.relay_checks(app, all_relays)

    region = obtain_region(app)
    # with temporary_variation(app):
    study_selections = ss.get_study_selections(app)
    feeders_devices, bu_devices, user_selection, _ = gi.get_input(app, region, study_selections)

    feeders_devices = model_checks.chk_empty_fdrs(app, feeders_devices)

    feeders_devices = convert_to_dataclass(feeders_devices)
    bu_devices = convert_to_dataclass(bu_devices)

    all_fault_studies = {}
    for feeder, devices in feeders_devices.items():
        feeder_obj = app.GetCalcRelevantObjects(feeder + ".ElmFeeder")[0]
        devices = fs.fault_study(app, region, feeder_obj, bu_devices, devices)
        if user_selection:
            selected_devices = [device for device in devices if device.object in user_selection[feeder]]
            if 'Conductor Damage Assessment' in study_selections:
                cd.cond_damage(app, selected_devices)
            if 'Protection Relay Coordination Plot' in study_selections:
                system_volts = feeder_obj.cn_bus.uknom
                pr.plot_all_relays(app, devices, selected_devices, system_volts)
        all_fault_studies[feeder] = devices
    gen_info = gi.get_grid_data(app)
    sr.save_dataframe(app, region, study_selections, gen_info, all_fault_studies)


def convert_to_dataclass(dictionary):
    """
    Convert a dictionary pf protection elements in to a dictionary of device dataclasses
    :param dictionary:
    :return:
    """

    new_dictionary = {}
    for key, value in dictionary.items():
        new_dictionary[key] = [
        dd.Device(
            element,
            element.fold_id,
            element.fold_id.cterm,
            ds.ph_attr_lookup(element.fold_id.cterm.phtech),
            round(element.fold_id.cterm.uknom, 2),
            None,
            None,
            None,
            None,
            None,
            None,
            [],
            [],
            [],
            [],
            []
            )
        for element in value
        ]
    return new_dictionary


def get_all_relays(app):
    net_mod = app.GetProjectFolder("netmod")
    # Filter for relays under network model recursively.
    all_relays = net_mod.GetContents("*.ElmRelay", True)
    relays = [
        relay
        for relay in all_relays
        if relay.cpGrid
        if relay.cpGrid.IsCalcRelevant()
        if relay.GetParent().GetClassName() == "StaCubic"
        if not relay.IsOutOfService()
    ]
    return relays


if __name__ == '__main__':
    start = time.time()

    # Configure logging
    logging.basicConfig(
        filename=cl.getpath() / 'prot_assess_log.txt',
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    app = pf.GetApplication()

    with app_manager(app, gui=True) as app:
        main(app)

    end = time.time()
    run_time = round(end - start, 6)
    app.PrintPlain(f"Script run time: {run_time} seconds")
