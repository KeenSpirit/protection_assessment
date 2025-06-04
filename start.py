import sys
import time
import powerfactory as pf
from pf_protection_helper import *
from fault_study import fault_level_study as fs, fault_level_study as tm
from user_inputs import get_inputs as gi
from user_inputs import study_selection as ss
from oc_plots import plot_relays as pr
from prot_audit import audit as pa
from devices import devices as ds
from cond_damage import conductor_damage as cd
from save_results import save_result as sr
import logging.config
from logging_config import configure_logging as cl
from test_package import test_module

from importlib import reload
reload(gi)
reload(fs)
reload(pr)
reload(pa)
reload(ds)
reload(cd)
reload(tm)
reload(sr)
reload(ss)
reload(test_module)


def main(app):
    """

    :param app:
    :return:
    """


    region = fs.obtain_region(app)

    # with temporary_variation(app):
    study_selections = ss.get_study_selections(app)
    feeders_devices, bu_devices, user_selection, _ = gi.get_input(app, region, study_selections)
    feeders_devices = chk_empty_fdrs(app, feeders_devices)
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
                pr.plot_all_relays(app, selected_devices)
            if 'Protection Audit' in study_selections:
                pa.audit_all_relays(app, selected_devices)
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
        ds.Device(
            element,
            element.fold_id,
            element.fold_id.cterm,
            fs.ph_attr_lookup(element.fold_id.cterm.phtech),
            round(element.fold_id.cterm.uknom, 2),
            None,
            None,
            None,
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


def chk_empty_fdrs(app, feeders_devices):
    """

    :param app:
    :param feeders_devices:
    :return:
    """

    empty_feeders = [feeder for feeder, devices in feeders_devices.items() if devices == []]

    if len(empty_feeders) == len(feeders_devices):
        app.PrintError("No protection devices were detected in the model for the selected feeders. \n"
                       "Please add and configure the required protection devices and re-run the script.")
        sys.exit(0)
    for empty_feeder in empty_feeders:
            app.PrintWarn(f"No protection devices were detected in the model for feeder {empty_feeder}. \n"
                          "This feeder will be excluded from the study.")
            del feeders_devices[empty_feeder]
    return feeders_devices


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
