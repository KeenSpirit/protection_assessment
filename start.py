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
# from save_results import save_regional_results as srr
# from save_results import save_seq_results as ssr

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


def main(app):
    """

    :param app:
    :return:
    """

    model = fs.obtain_region(app)

    # with temporary_variation(app):
    study_selections = ss.get_study_selections(app)
    feeders_devices, user_selection, _ = gi.get_input(app, model)
    all_fault_studies = {}
    for feeder, sites in feeders_devices.items():
        feeder_obj = app.GetCalcRelevantObjects(feeder + ".ElmFeeder")[0]
        devices = fs.fault_study(app, feeder_obj, sites)
        selected_devices = [device for device in devices if device.object in user_selection[feeder]]
        if 'Conductor Damage Assessment' in study_selections:
            cd.cond_damage(app, selected_devices)
        if 'Protection Relay Coordination Plot' in study_selections:
            pr.plot_all_relays(app, selected_devices)
        if 'Protection Audit' in study_selections:
            pa.audit_all_relays(app, selected_devices)
        all_fault_studies[feeder] = devices
    gen_info = gi.get_grid_data(app)
    sr.save_dataframe(app, gen_info, all_fault_studies)

    # fix relay plots
    # if 'Protection Relay Coordination Plot' in study_selections:
    #     time.sleep(2)
    #     pr.plot_fix(app, user_selection)


if __name__ == '__main__':
    start = time.time()

    # Configure logging
    logging.basicConfig(
        filename=cl.getpath() / 'prot_assess_log.txt',
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )

    with app_manager(pf.GetApplication(), gui=True, echo=True) as app:
        main(app)

    end = time.time()
    run_time = round(end - start, 6)
    app.PrintPlain(f"Script run time: {run_time} seconds")
