import time
import powerfactory as pf
from pf_protection_helper import *
from fault_study import fault_level_study as fs, fault_level_study as tm
import get_inputs as gi
from oc_plots import plot_relays as pr
from prot_audit import audit as pa
from devices import devices as ds
from cond_damage import conductor_damage as cd
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


def main(app):

    # pr.test_func(app)
    device = "CRB24A_OCEF_1"
    pr.plot_fix(app, device)
    # project = app.GetActiveProject()
    # derived_proj = project.der_baseproject
    # der_proj_name = derived_proj.GetFullName()
    #
    # regional_model = 'Regional Models'
    # seq_model = 'SEQ'
    #
    # if regional_model in der_proj_name:
    #     # This is a regional model
    #     model=regional_model
    # elif seq_model in der_proj_name:
    #     # This is a SEQ model
    #     model = seq_model
    # else:
    #     msg = (
    #         "The appropriate region for the model could not be found. "
    #         "Please contact the script administrator to resolve this issue."
    #     )
    #     raise RuntimeError(msg)
    #
    # # with temporary_variation(app):
    # all_fault_studies = {}
    # feeders_devices, user_selection, _ = gi.get_input(app, model)
    # for feeder, sites in feeders_devices.items():
    #     feeder_obj = app.GetCalcRelevantObjects(feeder + ".ElmFeeder")[0]
    #     devices = fs.fault_study(app, feeder_obj, sites)
    #     cd.cond_damage(app, devices)
    #     pr.plot_all_relays(app, devices)
    #     # pa.audit_all_relays(app, devices)
    #     #
    #     # all_fault_studies[feeder] = devices
    #
    #     # gen_info = gi.get_grid_data(app)


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
