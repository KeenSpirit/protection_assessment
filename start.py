import time
import powerfactory as pf
from importlib import reload
from pf_protection_helper import *
from fault_study import fault_level_study as fs
import get_inputs as gi
# from save_results import save_regional_results as srr
# from save_results import save_seq_results as ssr
from save_results import map_functs as mf
from oc_plots import plot_relays as pr
from prot_audit import audit as pa
from devices import devices as ds
from cond_damage import conductor_damage as cd
import test_module as tm
import logging.config
from logging_config import configure_logging as cl


reload(gi)
reload(fs)
reload(pr)
reload(pa)
reload(ds)
reload(cd)
reload(tm)

def main(app):

    # pr.plot_fix(app)
    project = app.GetActiveProject()
    derived_proj = project.der_baseproject
    der_proj_name = derived_proj.GetFullName()

    regional_model = 'Regional Models'
    seq_model = 'SEQ'

    if regional_model in der_proj_name:
        # This is a regional model
        # with temporary_variation(app):
        # all_fault_studies = {}
        feeders_devices, user_selection, _ = gi.get_input(app, regional_model)
        for feeder, sites in feeders_devices.items():
            site_name_map = mf.site_name_convert(sites)
            feeder_obj = app.GetCalcRelevantObjects(feeder + ".ElmFeeder")[0]
            # tm.fault_study(app, feeder_obj, sites)
            study_results, detailed_fls, line_fls = fs.fault_study(app, feeder_obj, site_name_map)
            cd.cond_damage(app, line_fls, site_name_map)
            device_list = ds.format_devices(study_results, user_selection, site_name_map)
            pr.plot_all_relays(app, device_list)
            pa.audit_all_relays(app, device_list)

            # all_fault_studies[feeder] = device_list


            # gen_info = gi.get_grid_data(app)
    elif seq_model in der_proj_name:
        # This is a SEQ model
        pass
    else:
        msg = (
            "The appropriate region for the model could not be found. "
            "Please contact the script administrator to resolve this issue."
        )
        raise RuntimeError(msg)



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
