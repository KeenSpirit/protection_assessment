import sys
from pf_config import pft
from typing import List, Dict
from fdr_open_points import fdr_open_user_input as foui
import script_classes as dd
from importlib import reload
reload(foui)

def main(app: pft.Application):

    # Enables the user to manually stop the script
    app.SetEnableUserBreak(1)
    app.ClearOutputWindow()
    radial_dic = foui.mesh_feeder_check(app)
    feeder_list = foui.get_feeders(app, radial_dic)
    for fdr in feeder_list:
        feeder = app.GetCalcRelevantObjects(fdr + ".ElmFeeder")[0]
        feeder = dd.initialise_fdr_dataclass(feeder)
        get_open_points(app, feeder)
        app.PrintPlain(f"Feeder {feeder.obj} open points:")
        if feeder.open_points:
            for site, switch in feeder.open_points.items():
                if switch.GetClassName() == dd.ElementType.SWITCH.value:
                    app.PrintPlain(f"\t{switch}")
                else:
                    app.PrintPlain(f"\t{site} / {switch}")
        else:
            app.PrintPlain(f"\t(None detected)")
    app.PrintPlain("NOTE: Open points to adjacent bulk supply substations may not be "
                   "detected unless the relevant grids are active under the current "
                   "project")


def get_open_points(app:pft.Application, feeder: dd.Feeder):

    all_staswitch = app.GetProjectFolder("netdat").GetContents("*.StaSwitch", 1)
    all_elmcoup = app.GetProjectFolder("netdat").GetContents("*.ElmCoup", 1)

    line_list = feeder.obj.GetObjs('ElmLne')
    terminal_list = []
    for line in line_list:
        terminal_list.extend(line.GetConnectedElements())
    terminal_list = list(set(terminal_list))
    open_switches = {}
    for switch in all_staswitch:
        # Get the terminal associated with the StaSwitch.
        cubicle = switch.GetAttribute("fold_id")
        switch_terminal = cubicle.GetAttribute("cterm")
        # If that terminal is in the feeder terminal_list list, add it to the feeder switch list.
        if switch.GetAttribute("on_off") == 0 and switch_terminal in terminal_list:
            open_switches[switch] = switch

    for switch in all_elmcoup:
        terminals = switch.GetConnectedElements()
        if switch.GetAttribute("on_off") == 0 and any(term in terminals for term in terminal_list):
            open_switches[switch.GetAttribute("fold_id")] = switch

    feeder.open_points = open_switches
    return


