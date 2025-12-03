from fdr_open_points import fdr_open_user_input as foui
import script_classes as dd


def main(app):

    # Enables the user to manually stop the script
    app.SetEnableUserBreak(1)
    app.ClearOutputWindow()
    # TODO: needs work
    radial_dic = foui.mesh_feeder_check(app)
    feeder_list = foui.get_feeders(app, radial_dic)
    for feeder in feeder_list:
        open_switches = get_open_points(app, feeder)
        app.PrintPlain(f"Feeder {feeder} open points:")
        if open_switches:
            for site, switch in open_switches.items():
                if switch.GetClassName() == dd.ElementType.SWITCH.value:
                    app.PrintPlain(f"\t{switch}")
                else:
                    app.PrintPlain(f"\t{site} / {switch}")
        else:
            app.PrintPlain(f"\t(None detected)")
        app.PrintPlain(f"Feeder {feeder} Lines out of service:")
    app.PrintPlain("NOTE: Open points to adjacent bulk supply substations may not be "
                   "detected unless the relevant grids are active under the current "
                   "project")


def get_open_points(app, feeder):

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


