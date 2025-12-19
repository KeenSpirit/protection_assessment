import math
import time
import sys
from pf_config import pft
from devices import fuses as ds
from oc_plots import get_rmu_fuses as grf
from pf_protection_helper import create_obj, obtain_region
import script_classes as dd
from typing import List, Dict, Union
from importlib import reload


def plot_all_relays(app: pft.Application, feeder: dd.Feeder, selected_devices: List[dd.Device]):
    """

    :param app:
    :param feeder: string
    :param selected_devices:
    :return:
    """

    app.PrintPlain(f"Generating device time overcurrent plots for {feeder.obj.loc_name}...")
    new_format = new_page_format(app)
    study_case = app.GetActiveStudyCase()
    graphics_board = app.GetFromStudyCase("Graphics Board.SetDesktop")
    study_case.Deactivate()
    drawing_format(graphics_board, new_format)
    plot_folder = create_plot_folder(feeder, graphics_board)

    update_ds_tr_data(app, selected_devices)
    colour_dic = create_colour_dic(feeder.devices)

    # Organise selected devices into parent terminals. Devices that share a parent terminal will be plotted together.
    terminals = set(device.term.loc_name for device in selected_devices)
    device_dic = {
        terminal:
            [device for device in selected_devices if device.term.loc_name == terminal]
        for terminal in terminals
    }

    for devices in device_dic.values():
        if any(device.obj.GetClassName() == dd.ElementType.RELAY.value for device in devices):
            vipage = create_plot(app, plot_folder, colour_dic, devices, feeder.sys_volts, f_type='Ground')
            create_draw_format(vipage)
            if any(device.min_fl_2ph > 0 for device in devices):
               vipage = create_plot(app, plot_folder, colour_dic, devices, feeder.sys_volts, f_type='Phase')
               create_draw_format(vipage)
        else:
            app.PrintPlain('No relays to plot')

    study_case.Activate()
    directory = plot_folder.GetFullName()
    app.PrintPlain(f"Time overcurrent plots saved in PowerFactory to {directory}")


def new_page_format(app: pft.Application):

    prjt = app.GetActiveProject()
    settings = prjt.GetContents('Settings.SetFold')[0]
    formats = create_obj(settings, "Page Formats", "SetFoldpage")
    new_format_name = 'A4'
    new_format = create_obj(formats, new_format_name, "SetFormat")
    new_format.iSizeX = 297
    new_format.iSizeY = 210
    new_format.iLeft = 0
    new_format.iRight = 0
    new_format.iTop = 0
    new_format.iBottom = 0

    return new_format


def drawing_format(graphics_board: pft.SetDesktop, format_graph: pft.SetFormat):
    """

    :param graphics_board:
    :param format_graph:
    :param reset:
    :return:
    """

    draw_form = create_obj(graphics_board, "Drawing Format", "SetGrfpage")
    draw_form.iDrwFrm = 1           # 0 = Portrait, 1 = Landscape
    draw_form.aDrwFrm = format_graph.loc_name     # Format


def update_ds_tr_data(app: pft.Application, device_list: List[dd.Device]):
    """
    Add RMU fuse data to SEQ device max ds trs
    :param app:
    :param device_list:
    :return:
    """

    region = obtain_region(app)
    if region == "SEQ":
        tr_strings_dic = {}
        for device in device_list:
            if device.obj.GetClassName() == dd.ElementType.RELAY.value:
                max_ds_tr = device.max_ds_tr
                if max_ds_tr is None:
                    continue
                tr_name = max_ds_tr.term.cpSubstat.loc_name
                if tr_name[:2] != "SP":
                    if max_ds_tr.load_kva in [1000, 1500]:
                        max_tr_string = tr_name + "_1"
                    else:
                        max_tr_string = tr_name + "_0"
                    tr_strings_dic[device.obj] = max_tr_string
        if tr_strings_dic:
            max_tr_strings = list(tr_strings_dic.values())
            results = grf.get_transformer_specifications(max_tr_strings)
            for device_object, string in tr_strings_dic.items():
                device = [device for device in device_list if device.obj == device_object][0]
                string_result = results[string]
                max_ds_tr = device.max_ds_tr
                max_ds_tr.insulation = string_result['insulation']
                max_ds_tr.impedance = string_result['impedance']


def create_colour_dic(devices: List[dd.Device]):
    """
    Ensure that each device has a uniquely coloured trip curve when OC plots are created
    :param devices:
    :return:
    """

    all_devices = []
    for device in devices:
        us_devices = device.us_devices
        ds_devices = device.ds_devices
        all_devices.append(device)
        all_devices.extend(us_devices+ds_devices)
    unique_devices = []
    for device in all_devices:
        if device not in unique_devices:
            unique_devices.append(device)

    # Colour i=0 is white, so start with i=1
    colour_dic = {device.obj: i+1 for i, device in enumerate(unique_devices)}
    return colour_dic


def create_plot(app: pft.Application, graphics_board: pft.SetDesktop, colour_dic, devices: List[dd.Device], sys_volts: str, f_type: str):
    """

    :param app:
    :param graphics_board:
    :param colour_dic:
    :param devices: list of protection devices with a common parent terminal
    :param sys_volts:
    :param f_type:
    :return:
    """

    date_string = time.strftime("%Y%m%d")

    devices_name = ""
    for device in devices:
        relay_name = device.obj.loc_name
        devices_name = devices_name + "_" + relay_name
    # Create the plot graphic object
    folder_name = f"{devices_name} {f_type} Coord Plot {date_string}"
    vipage = create_obj(graphics_board, folder_name, "SetVipage")
    plot_name = f"{devices_name} {f_type} Coord Plot {date_string}"
    plot = vipage.CreateObject("VisOcplot", plot_name)
    plot.Clear()

    # Add the max and min constants to the plot
    min_fl = plot.CreateObject("VisXvalue", f'{devices_name} min fl')
    max_fl = plot.CreateObject("VisXvalue", f'{devices_name} max fl')
    if f_type == 'Ground':
        min_fl_value = min(relay.min_fl_pg for relay in devices)
        max_fl_value = max(relay.max_fl_pg for relay in devices)
        xvalue_settings(min_fl, 'PG Min FL', min_fl_value)
        xvalue_settings(max_fl, 'PG Max FL', max_fl_value)
    else:
        min_fl_value = min(relay.min_fl_2ph for relay in devices)
        max_fl_value = max(max(relay.max_fl_3ph, relay.max_fl_2ph) for relay in devices)
        xvalue_settings(min_fl, "Ph Min FL", min_fl_value)
        xvalue_settings(max_fl, "Ph Max FL", max_fl_value)

    # Add the ds transformer constants to the plot
    device = [device for device in devices if device.max_ds_tr is not None][0]
    max_ds_tr = device.max_ds_tr
    if max_ds_tr is not None and max_ds_tr.term is not None:
        tr_term = max_ds_tr.term
        tr_name = tr_term.cpSubstat.loc_name
        tr_max_pg = max_ds_tr.max_pg
        tr_max_ph = max_ds_tr.max_ph
        max_tr_fl = plot.CreateObject("VisXvalue", f'{tr_name} max fl')
        if f_type == 'Ground':
            xvalue_settings(max_tr_fl, 'DS TR PG Max FL', tr_max_pg)
        else:
            xvalue_settings(max_tr_fl, "DS TR Ph Max FL",  tr_max_ph)
        tr_term_dataclass = [term for term in device.sect_terms if term.obj == max_ds_tr.term][0]
        ds_fuse = ds.create_fuse(app, max_ds_tr, tr_term_dataclass, sys_volts)
        if not ds_fuse:
            app.PrintPlain(
                f'Could not find fuse element for {tr_name} in PowerFactory'
            )
    else:
        ds_fuse = []
    # Add all the devices to the plot
    us_devices = [device.obj for device in devices[0].us_devices]
    ds_devices = [device.obj for device in devices[0].ds_devices]

    all_devices = [device.obj for device in devices] + us_devices + ds_devices + ds_fuse

    for device in all_devices:
        if device.GetClassName() == 'ElmRelay':
            colour = colour_dic[device]
        else:
            colour = 10      # fuse colour
        plot.AddRelay(device, colour)
    plot_settings(plot, devices[0], f_type)

    return vipage


def create_draw_format(vipage: pft.SetVipage):

    draw_format = create_obj(vipage, "Drawing Format", "SetGrfpage")
    draw_format.iDrwFrm = 1
    draw_format.aDrwFrm = 'A4'


def plot_settings(plot: pft.VisOcplot, relay: pft.ElmRelay, f_type: str):
    """
    Apply plot parameters
    :param plot:
    :param relay:
    :param f_type:
    :return:
    """

    def _get_bound(num: float, bound: str):
        order_of_mag = 10 ** int(math.log10(num))
        if bound == 'Min':
            val = math.floor(num / order_of_mag) * order_of_mag
        else:
            val = math.ceil(num / order_of_mag) * order_of_mag
        return val

    if f_type == 'Ground':
        x_min = _get_bound(relay.min_fl_pg, bound='Min')
        x_max = _get_bound(relay.max_fl_pg, bound='Max')
        # Set the curve style for earth fault elements to dashed
        plot.gStyle = [10 for _ in range(len(plot.gStyle))]
    else:
        max_fl_ph = max(relay.max_fl_2ph, relay.max_fl_3ph)
        x_min = _get_bound(relay.min_fl_2ph, bound='Min')
        x_max = _get_bound(max_fl_ph, bound='Max')

    setocplt(plot, f_type)

    plot.x_max = x_max
    plot.x_min = x_min
    plot.x_map = 1                      # log scale
    plot.y_max = 10
    plot.y_min = 0.01
    plot.y_map = 1                      # log scale
    plot.y_fit = 0                      # Adapt scale


def setocplt(plot: pft.VisOcplot, f_type: str):
    """
    Create plot settings
    :param plot:
    :param f_type:
    :return:
    """

    settings = create_obj(plot, "Overcurrent Plot Settings", "SetOcplt")
    settings.unit = 0                   # Show primary current
    if f_type == 'Ground':
            settings.ishow = 2          # Phase and Earth Realys
    else:
        settings.ishow = 1              # Phase Relays
    settings.ishowminmax = 0            # Characteristic - All
    settings.ishowdir = 0               # Direction - All
    settings.ishowtframe = 0            # Recloser Operation - All
    settings.ishowcalc = 1              # Display Automatically - Current

    settings.iTbrk = 0                  # Consider Breaker Opening Time - No
    settings.iushow = 0                 # Voltage Reference Axis - all
    settings.imarg = 1                  # Show Grading Marings while Drag & Drop - Yes


def xvalue_settings(constant: pft.VisXvalue, name: str, value: float):
    """
    Fault current settings
    :param constant:
    :param name:
    :param value:
    :return:
    """

    constant.loc_name = name            # name
    constant.label = 1
    constant.lab_text = [name]          # line label
    constant.show = 1                   # line with intersections
    constant.iopt_lab= 3                # position
    constant.value= value
    constant.color = 1                  # line with intersections
    #constant.style = 1
    constant.width = 5
    constant.xis = 0                    # Current

def create_plot_folder(feeder: dd.Feeder, graphics_board: pft.SetDesktop):

    date_string = time.strftime("%Y%m%d-%H%M%S")
    title = f'{date_string} {feeder.obj.loc_name} Prot Coord Plots'

    return graphics_board.CreateObject("IntFolder", title)

def create_folder(app: pft.Application):

    title = 'Protection Coordination Studies'

    prjt = app.GetActiveProject()
    existing_study_fold = [
        folder.loc_name for folder in prjt.GetContents("*.IntFolder")
    ]
    if title not in existing_study_fold:
        prjt.CreateObject("IntFolder", title)
