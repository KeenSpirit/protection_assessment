import math
from devices import fuses as ds
from oc_plots import get_rmu_fuses as grf
from pf_protection_helper import create_obj, obtain_region
from importlib import reload


def plot_all_relays(app, devices, selected_devices, system_volts):

    app.PrintPlain("Generating device time overcurrent plots...")
    new_format = new_page_format(app)
    study_case = app.GetActiveStudyCase()
    graphics_board = app.GetFromStudyCase("Graphics Board.SetDesktop")
    study_case.Deactivate()
    drawing_format(graphics_board, new_format)
    update_ds_tr_data(app, selected_devices)

    colour_dic = create_colour_dic(devices)

    for device in selected_devices:
        if device.object.GetClassName() == 'ElmRelay':
            vipage = create_plot(app, graphics_board, colour_dic, device, system_volts, f_type='Ground')
            create_draw_format(vipage)
            if device.min_fl_ph > 0:
               vipage = create_plot(app, graphics_board, colour_dic, device, system_volts, f_type='Phase')
        else:
            app.PrintPlain('No relays to plot')

    study_case.Activate()


def new_page_format(app):

    prjt = app.GetActiveProject()
    settings = prjt.GetContents('Settings.SetFold')[0]
    formats = create_obj(settings, "Page Formats", "SetFoldpage")
    new_format_name = '210_x_61'
    new_format = create_obj(formats, new_format_name, "SetFormat")

    return new_format


def drawing_format(graphics_board, format_graph):
    """

    :param graphics_board:
    :param format_graph:
    :param reset:
    :return:
    """

    draw_form = create_obj(graphics_board, "Drawing Format", "SetGrfpage")
    format_graph.iSizeX = 210
    format_graph.iSizeY = 61
    format_graph.iLeft = 0
    format_graph.iRight = 0
    format_graph.iTop = 0
    format_graph.iBottom = 0
    draw_form.iDrwFrm = 1           # 0 = Portrait, 1 = Landscape
    draw_form.aDrwFrm = format_graph.loc_name     # Format


def update_ds_tr_data(app, device_list):
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
            if device.object.GetClassName() == 'ElmRelay':
                max_ds_tr = device.max_ds_tr
                if max_ds_tr is None:
                    continue
                tr_name = max_ds_tr.term.object.cpSubstat.loc_name
                if tr_name[:2] != "SP":
                    if max_ds_tr.load_kva in [1000, 1500]:
                        max_tr_string = tr_name + "_1"
                    else:
                        max_tr_string = tr_name + "_0"
                    tr_strings_dic[device.object] = max_tr_string
        if tr_strings_dic:
            max_tr_strings = list(tr_strings_dic.values())
            results = grf.get_transformer_specifications(max_tr_strings)
            for device_object, string in tr_strings_dic.items():
                device = [device for device in device_list if device.object == device_object][0]
                string_result = results[string]
                max_ds_tr = device.max_ds_tr
                max_ds_tr.insulation = string_result['insulation']
                max_ds_tr.impedance = string_result['impedance']


def create_colour_dic(devices):
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
    colour_dic = {device.object: i+1 for i, device in enumerate(unique_devices)}
    return colour_dic


def create_plot(app, graphics_board, colour_dic, relay, system_volts, f_type: str):

    relay_name = relay.object.loc_name
    app.PrintPlain(relay_name)
    # Create the plot graphic object
    folder_name = f"{relay_name} Coordination Plot"
    vipage = create_obj(graphics_board, folder_name, "SetVipage")
    plot_name = f"{relay_name} {f_type} Coordination Plot"
    plot = vipage.CreateObject("VisOcplot", plot_name)
    plot.Clear()

    # Add the max and min constants to the plot
    min_fl = plot.CreateObject("VisXvalue", f'{relay_name} min fl')
    max_fl = plot.CreateObject("VisXvalue", f'{relay_name} max fl')
    if f_type == 'Ground':
        xvalue_settings(min_fl, 'PG Min FL', relay.min_fl_pg)
        xvalue_settings(max_fl, 'PG Max FL', relay.max_fl_pg)
    else:
        xvalue_settings(min_fl, "Ph Min FL", relay.min_fl_ph)
        xvalue_settings(max_fl, "Ph Max FL", relay.max_fl_ph)

    # Add the ds transformer constants to the plot
    if relay.max_ds_tr is not None:
        tr_term = relay.max_ds_tr.term
        max_ds_tr = tr_term.object.cpSubstat.loc_name
        tr_max_pg = relay.max_ds_tr.max_pg
        tr_max_ph = relay.max_ds_tr.max_ph
        max_tr_fl = plot.CreateObject("VisXvalue", f'{max_ds_tr} max fl')
        if f_type == 'Ground':
            xvalue_settings(max_tr_fl, 'DS TR PG Max FL', tr_max_pg)
        else:
            xvalue_settings(max_tr_fl, "DS TR Ph Max FL",  tr_max_ph)
        ds_fuse = ds.create_fuse(app, relay, system_volts)
        if not ds_fuse:
            app.PrintPlain(
                f'Downstream max transformer fuse size for {max_ds_tr} '
                f'could not be found in PowerFactory'
            )
    else:
        ds_fuse = []
    # Add all the devices to the plot
    us_device = [device.object for device in relay.us_devices]
    ds_device = [device.object for device in relay.ds_devices]

    all_devices = [relay.object] + us_device + ds_device + ds_fuse

    for device in all_devices:
        if device.GetClassName() == 'ElmRelay':
            colour = colour_dic[device]
        else:
            colour = 10      # fuse colour
        plot.AddRelay(device, colour)
    plot_settings(plot, relay, f_type)

    return vipage


def create_draw_format(vipage):

    draw_format = create_obj(vipage, "Drawing Format", "SetGrfpage")
    draw_format.iDrwFrm = 1
    draw_format.aDrwFrm = 'A4'


def plot_settings(plot, relay, f_type):
    """
    Apply plot parameters
    :param plot:
    :param relay:
    :param f_type:
    :return:
    """

    def _get_bound(num, bound):
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
        x_min = _get_bound(relay.min_fl_ph, bound='Min')
        x_max = _get_bound(relay.max_fl_ph, bound='Max')

    setocplt(plot, f_type)

    plot.x_max = x_max
    plot.x_min = x_min
    plot.x_map = 1                      # log scale
    plot.y_max = 10
    plot.y_min = 0.01
    plot.y_map = 1                      # log scale
    plot.y_fit = 0                      # Adapt scale


def setocplt(plot, f_type):
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



def xvalue_settings(constant, name, value):
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


def create_folder(app):

    title = 'Protection Coordination Studies'

    prjt = app.GetActiveProject()
    existing_study_fold = [
        folder.loc_name for folder in prjt.GetContents("*.IntFolder")
    ]
    if title not in existing_study_fold:
        prjt.CreateObject("IntFolder", title)
