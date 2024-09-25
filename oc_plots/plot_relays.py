import math
from devices import devices as ds
from oc_plots import sld_plot as sld
from pf_protection_helper import create_obj
from importlib import reload
reload(sld)


def plot_all_relays(app, device_list):

    app.PrintPlain("Generating device time overcurrent plots...")
    new_format = new_page_format(app)
    study_case = app.GetActiveStudyCase()
    graphics_board = app.GetFromStudyCase("Graphics Board.SetDesktop")
    study_case.Deactivate()
    drawing_format(graphics_board, new_format)

    for device in device_list:
        if device.object.GetClassName() == 'ElmRelay':
            vipage = create_plot(app, graphics_board, device, f_type='Ground')
            create_draw_format(vipage)
            if device.min_fl_ph > 0:
               vipage = create_plot(app, graphics_board, device, f_type='Phase')
            # Add the single line diagram
            # new_diagram = sld.create_slf(vipage, device)
            # if new_diagram:
            #     sld.update_slf(app, new_diagram, device)
            #     study_case.Activate()
            #     sld.layers(new_diagram)
            #     study_case.Deactivate()
            # vipage.Close()
        else:
            app.PrintPlain('No relays to plot')

    study_case.Activate()

        # Save results to study folder
        # prjt = app.GetActiveProject()
        # study_folder = prjt.GetContents("Protection Coordination Studies.IntFolder")[0]
        # study_folder.AddCopy(vipage, folder_name)
        # vipage.Delete()


def create_plot(app, graphics_board, relay, f_type: str):

    relay_name = relay.object.loc_name
    app.PrintPlain(relay_name)
    # Create the plot graphic object
    folder_name = f"{relay_name} Coordination Plot"
    vipage = create_obj(graphics_board, folder_name, "SetVipage")
    plot_name = f"{relay_name} {f_type} Coordination Plot"
    plot = vipage.CreateObject("VisOcplot", plot_name)
    plot.Clear()

    # Add the constants to the plot
    min_fl = plot.CreateObject("VisXvalue", f'{relay_name} min fl')
    max_fl = plot.CreateObject("VisXvalue", f'{relay_name} max fl')
    max_tr_fl = plot.CreateObject("VisXvalue", f'{relay.max_ds_tr} max fl')

    if f_type == 'Ground':
        xvalue_settings(min_fl, 'PG Min FL', relay.min_fl_pg)
        xvalue_settings(max_fl, 'PG Max FL', relay.max_fl_pg)
        xvalue_settings(max_tr_fl, 'DS TR PG Max FL', relay.tr_max_pg)
    else:
        xvalue_settings(min_fl, "Ph Min FL", relay.min_fl_ph)
        xvalue_settings(max_fl, "Ph Max FL", relay.max_fl_ph)
        xvalue_settings(max_tr_fl, "DS TR Ph Max FL", relay.tr_max_ph)

    # Add all the devices to the plot
    us_device = [device.object for device in relay.us_devices]
    ds_fuse = ds.create_fuse(app, relay)
    if not ds_fuse:
        app.PrintPlain(
            f'Downstream max transformer fuse size for {relay.max_ds_tr} could not be found in PowerFactory'
        )
    all_devices = [relay.object] + us_device + ds_fuse

    plot.AddRelays(all_devices)
    plot_settings(plot, relay, f_type)

    return vipage


def create_draw_format(vipage):

    draw_format = create_obj(vipage, "Drawing Format", "SetGrfpage")
    draw_format.iDrwFrm = 1
    draw_format.aDrwFrm = 'A4'


def plot_settings(plot, relay, f_type):

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
    plot.y_max = 5
    plot.y_min = 0.01
    plot.y_map = 0                      # linear scale
    plot.y_fit = 0                      # Adapt scale


def setocplt(plot, f_type):

    settings = create_obj(plot, "Overcurrent Plot Settings", "SetOcplt")
    if f_type == 'Ground':
        settings.ishow = 3
    else:
        settings.ishow = 1
    settings.iTbrk = 0


def xvalue_settings(constant, name, value):

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


def grid_settings(plot):

    grid_set = plot.IntPlot
    grid_set.XGridMain = 1
    grid_set.XGridHlp = 1
    grid_set.YGridMain = 1
    grid_set.YGridMHlp = 1
    grid_set.mxpos = 2
    grid_set.txpos = 2
    grid_set.mypos = 2
    grid_set.typos = 2


def create_folder(app):

    title = 'Protection Coordination Studies'

    prjt = app.GetActiveProject()
    existing_study_fold = [
        folder.loc_name for folder in prjt.GetContents("*.IntFolder")
    ]
    if title not in existing_study_fold:
        prjt.CreateObject("IntFolder", title)

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


def plot_fix(app, user_selection):

    relay_list = []
    for feeders, devices in user_selection.items():
        for device in devices:
            if device.GetClassName() == 'ElmRelay':
                relay_list.append(device.loc_name)
    study_case = app.GetActiveStudyCase()
    graphics_board = app.GetFromStudyCase("Graphics Board.SetDesktop")
    study_case.Deactivate()
    for relay in relay_list:
        folder_name = f"{relay} Coordination Plot"
        vipage = graphics_board.GetContents(f"{folder_name}.SetVipage")[0]
        draw_form = vipage.GetContents("Drawing Format.SetGrfpage")[0]
        draw_form.aDrwFrm = "A4"
        network_graphic = vipage.GetContents(f"{relay} Graphic.VisGrfnet")[0]
        network_graphic.fixPos = 1


def test_func(app):

    graphics_board = app.GetFromStudyCase("Graphics Board.SetDesktop")
    folder_name = "RC-1396404 Coordination Plot"
    vipage = graphics_board.GetContents(f"{folder_name}.SetVipage")[0]
    network_graphic = vipage.GetContents("RC-1396404 Graphic.VisGrfnet")[0]
    two_dev_rad = network_graphic.GetContents("2 Device Radial.IntGrfnet")[0]
    grf_net_obj = two_dev_rad.GetContents("Graphical Net Object9.IntGrf")[0]
    grf_con_1 = grf_net_obj.GetContents("GCO_1.IntGrfcon")[0]
    app.PrintPlain(f"{grf_con_1.rX}")
    app.PrintPlain(f"{grf_con_1.rY}")
    grf_con_2 = grf_net_obj.GetContents("GCO_2.IntGrfcon")[0]
    app.PrintPlain(f"{grf_con_2.rX}")
    app.PrintPlain(f"{grf_con_2.rY}")











