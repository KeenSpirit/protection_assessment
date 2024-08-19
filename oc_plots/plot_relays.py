import math
from devices import devices as ds
from oc_plots import sld_plot as sld
from importlib import reload
reload(sld)

def plot_all_relays(app, device_list):

    app.PrintPlain("Generating device time overcurrent plots")
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
            new_diagram = sld.create_slf(app, vipage, device)
            if new_diagram:
                sld.update_slf(app, new_diagram, device)
            vipage.Close()
            # sld.create_slf(app, vipage, device, fixpos=True)
        else:
            app.PrintPlain('No relays to plot')
    # drawing_format(graphics_board, new_format, reset=True)

    study_case.Activate()

        # Save results to study folder
        # prjt = app.GetActiveProject()
        # study_folder = prjt.GetContents("Protection Coordination Studies.IntFolder")[0]
        # study_folder.AddCopy(vipage, folder_name)
        # vipage.Delete()

def create_plot(app, graphics_board, relay, f_type: str):


    relay_name = relay.object.loc_name
    # Create the plot graphic object
    folder_name = f"{relay_name} Coordination Plot"
    vipage = graphics_board.GetContents(f"{folder_name}.SetVipage")
    if not vipage:
        vipage = graphics_board.CreateObject("SetVipage", folder_name)
    else:
        vipage = vipage[0]
    plot_name = f"{relay_name} {f_type} Coordination Plot"
    plot = vipage.CreateObject("VisOcplot", plot_name)
    plot.Clear()
    plot_settings(plot, relay, f_type)

    # Add the constants to the plot
    min_fl = plot.CreateObject("VisXvalue", f'{relay_name} min fl')
    max_fl = plot.CreateObject("VisXvalue", f'{relay_name} max fl')
    max_tr_fl = plot.CreateObject("VisXvalue", f'{relay.ds_tr_site} max fl')

    if f_type == 'Ground':
        xvalue_settings(min_fl, 'PG Min FL', relay.min_fl_pg)
        xvalue_settings(max_fl, 'PG Max FL', relay.max_fl_pg)
        xvalue_settings(max_tr_fl, 'DS TR PG Max FL', relay.ds_tr_pg)
    else:
        xvalue_settings(min_fl, "Ph Min FL", relay.min_fl_ph)
        xvalue_settings(max_fl, "Ph Max FL", relay.max_fl_ph)
        xvalue_settings(max_tr_fl, "DS TR Ph Max FL", relay.ds_tr_ph)

    # Add all the devices to the plot
    us_device = relay.us_device
    ds_fuse = ds.create_fuse(app, relay)
    if not ds_fuse:
        app.PrintPlain(
            f'Downstream max transformer fuse size for {relay.ds_tr_site} could not be found in PowerFactory'
        )
    all_devices = [relay.object] + us_device + ds_fuse

    plot.AddRelays(all_devices)

    return vipage


def create_draw_format(vipage):

    draw_format = vipage.GetContents("Drawing Format.SetGrfpage")
    if not draw_format:
        draw_format = vipage.CreateObject("SetGrfpage", "Drawing Format")
    else:
        draw_format = vipage[0]
    draw_format.iDrwFrm = 1
    draw_format.aDrwFrm = '210_x_61'


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
    else:
        x_min = _get_bound(relay.min_fl_ph, bound='Min')
        x_max = _get_bound(relay.max_fl_ph, bound='Max')

    plot.x_max = x_max
    plot.x_min = x_min
    plot.x_map = 1                      # log scale
    plot.y_max = 10
    plot.y_min = 0.01
    plot.y_map = 0                      # linear scale
    plot.y_fit = 0                      # Adapt scale


def xvalue_settings(constant, name, value):

    constant.loc_name = name            # name
    constant.label = 1
    constant.lab_text = [name]          # line label
    constant.show = 1                   # line with intersections
    constant.label = 0                  # line with intersections
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
    formats = settings.GetContents("Page Formats.SetFoldpage")
    if not formats:
        formats = settings.CreateObject("SetFoldpage", "Page Formats")
    else:
        formats = formats[0]
    new_format_name = '210_x_61'
    new_format = formats.GetContents(f"{new_format_name}.SetFormat")
    if not new_format:
        new_format = formats.CreateObject("SetFormat", new_format_name)
    else:
        return new_format[0]
    return new_format


def drawing_format(graphics_board, format_graph, reset=False):
    """

    :param graphics_board:
    :param format_graph:
    :param reset:
    :return:
    """

    draw_form_list = graphics_board.GetContents("Drawing Format.SetGrfpage")
    if not draw_form_list:
        draw_form = graphics_board.CreateObject("SetGrfpage", "Drawing Format")
    else:
        draw_form = draw_form_list[0]

    if reset:
        draw_form.aDrwFrm = 'A4'            # Format
        return

    format_graph.iSizeX = 210
    format_graph.iSizeY = 61
    format_graph.iLeft = 0
    format_graph.iRight = 0
    format_graph.iTop = 0
    format_graph.iBottom = 0

    draw_form.iDrwFrm = 1           # 0 = Portrait, 1 = Landscape
    draw_form.aDrwFrm = format_graph.loc_name     # Format


# Deactivate study case
# Under the SetVipage, get the "Drawing Format.SetGrfpage" object
# change object.aDrwFrm = "210_x_61"
# Under the SetVipage object, get the VisGrfnet object
# set VisGrfnet.fixPos = 0


def plot_fix(app):

    study_case = app.GetActiveStudyCase()
    graphics_board = app.GetFromStudyCase("Graphics Board.SetDesktop")
    study_case.Deactivate()
    folder_name = "CHTOSS-FB52-J01 Coordination Plot"
    vipage = graphics_board.GetContents(f"{folder_name}.SetVipage")[0]
    draw_form = vipage.GetContents("Drawing Format.SetGrfpage")[0]
    draw_form.aDrwFrm = "A4"
    network_graphic = vipage.GetContents("CHTOSS-FB52-J01 Graphic.VisGrfnet")[0]
    network_graphic.fixPos = 1