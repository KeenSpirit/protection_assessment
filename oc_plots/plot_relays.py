import math
from devices import devices as ds

def plot_all_relays(app, device_list):

    app.PrintPlain("Generating device time overcurrent plots")
    study_case = app.GetActiveStudyCase()
    graphics_board = app.GetFromStudyCase("Graphics Board.SetDesktop")
    drawing_format(graphics_board, "210_x_61")
    study_case.Deactivate()

    for device in device_list:
        if device.object.GetClassName() == 'ElmRelay':
            vipage = create_plot(app, graphics_board, device, f_type='Ground')
            if device.min_fl_ph > 0:
               vipage = create_plot(app, graphics_board, device, f_type='Phase')
            # Add the single line diagram
            new_diagram = create_slf(app, vipage, device)
            if new_diagram:
                update_slf(app, new_diagram, device)
                # graphic = vipage.GetContents(f"{device.object.loc_name} Graphic.VisGrfnet")[0]
                # graphic.SetAttribute("fixPos", 1)
        else:
            app.PrintPlain('No relays to plot')
    # drawing_format(graphics_board, "A4")
    app.Rebuild(2)

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


def create_slf(app, project_folder, relay):

    user = app.GetCurrentUser()
    name = relay.object.loc_name
    template_folder = user.GetContents('templates')[0]
    diagram = template_folder.GetContents('2 Device Radial.IntGrfnet')[0]

    network_graphic = project_folder.GetContents(f"{name} Graphic.VisGrfnet")
    if not network_graphic:
        network_graphic = project_folder.CreateObject("VisGrfnet", f'{name} Graphic')
        new_diagram = network_graphic.AddCopy(diagram)
        network_graphic.SetAttribute("sglGrf", new_diagram)
        network_graphic.SetAttribute("allowInteraction", 1)
        network_graphic.SetAttribute("FrmVis", 1)
        network_graphic.SetAttribute("fixPos", 0)
    else:
        new_diagram = False
    return new_diagram


def update_slf(app, new_diagram, relay):


    # Create a folder to store new net dat objects
    name = relay.object.loc_name
    network_data = app.GetProjectFolder("netdat")
    new_elements = network_data.GetContents("New Elements.ElmNet", True)[0]

    title = f'{name} Graphic Elements'
    relay_folder = new_elements.GetContents(f"{title}.IntFolder")
    if relay_folder: [folder.Delete() for folder in relay_folder]
    new_elements.CreateObject("IntFolder", title)
    relay_folder = new_elements.GetContents(f"{title}.IntFolder")[0]

    # Create new Net Dat objects
    term_1 = relay_folder.CreateObject("ElmTerm", f'{name} term 1')
    term_1_cub_0 = term_1.CreateObject("StaCubic", 'Cub_0')
    ct_1 = term_1_cub_0.CreateObject("StaCt", f'{name} CT 1')
    relay_1 = term_1_cub_0.CreateObject("ElmRelay", f'Upstream Relay')
    relay_1.pdiselm = [ct_1, None]
    cb_1 = term_1_cub_0.CreateObject("StaSwitch", f'{name} cb 1')
    cb_1.aUsage = 'cbk'
    cb_1.on_off = 1

    term_2 = relay_folder.CreateObject("ElmTerm", f'{name} term 2')
    term_2_cub_0 = term_2.CreateObject("StaCubic", 'Cub_0')
    term_2_cub_1 = term_2.CreateObject("StaCubic", 'Cub_1')

    term_3 = relay_folder.CreateObject("ElmTerm", f'{name} term 3')
    term_3_cub_0 = term_3.CreateObject("StaCubic", 'Cub_0')

    term_4 = relay_folder.CreateObject("ElmTerm", f'{name} term 4')
    term_4_cub_0 = term_4.CreateObject("StaCubic", 'Cub_0')

    term_5 = relay_folder.CreateObject("ElmTerm", f'{name} term 5')
    term_5_cub_0 = term_5.CreateObject("StaCubic", 'Cub_0')
    ct_2 = term_5_cub_0.CreateObject("StaCt", f'{name} CT 2')
    relay_2 = term_5_cub_0.CreateObject("ElmRelay", name)
    relay_2.pdiselm = [ct_2, None]
    cb_2 = term_5_cub_0.CreateObject("StaSwitch", f'{name} cb 2')
    cb_2.aUsage = 'cbk'
    cb_2.on_off = 1
    term_5_cub_1 = term_5.CreateObject("StaCubic", 'Cub_1')

    term_6 = relay_folder.CreateObject("ElmTerm", f'{name} term 6')
    term_6_cub_0 = term_6.CreateObject("StaCubic", 'Cub_0')
    term_6_cub_1 = term_6.CreateObject("StaCubic", 'Cub_1')
    term_6_cub_2 = term_6.CreateObject("StaCubic", 'Cub_2')

    term_7 = relay_folder.CreateObject("ElmTerm", f'{name} term 7')
    term_7_cub_0 = term_7.CreateObject("StaCubic", 'Cub_0')
    term_7_cub_1 = term_7.CreateObject("StaCubic", 'Cub_1')

    # Create and connect lines
    line_1 = relay_folder.CreateObject("ElmLne", f'{name} line 1')
    line_1.bus1 = term_6_cub_0
    line_1.bus2 = term_7_cub_0

    line_2 = relay_folder.CreateObject("ElmLne", f'{name} line 2')
    line_2.bus1 = term_6_cub_1
    line_2.bus2 = term_4_cub_0

    line_3 = relay_folder.CreateObject("ElmLne", f'{name} line 3')
    line_3.bus1 = term_1_cub_0
    line_3.bus2 = term_5_cub_0

    line_4 = relay_folder.CreateObject("ElmLne", f'{name} line 4')
    line_4.bus1 = term_5_cub_1
    line_4.bus2 = term_6_cub_2

    fuse_1 = relay_folder.CreateObject("RelFuse", f'{name} fuse 1')
    fuse_1.bus1 = term_7_cub_1
    fuse_1.bus2 = term_2_cub_0

    ds_tr = relay_folder.CreateObject("ElmTr2", f'{name} DS TR')
    ds_tr.bushv = term_2_cub_1
    ds_tr.buslv = term_3_cub_0


    # Get graphic objects and assign them to the new net dat objects
    objects = new_diagram.GetContents('*.IntGrf')

    go_term_1 = [obj for obj in objects if obj.loc_name == 'Graphical Net Object1'][0]
    go_term_1.pDataObj = term_1
    go_line_1 = [obj for obj in objects if obj.loc_name == 'Graphical Net Object10'][0]
    go_line_1.pDataObj = line_1
    go_term_2 = [obj for obj in objects if obj.loc_name == 'Graphical Net Object12'][0]
    go_term_2.pDataObj = term_2
    go_term_3 = [obj for obj in objects if obj.loc_name == 'Graphical Net Object14'][0]
    go_term_3.pDataObj = term_3

    go_tr_fuse = [obj for obj in objects if obj.loc_name == 'Graphical Net Object16'][0]
    go_tr_fuse.iCol = 4                         # Blue
    go_tr_fuse.pDataObj = fuse_1
    ds_tr_fuse = ds.get_fuse_element(app, relay.ds_tr_size)
    fuse_1.loc_name = ds_tr_fuse.loc_name

    go_ds_tr = [obj for obj in objects if obj.loc_name == 'Graphical Net Object18'][0]
    go_ds_tr.pDataObj = ds_tr
    ds_tr.loc_name = f"{relay.ds_tr_site} - {relay.ds_tr_pg} pg / {relay.ds_tr_ph} ph"

    go_term_4 = [obj for obj in objects if obj.loc_name == 'Graphical Net Object18(1)'][0]
    go_term_4.pDataObj = term_4
    term_4.loc_name = f"Sect FL Min - {relay.min_fl_pg} pg / {relay.min_fl_ph} ph"

    go_line_2 = [obj for obj in objects if obj.loc_name == 'Graphical Net Object19'][0]
    go_line_2.pDataObj = line_2
    go_line_3 = [obj for obj in objects if obj.loc_name == 'Graphical Net Object20'][0]
    go_line_3.pDataObj = line_3

    go_relay_1 = [obj for obj in objects if obj.loc_name == 'Graphical Net Object20(1)'][0]
    go_relay_1.iCol = 2                         # Red
    go_relay_1.pDataObj = relay_1
    if relay.us_device:
        relay_1.loc_name = relay.us_device[0].loc_name

    go_relay_2 = [obj for obj in objects if obj.loc_name == 'Graphical Net Object21'][0]
    go_relay_2.iCol = 3                         # Green
    go_relay_2.pDataObj = relay_2
    relay_2.loc_name = relay.object.loc_name

    go_ct_1 = [obj for obj in objects if obj.loc_name == 'Graphical Net Object22'][0]
    go_ct_1.pDataObj = ct_1
    go_ct_2 = [obj for obj in objects if obj.loc_name == 'Graphical Net Object23'][0]
    go_ct_2.pDataObj = ct_2

    go_term_5 = [obj for obj in objects if obj.loc_name == 'Graphical Net Object4'][0]
    go_term_5.pDataObj = term_5
    term_5.loc_name = f"Sect FL Max - {relay.max_fl_pg} pg / {relay.max_fl_ph} ph"

    go_term_6 = [obj for obj in objects if obj.loc_name == 'Graphical Net Object5'][0]
    go_term_6.pDataObj = term_6
    go_term_7 = [obj for obj in objects if obj.loc_name == 'Graphical Net Object7'][0]
    go_term_7.pDataObj = term_7
    go_line_4 = [obj for obj in objects if obj.loc_name == 'Graphical Net Object9'][0]
    go_line_4.pDataObj = line_4



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
    constant.lab_text = name          # line label
    constant.show = 1                   # line with intersections
    constant.label = 0                  # line with intersections
    constant.iopt_lab= 3                # position
    constant.value= value
    constant.color = 1                  # line with intersections
    #constant.style = 1
    constant.width = 0.4
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


def drawing_format(graphics_board, format):
    """

    :param graphics_board:
    :param format: "A4", "210_x_61"
    :return:
    """


    draw_form = graphics_board.GetContents("Drawing Format.SetGrfpage")[0]

    draw_form.iDrwFrm = 1           # 0 = Portrait, 1 = Landscape
    draw_form.aDrwFrm = format        # Format

