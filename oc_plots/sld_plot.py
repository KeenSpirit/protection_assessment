from dataclasses import dataclass
from typing import List, Dict, Tuple
from devices import devices as ds


def create_slf(project_folder, relay, fixpos=False):


    name = relay.object.loc_name
    network_graphic = project_folder.GetContents(f"{name} Graphic.VisGrfnet")
    if fixpos:
        draw_form = project_folder.GetContents("Drawing Format.SetGrfpage")[0]
        draw_form.aDrwFrm = "A4"
        # network_graphic = project_folder.GetContents("*.VisGrfnet")[0]
        network_graphic[0].fixPos = 1
        network_graphic[0].SetAttribute("fixPos", 0)
        return

    if not network_graphic:
        network_graphic = project_folder.CreateObject("VisGrfnet", f'{name} Graphic')
        new_diagram = network_graphic.CreateObject("IntGrfnet", '2 Device Radial')
        network_graphic.SetAttribute("sglGrf", new_diagram)
        network_graphic.SetAttribute("allowInteraction", 1)
        network_graphic.SetAttribute("FrmVis", 1)
        network_graphic.SetAttribute("fixPos", 0)
        set_fold = create_setting_fold(new_diagram)
        create_format(set_fold)
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
    if relay.us_devices:
        relay_1.loc_name = [relay.object.loc_name for relay in relay.us_devices][0]
    cb_1 = term_1_cub_0.CreateObject("StaSwitch", f'{name} cb 1')
    cb_1.aUsage = 'cbk'
    cb_1.on_off = 1

    term_2 = relay_folder.CreateObject("ElmTerm", f'{name} term 2')
    term_2_cub_0 = term_2.CreateObject("StaCubic", 'Cub_0')
    term_2_cub_1 = term_2.CreateObject("StaCubic", 'Cub_1')

    term_3 = relay_folder.CreateObject("ElmTerm", f'{name} term 3')
    term_3_cub_0 = term_3.CreateObject("StaCubic", 'Cub_0')

    term_4 = relay_folder.CreateObject("ElmTerm", f'{name} term 4')
    term_4.loc_name = f"Sect FL Min - {relay.min_fl_pg} pg / {relay.min_fl_ph} ph"
    term_4_cub_0 = term_4.CreateObject("StaCubic", 'Cub_0')

    term_5 = relay_folder.CreateObject("ElmTerm", f'{name} term 5')
    term_5_cub_0 = term_5.CreateObject("StaCubic", 'Cub_0')
    term_5.loc_name = f"Sect FL Max - {relay.max_fl_pg} pg / {relay.max_fl_ph} ph"
    ct_2 = term_5_cub_0.CreateObject("StaCt", f'{name} CT 2')
    relay_2 = term_5_cub_0.CreateObject("ElmRelay", name)
    relay_2.pdiselm = [ct_2, None]
    relay_2.loc_name = relay.object.loc_name
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
    ds_tr_fuse = ds.get_fuse_element(app, relay.max_tr_size)
    if ds_tr_fuse:
        fuse_1.loc_name = ds_tr_fuse.loc_name
    fuse_1.bus1 = term_7_cub_1
    fuse_1.bus2 = term_2_cub_0

    ds_tr = relay_folder.CreateObject("ElmTr2", f'{name} DS TR')
    ds_tr.loc_name = f"{relay.max_ds_tr} - {relay.tr_max_pg} pg / {relay.tr_max_ph} ph"
    ds_tr.bushv = term_2_cub_1
    ds_tr.buslv = term_3_cub_0


    # Get graphic objects and assign them to the new net dat objects
    def _create_net_obj(int_grf, name, sym_nam, x, y, data_obj, col=1):

        obj_lookup = {
             'TermStrip': 1,
             'PointTerm': 1,
             'd_lin': 1,
             'd_tr2': 1,
             'd_relay': 35,
             'd_ct': 36,
             'd_fuse': 1
            }

        obj = int_grf.CreateObject("IntGrf", name)
        obj.sSymNam = sym_nam
        obj.rCenterX = x
        obj.rCenterY = y
        obj.iLevel = obj_lookup[sym_nam]
        obj.pDataObj = data_obj
        obj.iCol = col
        return obj

    int_grf = _create_net_obj(new_diagram, 'NetObj1', 'TermStrip', 24.0625, 170.625, term_1)
    _create_labels(app, int_grf, -0.6562, 0.2625, 1, 1, 7, 3)
    int_grf = _create_net_obj(new_diagram, 'NetObj2', 'd_lin',  28.4375, 83.125, line_1)
    _create_labels(app, int_grf, -0.6562, 0.2625, 0, 1, 7, 3)
    _create_grfcon(int_grf, 'GCO_1', 0, 28.4375, 17.5, -1.0, 83.125, 83.125, -1.0)
    _create_grfcon(int_grf, 'GCO_2', 1, 28.4375, 39.375, 39.375, 83.125, 83.125, 65.625)
    int_grf = _create_net_obj(new_diagram, 'NetObj3', 'PointTerm',  39.375, 56.875, term_2)
    _create_labels(app, int_grf, -0.6562, 0.2625, 0, 1, 7, 3)
    int_grf = _create_net_obj(new_diagram, 'NetObj4', 'PointTerm',  39.375, 48.125, term_3)
    _create_labels(app, int_grf, -0.6562, 0.2625, 0, 1, 7, 3)
    int_grf = _create_net_obj(new_diagram, 'NetObj5', 'd_fuse',  39.375, 61.25, fuse_1, col=4)
    _create_labels(app, int_grf, 10.6527, 1.18824, 1, 1, 5, 7)
    int_grf = _create_net_obj(new_diagram, 'NetObj6', 'd_tr2',  39.375, 52.5, ds_tr)
    _create_labels(app, int_grf, 2.49169, -13.070, 1, 1, 5, 7)
    int_grf = _create_net_obj(new_diagram, 'NetObj7', 'TermStrip',  17.5, 35., term_4)
    _create_labels(app, int_grf, 32.2185, -4.0943, 1, 2, 5, 1)
    int_grf = _create_net_obj(new_diagram, 'NetObj8', 'd_lin',  17.5, 59.0625, line_2)
    _create_labels(app, int_grf, 32.2185, -4.0943, 0, 1, 7, 3)
    _create_grfcon(int_grf, 'GCO_1', 0, 17.5, 17.5, -1.0, 59.0625, 83.125, -1.0)
    _create_grfcon(int_grf, 'GCO_2', 1, 17.5, 17.5, -1.0, 59.0625, 35.0, -1.0)
    int_grf = _create_net_obj(new_diagram, 'NetObj9', 'd_lin',  26.25, 140., line_3)
    _create_labels(app, int_grf, 32.2185, -4.0943, 0, 1, 7, 3)
    _create_grfcon(int_grf, 'GCO_1', 0, 26.25, 26.25, -1.0, 140.0, 170.625, -1.0)
    _create_grfcon(int_grf, 'GCO_2', 1, 26.25, 26.25, -1.0, 140.0, 109.375, -1.0)
    int_grf = _create_net_obj(new_diagram, 'NetObj10', 'd_relay',  32.8125, 161.875, relay_1, col=2)
    _create_labels(app, int_grf, 0, -1, 1, 1, 7, 3)
    int_grf = _create_net_obj(new_diagram, 'NetObj11', 'd_relay',  32.8125, 118.125, relay_2, col=3)
    _create_labels(app, int_grf, 0, -1, 1, 1, 7, 3)
    int_grf = _create_net_obj(new_diagram, 'NetObj12', 'd_ct',  26.25, 166.25, ct_1)
    _create_grfcon(int_grf, 'GCO_1', 0, 26.25, 26.25, -1.0, 166.25, 170.625, -1.0)
    _create_grfcon(int_grf, 'GCO_2', 1, 28.4375, 32.8125, 32.8125, 166.25, 166, 162.875)
    int_grf = _create_net_obj(new_diagram, 'NetObj13', 'd_ct',  26.25, 113.75, ct_2)
    _create_grfcon(int_grf, 'GCO_1', 0, 26.25, 26.25, -1.0, 113.74, 109.375, -1.0)
    _create_grfcon(int_grf, 'GCO_2', 1, 28.4375, 32.8125, -32.8125, 113.75, 113.75, 117.125)
    int_grf = _create_net_obj(new_diagram, 'NetObj14', 'TermStrip',  24.0625, 109.375, term_5)
    _create_labels(app, int_grf, 44.2989, -3.8963, 1, 2, 5, 1)
    int_grf = _create_net_obj(new_diagram, 'NetObj15', 'PointTerm',  17.5, 83.125, term_6)
    _create_labels(app, int_grf, 44.2989, -3.8963, 0, 2, 5, 1)
    int_grf = _create_net_obj(new_diagram, 'NetObj16', 'PointTerm',  39.375, 65.625, term_7)
    _create_labels(app, int_grf, 44.2989, -3.8963, 0, 2, 5, 1)
    int_grf = _create_net_obj(new_diagram, 'NetObj17', 'd_lin',  17.5, 96.25, line_4)
    _create_labels(app, int_grf, 44.2989, -3.8963, 0, 2, 5, 1)
    _create_grfcon(int_grf, 'GCO_1', 0,17.5, 17.5, -1.0, 96.25, 109.375, -1.0)
    _create_grfcon(int_grf, 'GCO_2', 1, 17.5, 17.5, -1.0, 96.25, 83.125, -1.0)


def create_setting_fold(intgrfnet):

    set_fold = intgrfnet.GetContents("Settings.IntFolder", True)
    if not set_fold:
        set_fold = intgrfnet.CreateObject("IntFolder", "Settings")
    else:
        set_fold = set_fold[0]
    return set_fold


def create_format(set_fold):

    format = set_fold.GetContents("Format.SetGrfpage", True)
    if not format:
        format = set_fold.CreateObject("SetGrfpage", "Format")
    else:
        format = format[0]
    format.iDrwFrm = 0
    format.aDrwFrm = '210_x_61'


def _create_grfcon(int_grf, name, connr, a, b, c, d, e, f):

    obj = int_grf.CreateObject("IntGrfcon", name)

    obj.iDatConNr = connr
    obj.iLinSt = 1
    obj.rLinWd = 0
    obj.rX = [a, b, c, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0]
    obj.rY = [d, e, f, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0]


def _create_labels(app, int_grf, x, y, vis, orient, par, box):

    global_lib = app.GetGlobalLibrary()
    form = global_lib.SearchObject(r"\System\Standard\Settings\Formats\Graphic\Label.IntFormsel")

    setvitxt = int_grf.CreateObject("SetVitxt", "LabelTermStrip")
    setvitxt.format = form
    setvitxt.ifont = 11
    setvitxt.center_x = x
    setvitxt.center_y = y
    setvitxt.do_rot = 3
    setvitxt.zlvl = vis
    setvitxt.chars = 31
    setvitxt.glvl = 2
    setvitxt.iOrient = orient
    setvitxt.iLines = 1
    setvitxt.anr = 0
    setvitxt.vlpos = 0
    setvitxt.iParRefPt = par
    setvitxt.iBoxRefPt = box


def layers(int_grf):

    sites = int_grf.SearchObject(r"\Layers\Bays and Sites.IntGrflayer")
    sites.iVis = 1
    cts = int_grf.SearchObject(r"\Layers\Current/voltage transformers.IntGrflayer")
    cts.iVis = 1
    sites = int_grf.SearchObject(r"\Layers\Diagram layer.IntGrflayer")
    sites.iVis = 1
    sites = int_grf.SearchObject(r"\Layers\Labels.IntGrflayer")
    sites.iVis = 1
    sites = int_grf.SearchObject(r"\Layers\Net elements.IntGrflayer")
    sites.iVis = 1
    sites = int_grf.SearchObject(r"\Layers\Relays.IntGrflayer")
    sites.iVis = 1




























# @dataclass
# class Node:
#     name: str
#     parent: str
#     children: List[str]
#
#
# def generate_hierarchy(nodes: List[Node]) -> Dict[str, List[int]]:
#     positions = {}
#     root = next(node for node in nodes if node.parent == "")
#
#     def place_node(node: Node, x: int, y: int):
#         positions[node.name] = [x, y]
#         child_nodes = [n for n in nodes if n.name in node.children]
#         num_children = len(child_nodes)
#
#         if num_children == 0:
#             return
#
#         total_width = max(5 * (num_children - 1), 0)
#         start_x = x - total_width // 2
#
#         for i, child in enumerate(child_nodes):
#             child_x = start_x + i * 5
#             child_y = y + 5
#             place_node(child, child_x, child_y)
#
#     place_node(root, 0, 0)
#     return positions
#
#
# # Example usage:
# if __name__ == "__main__":
#     nodes = [
#         Node("A", "", ["B", "C", "D"]),
#         Node("B", "A", ["E", "F"]),
#         Node("C", "A", []),
#         Node("D", "A", ["G"]),
#         Node("E", "B", []),
#         Node("F", "B", []),
#         Node("G", "D", [])
#     ]
#
#     result = generate_hierarchy(nodes)
#     for node, position in result.items():
#         print(f"Node {node}: position {position}")