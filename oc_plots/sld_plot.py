from dataclasses import dataclass
from typing import List, Dict, Tuple
from devices import devices as ds


def create_slf(app, project_folder, relay, fixpos=False):

    user = app.GetCurrentUser()
    name = relay.object.loc_name
    network_graphic = project_folder.GetContents(f"{name} Graphic.VisGrfnet")
    if fixpos:
        network_graphic[0].SetAttribute("fixPos", 0)
        return
    # template_folder = user.GetContents('templates')
    # if not template_folder:
    #     app.PrintPlain(f"A SLD for {name} was not created as no template file was found in the user directory")
    #     return False
    # diagram = template_folder[0].GetContents('2 Device Radial.IntGrfnet')[0]


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
    if relay.us_device:
        relay_1.loc_name = relay.us_device[0].loc_name
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
    ds_tr_fuse = ds.get_fuse_element(app, relay.ds_tr_size)
    if ds_tr_fuse:
        fuse_1.loc_name = ds_tr_fuse.loc_name
    fuse_1.bus1 = term_7_cub_1
    fuse_1.bus2 = term_2_cub_0

    ds_tr = relay_folder.CreateObject("ElmTr2", f'{name} DS TR')
    ds_tr.loc_name = f"{relay.ds_tr_site} - {relay.ds_tr_pg} pg / {relay.ds_tr_ph} ph"
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

    _create_net_obj(new_diagram, 'NetObj1', 'TermStrip', 24.0625, 170.625, term_1)
    _create_net_obj(new_diagram, 'NetObj2', 'd_lin',  28.4375, 83.125, line_1)
    _create_net_obj(new_diagram, 'NetObj3', 'PointTerm',  39.375, 56.875, term_2)
    _create_net_obj(new_diagram, 'NetObj4', 'PointTerm',  39.375, 48.125, term_3)
    _create_net_obj(new_diagram, 'NetObj5', 'd_fuse',  39.375, 61.25, fuse_1, col=4)
    _create_net_obj(new_diagram, 'NetObj6', 'd_tr2',  39.375, 52.5, ds_tr)
    _create_net_obj(new_diagram, 'NetObj7', 'TermStrip',  17.5, 35., term_4)
    _create_net_obj(new_diagram, 'NetObj8', 'd_lin',  17.5, 59.0625, line_2)
    _create_net_obj(new_diagram, 'NetObj9', 'd_lin',  26.25, 140., line_3)
    _create_net_obj(new_diagram, 'NetObj10', 'd_relay',  32.8125, 161.875, relay_1, col=2)
    _create_net_obj(new_diagram, 'NetObj11', 'd_relay',  32.8125, 118.125, relay_2, col=3)
    _create_net_obj(new_diagram, 'NetObj12', 'd_ct',  26.25, 166.25, ct_1)
    _create_net_obj(new_diagram, 'NetObj13', 'd_ct',  26.25, 113.75, ct_2)
    _create_net_obj(new_diagram, 'NetObj14', 'TermStrip',  24.0625, 109.375, term_5)
    _create_net_obj(new_diagram, 'NetObj15', 'PointTerm',  17.5, 83.125, term_6)
    _create_net_obj(new_diagram, 'NetObj16', 'PointTerm',  39.375, 65.625, term_7)
    _create_net_obj(new_diagram, 'NetObj17', 'd_lin',  24.0625, 170.625, line_4)


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