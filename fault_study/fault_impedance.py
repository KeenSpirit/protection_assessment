import sys
from typing import List, Union
sys.path.append(r"\\Ecasd01\WksMgmt\PowerFactory\ScriptsDEV\PowerFactoryTyping")
import powerfactorytyping as pft
import script_classes as dd


def update_node_construction(devices: List[dd.Device]):

    all_nodes = get_all_terms(devices)
    update_construction(all_nodes)


def get_all_terms(devices: List[dd.Device]) -> List[dd.Termination]:

    all_nodes = []
    for device in devices:
        terms = device.sect_terms
        all_nodes.extend(terms)

    return all_nodes


def update_construction(all_nodes: List[dd.Termination]):

    for node in all_nodes:
        if node.constr is not None:
            continue
        # Get all lines connected to the node
        line_elements = [ele for ele in node.obj.GetConnectedElements() if ele.GetClassName() == dd.ElementType.LINE.value]
        # Sometimes the upstream connection is not a line (can be elmcoup)
        if not line_elements:
            try:
                substation = node.cpSubstat
                proxy_node = substation.pBusbar
                line_elements = [
                    ele for ele in proxy_node.GetConnectedElements() if ele.GetClassName() == dd.ElementType.LINE.value
                ]
            except (AttributeError, IndexError):
                line_elements = []
        # For all upstream lines, determine whether they are overhead or underground construction
        for line in line_elements:
            try:
                line_type = line.typ_id
                if 'SWER' in line_type.loc_name:
                    node.constr = "SWER"
                    break
            except AttributeError:
                pass
            if line.IsCable() and node.constr != "OH":
                node.constr = "UG"
            else:
                node.constr = "OH"
        if node.constr is None:
            node.constr = "OH"


def term_pg_fl(region: str, term: dd.Termination) -> float:
    """
    Determine the correct terminal minimum phase-ground fault current based on the region and connected line
    construction
    :param region:
    :param term:
    :return:
    """

    if region == 'SEQ':
        fault_level = term.min_fl_pg
    else:
        if term.constr == 'OH':
            fault_level = term.min_fl_pg50
        else:
            fault_level = term.min_fl_pg10
    return fault_level

def term_sn_pg_fl(region: str, term: dd.Termination) -> float:
    """
    Determine the correct terminal minimum phase-ground fault current based on the region and connected line
    construction
    :param region:
    :param term:
    :return:
    """

    if region == 'SEQ':
        fault_level = term.min_sn_fl_pg
    else:
        if term.constr == 'OH':
            fault_level = term.min_sn_fl_pg50
        else:
            fault_level = term.min_sn_fl_pg10
    return fault_level