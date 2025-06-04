

def update_node_construction(devices):




    all_nodes = get_all_terms(devices)
    update_construction(all_nodes)


def get_all_terms(devices):

    all_nodes = []
    for device in devices:
        terms = device.sect_terms
        all_nodes.extend(terms)

    return all_nodes


def update_construction(all_nodes):

    for node in all_nodes:
        if node.constr is not None:
            continue
        upstream_lines = []
        # Get all lines connected to the node
        #TODO: Sometimes the upstream connection is not a line (can be elmcoup)
        # Need to keep going upstream until we find a line
        line_elements = [ele for ele in node.object.GetConnectedElements() if ele.GetClassName() == 'ElmLne']
        for ElmLne in line_elements:
            # For each connected line, determine, if it's an upstream line or a downstream line
            cub_1 = ElmLne.bus1
            cub_2 = ElmLne.bus2
            if cub_1 and cub_2:
                term_1 = cub_1.cterm
                term_2 = cub_2.cterm
                if node.object == term_1:
                    remote_node = term_2
                else:
                    remote_node = term_1
                try:
                    remote_node_fl = [node.min_fl_pg for node in all_nodes if node.object == remote_node][0]
                except IndexError:
                    # Probably a substation bus node
                    remote_node_fl = 0
                if node.min_fl_pg < remote_node_fl:
                    # This is an upstream line
                    upstream_lines.append(ElmLne)
        # For all upstream lines, determine whether they are overhead or underground construction
        for line in upstream_lines:
            if not line.IsCable():
                node.constr = "OH"
        if node.constr is None:
            node.constr = "OH"


def term_pg_fl(region, term):
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

    




