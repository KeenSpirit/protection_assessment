
def get_floating_terminals(feeder, devices) -> dict[object:dict[object:object]]:
    """
    Outputs all floating terminal objects with their associated line objects for all devices
    :param feeder:
    :param devices:
    :return:
    """

    floating_terms= {}
    floating_lines = find_end_points(feeder)
    for device in devices:
        terms = [term.object for term in device.sect_terms]
        floating_terms[device.object] = {}
        for line in floating_lines:
            try:
                t1, t2 = line.GetConnectedElements()
            except:
                continue
            t3 = line.GetConnectedElements(1,1,0)
            if len(t3) == 1 and t3[0] == t2 and t2 in terms and t1 not in terms:
                floating_terms[device.object][line] = t1
            elif len(t3) == 1 and t3[0] == t1 and t1 in terms and t2 not in terms:
                floating_terms[device.object][line] = t2

    return floating_terms


def find_end_points(feeder: object) -> list[object]:
    """
    Returns a list of sections with only one connection (i.e. end points).
    :param feeder: The feeder being investigated.
    :return: A list of line sections that have only one connection (i.e. end points).
    """

    floating_lines = []

    # Get all the sections that make up the selected feeder.
    feeder_lines = feeder.GetObjs('ElmLne')

    for ElmLne in feeder_lines:
        if ElmLne.bus1:
            bus1 = [x.obj_id for x in ElmLne.bus1.cterm.GetConnectedCubicles()
                    if x is not ElmLne.bus1
                    if x.obj_id.GetClassName() == 'ElmLne']
        else:
            bus1 = []

        if ElmLne.bus2:
            if ElmLne.bus2.HasAttribute('cterm'):
                bus2 = [x.obj_id for x in ElmLne.bus2.cterm.GetConnectedCubicles()
                        if x is not ElmLne.bus2
                        if x.obj_id.GetClassName() == 'ElmLne']
            else:
                bus2 = []
        else:
            bus2 = []

        if len(bus1) == 1 or len(bus2) == 1 \
                or (len(bus1) > 1 and ElmLne not in bus1) \
                or (len(bus2) > 1 and ElmLne not in bus2):
            floating_lines.append(ElmLne)

    return floating_lines
