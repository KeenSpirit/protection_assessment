import sys
from typing import List, Dict, Union
from pf_config import pft
import domain as dd


def get_floating_terminals(feeder: dd.Feeder, devices: List[dd.Device]) -> Dict:
    """
    Outputs all floating terminal objects with their associated line objects for all devices
    :param feeder:
    :param devices:
    :return:
    """

    floating_terms= {}
    floating_lines = find_end_points(feeder)
    for device in devices:
        terms = [term.obj for term in device.sect_terms]
        floating_terms[device.term] = {}
        for line in floating_lines:
            try:
                t1, t2 = line.GetConnectedElements()
            except AttributeError:
                continue
            t3 = line.GetConnectedElements(1,1,0)
            if len(t3) == 1 and t3[0] == t2 and t2 in terms and t1 not in terms:
                floating_terms[device.term][line] = t1
            elif len(t3) == 1 and t3[0] == t1 and t1 in terms and t2 not in terms:
                floating_terms[device.term][line] = t2

    return floating_terms


def find_end_points(feeder: pft.ElmFeeder) -> List[pft.ElmLne]:
    """
    Returns a list of sections with only one connection (i.e. end points).
    :param feeder: The feeder being investigated.
    :return: A list of line sections that have only one connection (i.e. end points).
    """

    floating_lines = []

    # Get all the sections that make up the selected feeder.
    feeder_lines = feeder.GetObjs('ElmLne')

    for elmlne in feeder_lines:
        if (elmlne.GetAttribute('bus1') is not None
                and elmlne.bus1.GetAttribute('cterm') is not None):
            bus1 = [x.GetAttribute('obj_id') for x in elmlne.bus1.cterm.GetConnectedCubicles()
                    if x is not elmlne.GetAttribute('bus1')
                    if x.obj_id.GetClassName() == dd.ElementType.LINE.value]
        else:
            bus1 = []
        if (elmlne.GetAttribute('bus2') is not None
                and elmlne.bus2.GetAttribute('cterm') is not None):
            bus2 = [x.GetAttribute('obj_id') for x in elmlne.bus2.cterm.GetConnectedCubicles()
                    if x is not elmlne.GetAttribute('bus2')
                    if x.obj_id.GetClassName() == dd.ElementType.LINE.value]
        else:
            bus2 = []

        if len(bus1) == 1 or len(bus2) == 1 \
                or (len(bus1) > 1 and elmlne not in bus1) \
                or (len(bus2) > 1 and elmlne not in bus2):
            floating_lines.append(elmlne)

    return floating_lines

