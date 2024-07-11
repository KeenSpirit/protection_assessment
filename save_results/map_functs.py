from typing import Union
import sys
sys.path.append(r"\\Ecasd01\WksMgmt\PowerFactory\ScriptsDEV\PowerFactoryTyping")
import powerfactorytyping as pft


def site_name_convert(site_names: list[Union[pft.ElmCoup, pft.StaSwitch, pft.ElmFeeder, pft.ElmRelay, pft.RelFuse]]) \
        -> dict[Union[pft.ElmCoup, pft.StaSwitch, pft.ElmFeeder]:dict[pft.StaCubic: pft.ElmTerm]]:
    """
    Convert a site name to a PowerFactory object by matching it with the equvalent switch/breaker name.
    Convert the switch/breaker to a StaCubic and ElmTerm object.
    :param app:
    :param site_names:
    :return:
    """

    site_name_map = {}
    for name in site_names:
        cubicle, device_term = obj_to_term(name)
        site_name_map[name] = {cubicle: device_term}

    return site_name_map


def obj_to_term(switch: Union[pft.ElmCoup, pft.StaSwitch, pft.ElmFeeder, pft.ElmRelay, pft.RelFuse]) -> tuple[pft.StaCubic, pft.ElmTerm]:
    """
    Convert a powrfactory object to an equivalment StaCubic and ElmTerm object
    """
    if switch.GetClassName() == "ElmCoup":
        if switch.HasAttribute('bus1'):
            cubicle = switch.bus1
            device_term_1 = cubicle.cterm
            if device_term_1.iUsage == 1:
                return cubicle, device_term_1
        if switch.HasAttribute('bus2'):
            cubicle = switch.bus2
            device_term_2 = cubicle.cterm
            return cubicle, device_term_2
    elif switch.GetClassName() in ["StaSwitch", "ElmRelay", "RelFuse"]:
        cubicle = switch.fold_id
        device_term = cubicle.cterm
        return cubicle, device_term
    else:                                           # ElmFeeder
        cubicle = switch.obj_id
        device_term = cubicle.cterm
        return cubicle, device_term


def term_element(term: pft.ElmTerm, site_name_map, element=False) -> Union[pft.ElmCoup, pft.StaSwitch, pft.ElmFeeder]:
    """
    Map an ElmTerm to it's corresponding ElmCoup/StaSwitch/ElmFeeder
    :param term:
    :param site_name_map:
    :param element:
    :return:
    """

    for device, value in site_name_map.items():
        (elmterm,) = value.values()
        if element and term == elmterm:
            return device
        elif term == elmterm:
            return device