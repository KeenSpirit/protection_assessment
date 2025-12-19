from typing import Union
import sys
from pf_config import pft
from fault_study import study_templates


def short_circuit(app, bound: str, f_type: str, location: Union[pft.ElmTerm, None] = None, ppro: int = 0) -> pft.ComShc:
    """
    Set the Short-circuit command module and perform a short-circuit calculation
    :param app:
    :param bound: 'Max', 'Min'
    :param f_type: '3-Phase', '2-Phase', "Phase-Ground"
    :param location: element location of fault. None if All Busbars
    :param ppro: fault distance from terminal
    :return: Short-Circuit Command
    """

    ComShc = app.GetFromStudyCase("Short_Circuit.ComShc")
    study_templates.apply_sc(ComShc, bound, f_type)
    if location:
        ComShc.SetAttribute("e:iopt_allbus", 0)
        ComShc.SetAttribute("e:shcobj", location)
        ComShc.SetAttribute("e:iopt_dfr", 0)
        ComShc.SetAttribute("e:ppro", ppro)
    else:
        ComShc.SetAttribute("e:iopt_allbus", 1)

    return ComShc.Execute()


def get_terminal_current(elmterm: pft.ElmTerm) -> float:

    def _check_att(obj, attribute):
        if obj.HasAttribute(attribute):
            terminal_fl = round(obj.GetAttribute(attribute) * 1000)
        else:
            terminal_fl = 0
        return terminal_fl

    Ia = _check_att(elmterm, 'm:Ikss:A')
    Ib = _check_att(elmterm, 'm:Ikss:B')
    Ic = _check_att(elmterm, 'm:Ikss:C')

    return max(Ia, Ib, Ic)


def get_line_current(elmlne: pft.ElmLne) -> float:
    currents = []

    if elmlne.HasAttribute('bus1'):
        currents.extend([
            elmlne.GetAttribute('m:Ikss:bus1:A'),
            elmlne.GetAttribute('m:Ikss:bus1:B'),
            elmlne.GetAttribute('m:Ikss:bus1:C')
        ])

    if elmlne.HasAttribute('bus2'):
        currents.extend([
            elmlne.GetAttribute('m:Ikss:bus2:A'),
            elmlne.GetAttribute('m:Ikss:bus2:B'),
            elmlne.GetAttribute('m:Ikss:bus2:C')
        ])
    return round(max(currents) * 1000) if currents else None


