from typing import Union
from fault_study import study_templates


def short_circuit(app, bound: str, f_type: str, location: Union[object, None] = None, ppro: int = 0) -> object:
    """
    Set the Short-circuit command module and perform a short-circuit calculation
    :param app:
    :param bound: 'Max', 'Min'
    :param f_type: '3 Phase', '2 Phase', 'Ground'
    :param location: element location of fault. None if All Busbars
    :param ppro: fault distance from terminal
    :return: Short-Circuit Command
    """

    ComShc = app.GetFromStudyCase("Short_Circuit.ComShc")
    study_templates.apply_sc(ComShc, bound, f_type)
    ComShc.SetAttribute("e:iopt_prot", 0)
    if location:
        ComShc.SetAttribute("e:iopt_allbus", 0)
        ComShc.SetAttribute("e:shcobj", location)
        ComShc.SetAttribute("e:iopt_dfr", 0)
        ComShc.SetAttribute("e:ppro", ppro)
    else:
        ComShc.SetAttribute("e:iopt_allbus", 1)

    return ComShc.Execute()


def get_line_current(elmlne: object) -> float:
    if elmlne.HasAttribute('bus1'):
        Ia1 = elmlne.GetAttribute('m:Ikss:bus1:A')
        Ib1 = elmlne.GetAttribute('m:Ikss:bus1:B')
        Ic1 = elmlne.GetAttribute('m:Ikss:bus1:C')
    if elmlne.HasAttribute('bus2'):
        Ia2 = elmlne.GetAttribute('m:Ikss:bus2:A')
        Ib2 = elmlne.GetAttribute('m:Ikss:bus2:B')
        Ic2 = elmlne.GetAttribute('m:Ikss:bus2:C')

    if elmlne.HasAttribute('bus1') and elmlne.HasAttribute('bus2'):
        return round(max(Ia1, Ib1, Ic1, Ia2, Ib2, Ic2))
    elif elmlne.bus1:
        return round(max(Ia1, Ib1, Ic1))
    elif elmlne.bus2:
        return round(max(Ia2, Ib2, Ic2))
    else:
        return None
