import sys
sys.path.append(r"\\Ecasd01\WksMgmt\PowerFactory\ScriptsDEV\PowerFactoryTyping")
import powerfactorytyping as pft


def get_conductor(line):
    """Looks at the tower geometry or the cable system to return the conductor type and thermal rating
    """
    construction = line.typ_id.GetClassName()

    if construction == 'TypGeo':
        TypCon = line.GetAttribute("e:pCondCir")
        conductor_type = TypCon.loc_name
        thermal_rating = round(TypCon.GetAttribute("e:Ithr"), 3) * 1000
    elif construction == 'TypCabsys':
        conductor_type = 'NA'
        thermal_rating = 'NA'
    elif construction == "TypLne":
        conductor_type = line.typ_id.loc_name
        thermal_rating = round(line.typ_id.GetAttribute("e:Ithr"), 3) * 1000
    elif construction == "TypTow":
        conductor_type = 'NA'
        thermal_rating = 'NA'
    else:
        conductor_type = 'NA'
        thermal_rating = 'NA'
    return conductor_type, thermal_rating


def get_phases(line) -> int:
    """Looks at the tower geometry or the cable system to return the number
    of phases.
    """
    construction = line.typ_id.GetClassName()

    if construction == 'TypGeo':
        TypGeo = line.typ_id
        num_phases = TypGeo.xy_c[0][0]
    elif construction == 'TypCabsys':
        num_phases = line.typ_id.GetAttribute('nphas')[0]
    elif construction == "TypLne":
        num_phases = line.typ_id.GetAttribute('nlnph')
    else:
        raise TypeError(f'{construction} Unhandelled construction')

    return int(num_phases)


def get_voltage(line):
    """
    Get the line-line operating voltage of the given ElmLne element.
    :param line:
    :return:
    """

    terms = line.GetConnectedElements()
    l_l_volts = 0
    for term in terms:
        try:
            l_l_volts = term.uknom
            break
        except AttributeError:
            pass
    return l_l_volts



