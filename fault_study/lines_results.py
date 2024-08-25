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
