import sys
sys.path.append(r"\\Ecasd01\WksMgmt\PowerFactory\ScriptsDEV\PowerFactoryTyping")
from dataclasses import dataclass
import powerfactorytyping as pft


def lines(app, ppp_max_all, pp_max_all, pg_max_all, ppp_min_all, pp_min_all, pg_min_all, sys_norm_min_pp_all,
          sys_norm_min_pg_all):
    """For each line in the feeder, obtain its maximum and minimum fault levels"""

    line_sections = get_line_sections(app)

    lines_max_3p = {fdr: {} for fdr in ppp_max_all}
    lines_max_2p = {fdr: {} for fdr in ppp_max_all}
    lines_max_pg = {fdr: {} for fdr in ppp_max_all}
    lines_min_3p = {fdr: {} for fdr in ppp_max_all}
    lines_min_2p = {fdr: {} for fdr in ppp_max_all}
    lines_min_pg = {fdr: {} for fdr in ppp_max_all}
    lines_sys_norm_min_2p = {fdr: {} for fdr in ppp_max_all}
    lines_sys_norm_min_pg = {fdr: {} for fdr in ppp_max_all}
    lines_type = {fdr: {} for fdr in ppp_max_all}
    lines_therm_rating = {fdr: {} for fdr in ppp_max_all}
    for feeders, feeder in ppp_max_all.items():
        for sections, section in feeder.items():
            lines_max_3p[feeders][sections] = {}
            lines_max_2p[feeders][sections] = {}
            lines_max_pg[feeders][sections] = {}
            lines_min_3p[feeders][sections] = {}
            lines_min_2p[feeders][sections] = {}
            lines_min_pg[feeders][sections] = {}
            lines_sys_norm_min_2p[feeders][sections] = {}
            lines_sys_norm_min_pg[feeders][sections] = {}
            lines_type[feeders][sections] = {}
            lines_therm_rating[feeders][sections] = {}
            terminal_list = []
            for terminal in section:
                terminal_list.append(terminal)
            for line, line_terms in line_sections.items():
                check = all(item in terminal_list for item in line_terms)
                if check is True:
                    max_3p, max_2p, max_pg = 0, 0, 0
                    min_3p, min_2p, min_pg, sys_norm_min_2p, sys_norm_min_pg = 99999, 99999, 99999, 99999, 99999
                    for term in line_terms:
                        if ppp_max_all[feeders][sections][term] >= max_3p:
                            max_3p = section[term]
                        if pp_max_all[feeders][sections][term] >= max_2p:
                            max_2p = pp_max_all[feeders][sections][term]
                        if pg_max_all[feeders][sections][term] >= max_pg:
                            max_pg = pg_max_all[feeders][sections][term]
                        if ppp_min_all[feeders][sections][term] <= min_3p:
                            min_3p = ppp_min_all[feeders][sections][term]
                        if pp_min_all[feeders][sections][term] <= min_2p:
                            min_2p = pp_min_all[feeders][sections][term]
                        if pg_min_all[feeders][sections][term] <= min_pg:
                            min_pg = pg_min_all[feeders][sections][term]
                        if sys_norm_min_pp_all[feeders][sections][term] <= sys_norm_min_2p:
                            sys_norm_min_2p = sys_norm_min_pp_all[feeders][sections][term]
                        if sys_norm_min_pg_all[feeders][sections][term] <= sys_norm_min_pg:
                            sys_norm_min_pg = sys_norm_min_pg_all[feeders][sections][term]
                    lines_max_3p[feeders][sections][line] = max_3p
                    lines_max_2p[feeders][sections][line] = max_2p
                    lines_max_pg[feeders][sections][line] = max_pg
                    lines_min_3p[feeders][sections][line] = min_3p
                    lines_min_2p[feeders][sections][line] = min_2p
                    lines_min_pg[feeders][sections][line] = min_pg
                    lines_sys_norm_min_2p[feeders][sections][line] = sys_norm_min_2p
                    lines_sys_norm_min_pg[feeders][sections][line] = sys_norm_min_pg
                    lines_type[feeders][sections][line], lines_therm_rating[feeders][sections][line]  \
                        = get_conductor(line)

    return (lines_max_3p, lines_max_2p, lines_max_pg, lines_min_3p, lines_min_2p, lines_min_pg,
            lines_sys_norm_min_2p, lines_sys_norm_min_pg, lines_type, lines_therm_rating)


def regional_lines(app, device_lines, ppp_max_all, pg_max_all, pp_min_all, pg_min_all):
    """
    For each line in the feeder, obtain its maximum and minimum fault levels
    THIS FUNCTION CURRENTLY DOES NOT WORK. TO BE FIXED AT A LATER DATE.
    :param app:
    :param ppp_max_all:
    :param pg_max_all:
    :param pp_min_all:
    :param pg_min_all:
    :return:
    """

    all_grids = app.GetCalcRelevantObjects('*.ElmXnet')
    grids = [grid for grid in all_grids if grid.outserv == 0]

    lines_fls = {section: [] for section in ppp_max_all}
    for section, terminals in ppp_max_all.items():
        terminal_list = [key for key in terminals.keys()]
        lines = device_lines[section]
        sect_lines = []
        for line in lines:
            cub_1 = line.bus1
            ds_1 = cub_1.GetAll(1, 0)
            if any(item in grids for item in ds_1):
                ds_2 = cub_1.GetAll(0, 0)
                ds_terms = [object for object in ds_2 if object.GetClassName() == "ElmTerm" and object.uknom > 1]
            else:
                ds_terms = [object for object in ds_1 if object.GetClassName() == "ElmTerm" and object.uknom > 1]
            line_sect_terms = [element for element in ds_terms if element in terminal_list]
            min_2p, min_pg, = 99999, 99999
            for term in line_sect_terms:
                if 1 < pp_min_all[section][term] < min_2p:
                    min_2p = pp_min_all[section][term]
                if 1 < pg_min_all[section][term] < min_pg:
                    min_pg = pg_min_all[section][term]

            terminals = [element for element in line.GetConnectedElements()]
            check = any(item in terminal_list for item in terminals)
            if check is True:
                max_3p, max_pg = 0, 0
                for term in terminals:
                    if term in terminal_list:
                        if ppp_max_all[section][term] >= max_3p:
                            max_3p = ppp_max_all[section][term]
                        if pg_max_all[section][term] >= max_pg:
                            max_pg = pg_max_all[section][term]
                line_type, line_therm_rating = get_conductor(line)

                line = Line(
                    line,
                    round(max_3p),
                    round(max_pg),
                    round(min_2p),
                    round(min_pg),
                    line_type,
                    line_therm_rating,
                    None,
                    None,
                    None,
                    None
                )
                sect_lines.append(line)
        lines_fls[section] = sect_lines


    return lines_fls


def get_line_sections(app) -> dict[pft.ElmLne: list[str]]:
    """

    :param app:
    :return:
    """

    all_active_lines = []
    line_sections = {}
    for grid in app.GetSummaryGrid().GetContents():
        all_active_lines += [
            line
            for line in grid.obj_id.GetContents("*.ElmLne")
            if "HV" in line.loc_name or "TR" in line.loc_name or "LN" in line.loc_name
            if not line.IsOutOfService()
            if line.IsEnergized()
    ]
    for line in all_active_lines:
        terminals = [element.loc_name for element in line.GetConnectedElements()]
        line_sections[line] = terminals

    return line_sections


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

@dataclass
class Line:
    object: object
    max_fl_ph: float
    max_fl_pg: float
    min_fl_ph: float
    min_fl_pg: float
    line_type: str
    thermal_rating: float
    ph_clear_time: float
    ph_fl: float
    pg_clear_time: float
    pg_fl: float
