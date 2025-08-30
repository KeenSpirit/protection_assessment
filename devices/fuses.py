from devices import fuse_mapping as fm
import pf_protection_helper as pph
from importlib import reload
reload(fm)


def create_fuse(app, ds_tr, system_volts):

    if ds_tr is None:
        return []
    typfuse = get_fuse_element(app, ds_tr, system_volts)
    if not typfuse:
        return []
    fuse_name = ds_tr.term.object.cpSubstat.loc_name
    equip = app.GetProjectFolder("equip")
    protection = pph.create_obj(equip, "Protection", "IntFolder")
    fuse_folder = pph.create_obj(protection, "Fuses", "IntFolder")
    lib_contents = fuse_folder.GetContents("*.RelFuse", 0)
    if lib_contents:
        for fuse in lib_contents:
            if fuse.loc_name == fuse_name:
                return [fuse]
    rel_fuse = fuse_folder.CreateObject("RelFuse", f'{fuse_name}')
    rel_fuse.loc_name = fuse_name
    rel_fuse.SetAttribute("typ_id", typfuse)
    return [rel_fuse]


def get_fuse_element(app, tfmr: object, system_volts):
    """
    Match the transformer size to the fuse type using the appropriate fuse mapping dictionary.
    Matching criteria is specified according to TS0013i: RMU Fuse Selection Guide
    :param app:
    :param tfmr:
    :param system_volts:
    :return:
    """

    fuse_object = None
    region = pph.obtain_region(app)
    term = tfmr.term
    if region == 'SEQ':
        if term.constr == "SWER":
            try:
                fuse_string = fm.ex_SWER_f_sizes[tfmr.load_kva]
            except KeyError:
                return None
        elif term.constr == "OH":
            # OH fuse
            if ph_attr_lookup(term.phases) == 1:
                try:
                    fuse_string = fm.ex_pole_1p_fuses[tfmr.load_kva]
                except KeyError:
                    return None
            else:
                try:
                    fuse_string = fm.ex_pole_3p_fuses[tfmr.load_kva]
                except KeyError:
                    return None
            fuse_types = f_types(app, 0)
        else:
            # RMU fuse
            if tfmr.insulation == 'air':
                if tfmr.load_kva < 750:
                    key = tfmr.load_kva
                elif tfmr.impedance == 'high':
                    key = f"{tfmr.load_kva} HighZ"
                else:
                    key = f"{tfmr.load_kva} LowZ"
                try:
                    fuse_string = fm.ex_rmu_air_fuses[key]
                except KeyError:
                    return None
            else:
                try:
                    fuse_string = fm.ex_rmu_oil_fuses[tfmr.load_kva]
                except KeyError:
                    return None
            fuse_types = f_types(app, 1)
    else:    # region == 'Regional Models':
        fuse_types = f_types(app, 0)
        term_volts = round(term.l_l_volts, 2)
        if term.constr == "SWER":
            if tfmr.load_kva >= 100:
                # SWER isolating transformer
                if round(term_volts) == 11 and round(system_volts) == 11:
                    fuse_string = fm.ee_swer_isol_tr_11_11[tfmr.load_kva]
                elif round(term_volts) == 22 and round(system_volts) == 11:
                    fuse_string = fm.ee_swer_isol_tr_11_127[tfmr.load_kva]
                elif round(term_volts) == 33 and round(system_volts) == 11:
                    fuse_string = fm.ee_swer_isol_tr_11_191[tfmr.load_kva]
                elif round(term_volts) == 22 and round(system_volts) == 22:
                    fuse_string = fm.ee_swer_isol_tr_22_127[tfmr.load_kva]
                elif round(term_volts) == 33 and round(system_volts) == 22:
                    fuse_string = fm.ee_swer_isol_tr_22_191[tfmr.load_kva]
                elif round(term_volts) == 22 and round(system_volts) == 33:
                    fuse_string = fm.ee_swer_isol_tr_33_127[tfmr.load_kva]
                else: # term_volts == 33 and system_volts == 33:
                    fuse_string = fm.ee_swer_isol_tr_33_191[tfmr.load_kva]
            else:
                # SWER transformer
                if round(term_volts) == 11:
                    fuse_string = fm.ee_swer_dist_tr_11[tfmr.load_kva]
                elif round(term_volts) == 22:
                    fuse_string = fm.ee_swer_dist_tr_127[tfmr.load_kva]
                else: # 33kV
                    fuse_string = fm.ee_swer_dist_tr_191[tfmr.load_kva]
        else:
            # TR HV fuse
            # Assumed that all TR HV fuses are EDO, because pf doesn't currently contain any OH current limiting fuses
            if round(term_volts) == 11:
                if ph_attr_lookup(term.phases) == 1:
                    fuse_string = fm.ee_tr_11_1p[tfmr.load_kva]
                else:
                    fuse_string = fm.ee_tr_11_3p[tfmr.load_kva]
            elif round(term_volts) == 22:
                if ph_attr_lookup(term.phases) == 1:
                    fuse_string = fm.ee_tr_22_1p[tfmr.load_kva]
                else:
                    fuse_string = fm.ee_tr_22_3p[tfmr.load_kva]
            else:   # 33kV
                if ph_attr_lookup(term.phases) == 1:
                    fuse_string = fm.ee_tr_33_1p[tfmr.load_kva]
                else:
                    fuse_string = fm.ee_tr_33_3p[tfmr.load_kva]

    for fuse in fuse_types:
        if fuse.loc_name == fuse_string:
            fuse_object = fuse
            break
    return fuse_object


def f_types(app, recursive: int):
    """

    :param app:
    :param recursive: 0, 1
    :return:
    """
    # Create a list of all the fuse types
    ergon_lib = app.GetGlobalLibrary()
    fuse_folder = ergon_lib.SearchObject(r"\ErgonLibrary\Protection\Fuses.IntFolder")
    fuse_types = fuse_folder.GetContents("*.TypFuse", recursive)

    # Temporary fuse storage location until RMU fuses have been transferred to master folder
    user = app.GetCurrentUser()
    database = user.fold_id
    try:
        user_fuse_folder = database.SearchObject(f"\\{user}\\_Fuse Models.IntFolder")
        rmu_fuse_types = user_fuse_folder.GetContents("*.TypFuse", 1)
        return fuse_types + rmu_fuse_types
    except Exception:
        return fuse_types


def ph_attr_lookup(attr):
    """
    Convert the terminal phase technology attribute phtech to a meaningful value
    :param attr:
    :return:
    """
    phases = {1:[6, 7, 8], 2:[2, 3, 4, 5], 3:[0, 1]}
    for phase, attr_list in phases.items():
        if attr in attr_list:
            return phase
