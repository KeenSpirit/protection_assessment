from devices import fuse_mapping as fm
import pf_protection_helper as helper
from importlib import reload
import sys
sys.path.append(r"\\Ecasd01\WksMgmt\PowerFactory\ScriptsDEV\PowerFactoryTyping")
import powerfactorytyping as pft
from typing import Union, List
import script_classes as dd
reload(fm)
reload(dd)


def get_all_relays(app: pft.Application) -> List[pft.ElmRelay]:
    net_mod = app.GetProjectFolder("netmod")
    # Filter for relays under network model recursively.
    all_relays = net_mod.GetContents("*.ElmRelay", True)
    relays = [
        relay
        for relay in all_relays
        if relay.cpGrid
        if relay.cpGrid.IsCalcRelevant()
        if relay.GetParent().GetClassName() == "StaCubic"
        if not relay.IsOutOfService()
    ]
    return relays


def get_all_fuses(app: pft.Application) -> List[pft.RelFuse]:
    net_mod = app.GetProjectFolder("netmod")
    all_fuses = net_mod.GetContents("*.RelFuse", True)
    fuses = [
        fuse
        for fuse in all_fuses
        if fuse.cpGrid
        if fuse.cpGrid.IsCalcRelevant()
        if fuse.fold_id.HasAttribute("cterm")
        if fuse.fold_id.cterm.IsEnergized()
        if not fuse.IsOutOfService()
        if determine_fuse_type(fuse)
    ]
    return fuses


def create_fuse(app: pft.Application, ds_tr: Union[None, dd.Tfmr], tr_term_dataclass: dd.Termination, sys_volts: str) -> List[pft.RelFuse]:

    if ds_tr is None:
        return []
    typfuse = get_fuse_element(app, ds_tr, tr_term_dataclass, sys_volts)
    if typfuse is None:
        return []
    fuse_name = ds_tr.term.cpSubstat.loc_name
    equip = app.GetProjectFolder("equip")
    protection = helper.create_obj(equip, "Protection", "IntFolder")
    fuse_folder = helper.create_obj(protection, "Fuses", "IntFolder")
    lib_contents = fuse_folder.GetContents("*.RelFuse", 0)
    if lib_contents:
        for fuse in lib_contents:
            if fuse.loc_name == fuse_name:
                return [fuse]
    rel_fuse = fuse_folder.CreateObject("RelFuse", f'{fuse_name}')
    rel_fuse.loc_name = fuse_name
    rel_fuse.SetAttribute("typ_id", typfuse)
    return [rel_fuse]


def get_fuse_element(app: pft.Application, tfmr: dd.Tfmr, tr_term_dataclass: dd.Termination, system_volts: str) -> Union[None, pft.TypFuse]:
    """
    Match the transformer size to the fuse type using the appropriate fuse mapping dictionary.
    Matching criteria is specified according to TS0013i: RMU Fuse Selection Guide
    :param app:
    :param tfmr:
    :param tr_term_dataclass:
    :param system_volts:
    :return:
    """

    def _safe_string(dic, key):
        try:
            fuse_string = dic[key]
        except KeyError:
            fuse_string = None
        return fuse_string

    fuse_object = None
    region = helper.obtain_region(app)
    term = tr_term_dataclass
    if region == 'SEQ':
        if term.constr == dd.ConstructionType.SWER.value:
            fuse_string = _safe_string(fm.ex_SWER_f_sizes, tfmr.load_kva)
        elif term.constr == dd.ConstructionType.OVERHEAD.value:
            # OH fuse
            if term.phases == 1:
                fuse_string = _safe_string(fm.ex_pole_1p_fuses, tfmr.load_kva)
            else:
                fuse_string = _safe_string(fm.ex_pole_3p_fuses, tfmr.load_kva)
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
                fuse_string = _safe_string(fm.ex_rmu_air_fuses, key)
            else:
                fuse_string = _safe_string(fm.ex_rmu_oil_fuses, tfmr.load_kva)
            fuse_types = f_types(app, 1)
    else:    # region == 'Regional Models':
        fuse_types = f_types(app, 0)
        term_volts = round(term.l_l_volts, 2)
        if term.constr == dd.ConstructionType.SWER.value:
            if tfmr.load_kva >= 100:
                # SWER isolating transformer
                if round(term_volts) == 11 and round(system_volts) == 11:
                    fuse_string = _safe_string(fm.ee_swer_isol_tr_11_11, tfmr.load_kva)
                elif round(term_volts) == 22 and round(system_volts) == 11:
                    fuse_string = _safe_string(fm.ee_swer_isol_tr_11_127, tfmr.load_kva)
                elif round(term_volts) == 33 and round(system_volts) == 11:
                    fuse_string = _safe_string(fm.ee_swer_isol_tr_11_191, tfmr.load_kva)
                elif round(term_volts) == 22 and round(system_volts) == 22:
                    fuse_string = _safe_string(fm.ee_swer_isol_tr_22_127, tfmr.load_kva)
                elif round(term_volts) == 33 and round(system_volts) == 22:
                    fuse_string = _safe_string(fm.ee_swer_isol_tr_22_191, tfmr.load_kva)
                elif round(term_volts) == 22 and round(system_volts) == 33:
                    fuse_string = _safe_string(fm.ee_swer_isol_tr_33_127, tfmr.load_kva)
                else: # term_volts == 33 and system_volts == 33:
                    fuse_string = _safe_string(fm.ee_swer_isol_tr_33_191, tfmr.load_kva)
            else:
                # SWER transformer
                if round(term_volts) == 11:
                    fuse_string = _safe_string(fm.ee_swer_dist_tr_11, tfmr.load_kva)
                elif round(term_volts) == 22:
                    fuse_string = _safe_string(fm.ee_swer_dist_tr_127, tfmr.load_kva)
                else: # 33kV
                    fuse_string = _safe_string(fm.ee_swer_dist_tr_191, tfmr.load_kva)
        else:
            # TR HV fuse
            # Assumed that all TR HV fuses are EDO, because pf doesn't currently contain any OH current limiting fuses
            if round(term_volts) == 11:
                if term.phases == 1:
                    fuse_string = _safe_string(fm.ee_tr_11_1p, tfmr.load_kva)
                else:
                    fuse_string = _safe_string(fm.ee_tr_11_3p, tfmr.load_kva)
            elif round(term_volts) == 22:
                if term.phases == 1:
                    fuse_string = _safe_string(fm.ee_tr_22_1p, tfmr.load_kva)
                else:
                    fuse_string = _safe_string(fm.ee_tr_22_3p, tfmr.load_kva)
            else:   # 33kV
                if term.phases == 1:
                    fuse_string = _safe_string(fm.ee_tr_33_1p, tfmr.load_kva)
                else:
                    fuse_string = _safe_string(fm.ee_tr_33_3p, tfmr.load_kva)
    for fuse in fuse_types:
        if fuse.loc_name == fuse_string:
            fuse_object = fuse
            break
    return fuse_object


def f_types(app: pft.Application, recursive: int):
    """

    :param app:
    :param recursive: 0, 1
    :return:
    """
    # Create a list of all the fuse types
    ergon_lib = app.GetGlobalLibrary()
    fuse_folder = ergon_lib.SearchObject(r"\ErgonLibrary\Protection\Fuses.IntFolder")
    fuse_types = fuse_folder.GetContents("*.TypFuse", recursive)
    return fuse_types


def determine_fuse_type(fuse: pft.RelFuse) -> bool:
    """This function will observe the fuse location and determine if it is
    a Distribution transformer fuse, SWER isolating fuse or a line fuse"""
    # First check is that if the fuse exists in a terminal that is in the
    # System Overiew then it will be a line fuse.
    fuse_active = fuse.HasAttribute("r:fold_id:r:obj_id:e:loc_name")
    if not fuse_active:
        return True
    fuse_grid = fuse.cpGrid
    if (
            fuse.GetAttribute("r:fold_id:r:cterm:r:fold_id:e:loc_name")
            == fuse_grid.loc_name
    ):
        # This would indicate it is in a line cubical
        return True
    if fuse.loc_name not in fuse.GetAttribute("r:fold_id:r:obj_id:e:loc_name"):
        # This indicates that the fuse is not in a switch object
        return True
    secondary_sub = fuse.fold_id.cterm.fold_id
    contents = secondary_sub.GetContents()
    for content in contents:
        if content.GetClassName() == "ElmTr2":
            return False
    else:
        return True