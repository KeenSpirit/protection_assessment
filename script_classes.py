"""
All element data used by the script is stored in dataclasses.
There are four dataclass types used by the script:
- devices
- terminations
- lines
- transformers
This module is used for defining and initiliasing each dataclass
It also defines various element types used by the script
"""

from enum import Enum
from dataclasses import dataclass
from devices import fuses as ds
from typing import Optional, Dict


class ElementType(Enum):
    TERM = "ElmTerm"
    LOAD = "ElmLod"
    TFMR = "ElmTr2"
    LINE = "ElmLne"
    GRID = "ElmXnet"
    COUPLER = "ElmCoup"
    SWITCH ="StaSwitch"
    RELAY = 'ElmRelay'
    FUSE = 'RelFuse'
    FEEDER = 'ElmFeeder'


@dataclass
class Feeder:
    obj: object
    cubicle: object
    term: object
    sys_volts: float
    devices: list
    bu_devices: list
    open_points: Dict


@dataclass
class Device:
    obj: object
    cubicle: object
    term: object
    phases: int
    l_l_volts: float
    ds_capacity: Optional[float]
    max_fl_3ph: Optional[float]
    max_fl_2ph: Optional[float]
    max_fl_pg: Optional[float]
    min_fl_3ph: Optional[float]
    min_fl_2ph: Optional[float]
    min_fl_pg: Optional[float]
    min_sn_fl_2ph: Optional[float]
    min_sn_fl_pg: Optional[float]
    max_ds_tr: Optional[float]
    sect_terms: list
    sect_loads: list
    sect_lines: list
    us_devices: list
    ds_devices: list


@dataclass
class Termination:
    obj: object
    constr: Optional[str]
    phases: int
    l_l_volts: float
    max_fl_3ph: float
    max_fl_2ph: float
    max_fl_pg: float
    min_fl_3ph: float
    min_fl_2ph: float
    min_fl_pg: float
    min_fl_pg10: float
    min_fl_pg50: float
    min_sn_fl_2ph: float
    min_sn_fl_pg: float
    min_sn_fl_pg10: float
    min_sn_fl_pg50: float


@dataclass
class Line:
    obj: object
    phases: int
    l_l_volts: float
    max_fl_3ph: float
    max_fl_2ph: float
    max_fl_pg: float
    min_fl_3ph: float
    min_fl_2ph: float
    min_fl_pg: float
    min_sn_fl_2ph: float
    min_sn_fl_pg: float
    line_type: str
    thermal_rating: float
    ph_energy: float
    ph_clear_time: float
    ph_fl: float
    pg_energy: float
    pg_clear_time: float
    pg_fl: float


@dataclass
class Tfmr:
    obj: object
    term: object
    load_kva: Optional[float]
    max_ph: Optional[str]
    max_pg: Optional[str]
    fuse: Optional[str]
    insulation: Optional[str]
    impedance: Optional[str]


def initialise_fdr_dataclass(element):

    dataclass = Feeder(
            element,
            element.obj_id,
            element.cn_bus,
            element.cn_bus.uknom,
            [],
            [],
            []
            )
    return dataclass


def initialise_dev_dataclass(element):

    if element is None:
        return None
    if element.GetClassName() == ElementType.FEEDER.value:
        cubicle = element.obj_id
    elif element.GetClassName() == ElementType.COUPLER.value:
        cubicle = element.bus1
    else:
        cubicle = element.fold_id

    dataclass = Device(
            element,
            cubicle,
            cubicle.cterm,
            ds.ph_attr_lookup(cubicle.cterm.phtech),
            round(cubicle.cterm.uknom, 2),
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            [],
            [],
            [],
            [],
            []
            )
    return dataclass


def initialise_term_dataclass(elmterm):

    if elmterm is None:
        return None
    dataclass = Termination(
            elmterm,
            None,
            ds.ph_attr_lookup(elmterm.phtech),
            round(elmterm.uknom,2),
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None
            )
    return dataclass


def initialise_line_dataclass(elmlne):

    if elmlne is None:
        return None

    def _get_phases(elmlne) -> int:
        """Looks at the tower geometry or the cable system to return the number
        of phases.
        """
        construction = elmlne.typ_id.GetClassName()

        if construction == 'TypGeo':
            TypGeo = elmlne.typ_id
            num_phases = TypGeo.xy_c[0][0]
        elif construction == 'TypCabsys':
            num_phases = elmlne.typ_id.GetAttribute('nphas')[0]
        elif construction == "TypLne":
            num_phases = elmlne.typ_id.GetAttribute('nlnph')
        else:
            raise TypeError(f'{construction} Unhandelled construction')

        return int(num_phases)

    def _get_voltage(elmlne):
        """
        Get the line-line operating voltage of the given ElmLne element.
        :param line:
        :return:
        """

        terms = elmlne.GetConnectedElements()
        l_l_volts = 0
        for term in terms:
            try:
                l_l_volts = term.uknom
                break
            except AttributeError:
                pass
        return l_l_volts

    def _get_conductor(elmlne):
        """Looks at the tower geometry or the cable system to return the conductor type and thermal rating
        """
        construction = elmlne.typ_id.GetClassName()

        if construction == 'TypGeo':
            TypCon = elmlne.GetAttribute("e:pCondCir")
            conductor_type = TypCon.loc_name
            thermal_rating = round(TypCon.GetAttribute("e:Ithr"), 3) * 1000
        elif construction == 'TypCabsys':
            conductor_type = 'NA'
            thermal_rating = 'NA'
        elif construction == "TypLne":
            conductor_type = elmlne.typ_id.loc_name
            thermal_rating = round(elmlne.typ_id.GetAttribute("e:Ithr"), 3) * 1000
        elif construction == "TypTow":
            conductor_type = 'NA'
            thermal_rating = 'NA'
        else:
            conductor_type = 'NA'
            thermal_rating = 'NA'
        return conductor_type, thermal_rating

    phases = _get_phases(elmlne)
    voltage = _get_voltage(elmlne)
    line_type, line_therm_rating = _get_conductor(elmlne)

    dataclass = Line(
            elmlne,
            phases,
            voltage,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            line_type,
            line_therm_rating,
            None,
            None,
            None,
            None,
            None,
            None)
    return dataclass


def initialise_load_dataclass(load):
    if load is None:
        return Tfmr(None, None, None, None, None, None, None, None)
    if load.GetClassName() == ElementType.LOAD.value:
        return Tfmr(load, load.bus1.cterm, round(load.Strat), None, None, None, None, None)
    if load.GetClassName() == ElementType.TFMR.value:
        return Tfmr(load, load.bushv.cterm, round(load.Snom_a * 1000), None, None, None, None, None)