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
import sys
sys.path.append(r"\\Ecasd01\WksMgmt\PowerFactory\ScriptsDEV\PowerFactoryTyping")
import powerfactorytyping as pft
from typing import List, Dict, Union, Any, Optional
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


class ConstructionType(Enum):
    """Line construction types affecting fault impedance."""
    OVERHEAD = "OH"
    UNDERGROUND = "UG"
    SWER = "SWER"


class FaultType(Enum):
    """Types of faults for analysis."""
    THREE_PHASE = "3-Phase"
    TWO_PHASE = "2-Phase"
    PHASE_GROUND = "Phase-Ground"


@dataclass(frozen=True)
class FaultCurrents:
    """
    Immutable container for fault current values at a location.

    Once created, values cannot be modified, preventing accidental corruption.
    All values in Amperes (primary).

    Usage:
        fc = FaultCurrents(three_phase=1500, two_phase=1300, phase_ground=800)
        print(fc.three_phase)  # 1500
        print(fc.max_phase)    # 1500
        fc.three_phase = 999   # ERROR! Cannot modify frozen dataclass
    """
    three_phase: float
    two_phase: float
    phase_ground: float

    @property
    def max_phase(self) -> float:
        """Maximum of 3-phase and 2-phase fault currents."""
        return max(self.three_phase, self.two_phase)


@dataclass
class Feeder:
    obj: pft.ElmFeeder
    cubicle: pft.StaCubic
    term: pft.ElmTerm
    sys_volts: float
    devices: list
    bu_devices: Dict
    open_points: Dict


@dataclass
class Device:
    obj: Any
    cubicle: pft.StaCubic
    term: pft.ElmTerm
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
    obj: pft.ElmTerm
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
    max_faults: Optional[FaultCurrents] = None
    min_faults: Optional[FaultCurrents] = None

@dataclass
class Line:
    obj: pft.ElmLne
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
    obj: Union[pft.ElmLod, pft.ElmTr2]
    term: pft.ElmTerm
    load_kva: Optional[float]
    max_ph: Optional[str]
    max_pg: Optional[str]
    fuse: Optional[str]
    insulation: Optional[str]
    impedance: Optional[str]


def initialise_fdr_dataclass(element: Any) -> Feeder:

    dataclass = Feeder(
            element,
            element.obj_id,
            element.cn_bus,
            element.cn_bus.uknom,
            [],
            [],
            {}
            )
    return dataclass


def initialise_dev_dataclass(element: Any) -> Device:

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
            ph_attr_lookup(cubicle.cterm.phtech),
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


def initialise_term_dataclass(elmterm: pft.ElmTerm) -> Union[None, Termination]:

    if elmterm is None:
        return None
    dataclass = Termination(
            elmterm,
            None,
            ph_attr_lookup(elmterm.phtech),
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


def initialise_line_dataclass(elmlne: pft.ElmLne) -> Union[None, Line]:

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

    def _get_voltage(elmlne: pft.ElmLne) -> float:
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

    def _get_conductor(elmlne: pft.ElmLne) -> str:
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


def initialise_load_dataclass(load: Union[None, pft.ElmLod, pft.ElmTr2]) -> Union[None, Tfmr]:
    if load is None:
        return Tfmr(None, None, None, None, None, None, None, None)
    if load.GetClassName() == ElementType.LOAD.value:
        return Tfmr(load, load.bus1.cterm, round(load.Strat), None, None, None, None, None)
    if load.GetClassName() == ElementType.TFMR.value:
        return Tfmr(load, load.bushv.cterm, round(load.Snom_a * 1000), None, None, None, None, None)


def ph_attr_lookup(attr: int):
    """
    Convert the terminal phase technology attribute phtech to a meaningful value
    :param attr:
    :return:
    """
    phases = {1:[6, 7, 8], 2:[2, 3, 4, 5], 3:[0, 1]}
    for phase, attr_list in phases.items():
        if attr in attr_list:
            return phase


def populate_fault_currents(terminal: 'Termination') -> None:
    """
    Create immutable FaultCurrents objects from a terminal's existing fault values.

    Call this AFTER all fault studies have populated the individual fields.
    This groups the fault values into immutable objects for safer access.

    Args:
        terminal: A Termination dataclass with fault values already populated

    Example:
        # After your fault studies complete:
        for device in devices:
            for terminal in device.sect_terms:
                populate_fault_currents(terminal)

        # Now you can access fault currents either way:
        old_way = terminal.max_fl_3ph           # Still works
        new_way = terminal.max_faults.three_phase  # Also works, immutable
    """
    # Create max faults object (or None if values are missing)
    if terminal.max_fl_3ph is not None:
        terminal.max_faults = FaultCurrents(
            three_phase=terminal.max_fl_3ph or 0,
            two_phase=terminal.max_fl_2ph or 0,
            phase_ground=terminal.max_fl_pg or 0
        )

    # Create min faults object (or None if values are missing)
    if terminal.min_fl_3ph is not None:
        terminal.min_faults = FaultCurrents(
            three_phase=terminal.min_fl_3ph or 0,
            two_phase=terminal.min_fl_2ph or 0,
            phase_ground=terminal.min_fl_pg or 0
        )