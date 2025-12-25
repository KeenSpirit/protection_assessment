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
from pf_config import pft
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
    """
    Represents a distribution feeder and its associated devices.

    Attributes:
        obj: The PowerFactory ElmFeeder object
        cubicle: The feeder's source cubicle
        term: The source terminal
        sys_volts: System voltage in kV
        devices: List of protection devices on this feeder
        bu_devices: Backup devices dictionary
        open_points: Network open points dictionary
    """
    obj: pft.ElmFeeder
    cubicle: pft.StaCubic
    term: pft.ElmTerm
    sys_volts: float
    devices: List[Any] = field(default_factory=list)
    bu_devices: Dict[str, Any] = field(default_factory=dict)
    open_points: Dict[str, Any] = field(default_factory=dict)


@dataclass
class Device:
    """
    Represents a protection device (relay or fuse) with its electrical context.

    Attributes:
        obj: The PowerFactory device object (ElmRelay or RelFuse)
        cubicle: The device's cubicle location
        term: The connected terminal
        phases: Number of phases (1, 2, or 3)
        l_l_volts: Line-to-line voltage in kV

    Fault Currents (populated by fault studies):
        max_fl_3ph, max_fl_2ph, max_fl_pg: Maximum fault currents
        min_fl_3ph, min_fl_2ph, min_fl_pg: Minimum fault currents
        min_sn_fl_2ph, min_sn_fl_pg: Minimum system normal fault currents

    Topology (populated by network tracing):
        sect_terms: Downstream terminals in this section
        sect_loads: Downstream loads in this section
        sect_lines: Lines in this section
        us_devices: Upstream protection devices
        ds_devices: Downstream protection devices
    """
    # Core identification - always required
    obj: Any
    cubicle: pft.StaCubic
    term: pft.ElmTerm
    phases: int
    l_l_volts: float

    # Capacity data - populated during analysis
    ds_capacity: Optional[float] = None
    max_ds_tr: Optional[float] = None

    # Maximum fault currents - populated by fault study
    max_fl_3ph: Optional[float] = None
    max_fl_2ph: Optional[float] = None
    max_fl_pg: Optional[float] = None

    # Minimum fault currents - populated by fault study
    min_fl_3ph: Optional[float] = None
    min_fl_2ph: Optional[float] = None
    min_fl_pg: Optional[float] = None

    # Minimum system normal fault currents - populated by fault study
    min_sn_fl_2ph: Optional[float] = None
    min_sn_fl_pg: Optional[float] = None

    # Topology - populated by network tracing
    sect_terms: List[Any] = field(default_factory=list)
    sect_loads: List[Any] = field(default_factory=list)
    sect_lines: List[Any] = field(default_factory=list)
    us_devices: List[Any] = field(default_factory=list)
    ds_devices: List[Any] = field(default_factory=list)


@dataclass
class Termination:
    """
    Represents a network terminal with fault current data.

    The terminal collects fault study results for protection coordination.

    Attributes:
        obj: The PowerFactory ElmTerm object
        phases: Number of phases (1, 2, or 3)
        l_l_volts: Line-to-line voltage in kV
        constr: Construction type (OH, UG, SWER) - set during analysis
        fault_data: Complete fault current data - set after fault studies
    """
    # Core identification - always required
    obj: pft.ElmTerm
    phases: int
    l_l_volts: float

    # Construction type - determined during analysis
    constr: Optional[str] = None

    # Fault currents - populated by fault study
    max_fl_3ph: Optional[float] = None
    max_fl_2ph: Optional[float] = None
    max_fl_pg: Optional[float] = None
    min_fl_3ph: Optional[float] = None
    min_fl_2ph: Optional[float] = None
    min_fl_pg: Optional[float] = None
    min_fl_pg10: Optional[float] = None
    min_fl_pg50: Optional[float] = None
    min_sn_fl_2ph: Optional[float] = None
    min_sn_fl_pg: Optional[float] = None
    min_sn_fl_pg10: Optional[float] = None
    min_sn_fl_pg50: Optional[float] = None
    max_faults: Optional[FaultCurrents] = None
    min_faults: Optional[FaultCurrents] = None

@dataclass
class Line:
    """
        Represents a distribution line with fault current and energy data.

        Attributes:
            obj: The PowerFactory ElmLne object
            phases: Number of phases (1, 2, or 3)
            l_l_volts: Line-to-line voltage in kV
            line_type: Conductor type description
            thermal_rating: Thermal current rating in Amps

        Fault Currents (populated by fault studies):
            max_fl_*, min_fl_*, min_sn_fl_*: Various fault current values

        Energy Data (populated by energy calculations):
            ph_energy, pg_energy: Let-through energy values
            ph_clear_time, pg_clear_time: Clearing times
            ph_fl, pg_fl: Fault levels used for energy calculation
        """
    # Core identification - always required
    obj: pft.ElmLne
    phases: int
    l_l_volts: float
    line_type: str
    thermal_rating: Union[float, str]  # Can be 'NA' for cables

    # Maximum fault currents - populated by fault study
    max_fl_3ph: Optional[float] = None
    max_fl_2ph: Optional[float] = None
    max_fl_pg: Optional[float] = None

    # Minimum fault currents - populated by fault study
    min_fl_3ph: Optional[float] = None
    min_fl_2ph: Optional[float] = None
    min_fl_pg: Optional[float] = None

    # Minimum system normal fault currents
    min_sn_fl_2ph: Optional[float] = None
    min_sn_fl_pg: Optional[float] = None

    # Energy calculations - populated during analysis
    ph_energy: Optional[float] = None
    ph_clear_time: Optional[float] = None
    ph_fl: Optional[float] = None
    pg_energy: Optional[float] = None
    pg_clear_time: Optional[float] = None
    pg_fl: Optional[float] = None


@dataclass
class Tfmr:
    """
    Represents a transformer or load for fusing coordination.

    Attributes:
        obj: The PowerFactory object (ElmLod or ElmTr2), or None
        term: The HV terminal connection
        load_kva: Rated capacity in kVA
        max_ph: Maximum phase fault at transformer
        max_pg: Maximum phase-ground fault at transformer
        fuse: Associated fuse type/size
        insulation: Insulation class
        impedance: Transformer impedance
    """
    obj: Optional[Union[pft.ElmLod, pft.ElmTr2]] = None
    term: Optional[pft.ElmTerm] = None
    load_kva: Optional[float] = None
    max_ph: Optional[str] = None
    max_pg: Optional[str] = None
    fuse: Optional[str] = None
    insulation: Optional[str] = None
    impedance: Optional[str] = None


def initialise_fdr_dataclass(element: pft.ElmFeeder) -> Feeder:
    """
    Initialize a Feeder dataclass from a PowerFactory ElmFeeder.

    Args:
        element: The PowerFactory feeder object

    Returns:
        Initialized Feeder dataclass
    """
    return Feeder(
        obj=element,
        cubicle=element.obj_id,
        term=element.cn_bus,
        sys_volts=element.cn_bus.uknom,
    )


def initialise_dev_dataclass(element: Any) -> Device:
    """
    Initialize a Device dataclass from a PowerFactory protection device.

    Args:
        element: The PowerFactory device object (ElmRelay, RelFuse, etc.)

    Returns:
        Initialized Device dataclass, or None if element is None
    """
    if element is None:
        return None

    # Determine the cubicle based on element type
    class_name = element.GetClassName()
    if class_name == ElementType.FEEDER.value:
        cubicle = element.obj_id
    elif class_name == ElementType.COUPLER.value:
        cubicle = element.bus1
    else:
        cubicle = element.fold_id

    return Device(
        obj=element,
        cubicle=cubicle,
        term=cubicle.cterm,
        phases=ph_attr_lookup(cubicle.cterm.phtech),
        l_l_volts=round(cubicle.cterm.uknom, 2),
    )


def initialise_term_dataclass(elmterm: pft.ElmTerm) -> Optional[Termination]:
    """
    Initialize a Termination dataclass from a PowerFactory ElmTerm.

    Args:
        elmterm: The PowerFactory terminal object

    Returns:
        Initialized Termination dataclass, or None if elmterm is None
    """
    if elmterm is None:
        return None

    return Termination(
        obj=elmterm,
        phases=ph_attr_lookup(elmterm.phtech),
        l_l_volts=round(elmterm.uknom, 2),
    )


def initialise_line_dataclass(elmlne: pft.ElmLne) -> Optional[Line]:
    """
    Initialize a Line dataclass from a PowerFactory ElmLne.

    Args:
        elmlne: The PowerFactory line object

    Returns:
        Initialized Line dataclass, or None if elmlne is None
    """
    if elmlne is None:
        return None

    def _get_phases(line: pft.ElmLne) -> int:
        """Get number of phases from line construction type."""
        construction = line.typ_id.GetClassName()

        if construction == 'TypGeo':
            return int(line.typ_id.xy_c[0][0])
        elif construction == 'TypCabsys':
            return int(line.typ_id.GetAttribute('nphas')[0])
        elif construction == "TypLne":
            return int(line.typ_id.GetAttribute('nlnph'))
        else:
            raise TypeError(f'{construction} Unhandled construction type')

    def _get_voltage(line: pft.ElmLne) -> float:
        """Get line-to-line operating voltage."""
        for term in line.GetConnectedElements():
            try:
                return term.uknom
            except AttributeError:
                pass
        return 0.0

    def _get_conductor_info(line: pft.ElmLne) -> tuple:
        """Get conductor type and thermal rating."""
        construction = line.typ_id.GetClassName()

        if construction == 'TypGeo':
            typ_con = line.GetAttribute("e:pCondCir")
            return typ_con.loc_name, round(typ_con.GetAttribute("e:Ithr"), 3) * 1000
        elif construction == "TypLne":
            return line.typ_id.loc_name, round(line.typ_id.GetAttribute("e:Ithr"), 3) * 1000
        else:
            return 'NA', 'NA'

    line_type, thermal_rating = _get_conductor_info(elmlne)

    return Line(
        obj=elmlne,
        phases=_get_phases(elmlne),
        l_l_volts=_get_voltage(elmlne),
        line_type=line_type,
        thermal_rating=thermal_rating,
    )


def initialise_load_dataclass(load: Union[None, pft.ElmLod, pft.ElmTr2]) -> Optional[Tfmr]:
    """
    Initialize a Tfmr dataclass from a PowerFactory load or transformer.

    Args:
        load: The PowerFactory load (ElmLod) or transformer (ElmTr2) object

    Returns:
        Initialized Tfmr dataclass (empty if load is None)
    """
    if load is None:
        return Tfmr()
    if load.GetClassName() == ElementType.LOAD.value:
        return Tfmr(
            obj=load,
            term=load.bus1.cterm,
            load_kva=round(load.Strat),
        )
    if load.GetClassName() == ElementType.TFMR.value:
        return Tfmr(
            obj=load,
            term=load.bushv.cterm,
            load_kva=round(load.Snom_a * 1000),
        )
    return Tfmr()


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