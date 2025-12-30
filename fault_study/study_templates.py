"""Templates for PowerFactory short-circuit study configurations.

This module provides configuration templates for different fault study types.
Uses a single dataclass with a factory function.
"""

from dataclasses import dataclass, field
from typing import Dict, Any, Optional
from enum import Enum

from pf_config import pft


# =============================================================================
# ENUMS FOR STUDY CONFIGURATION
# =============================================================================

class SCMethod(Enum):
    """Short-circuit calculation method."""
    IEC60909 = 1
    Complete = 3


class FaultType(Enum):
    """PowerFactory fault type codes."""
    THREE_PHASE = '3psc'
    TWO_PHASE = '2psc'
    PHASE_TO_GROUND = 'spgf'
    TWO_PHASE_TO_GROUND = '2pgf'
    THREE_PHASE_UNBALANCED = '3rst'


class CalculationBound(Enum):
    """Maximum or minimum fault level calculation."""
    MAXIMUM = 0
    MINIMUM = 1


class ProtTrippingCurrent(Enum):
    """Protection tripping current calculation mode."""
    SUB_TRANSIENT = 0
    TRANSIENT = 1
    MIXED_MODE = 2


class EvtShcType(Enum):
    """Event short-circuit type codes."""
    THREE_PHASE = 0
    TWO_PHASE = 1
    PHASE_TO_GROUND = 2
    TWO_PHASE_TO_GROUND = 3


class ConsiderProt(Enum):
    """Consider Protection Devices."""
    NONE = 0
    ALL = 1


class FaultLocation(Enum):
    """Consider Protection Devices."""
    USER_SELECTION = 0
    BUSBARS_JUNCTIONS = 1


# =============================================================================
# UNIFIED SHORT-CIRCUIT CONFIGURATION DATACLASS
# =============================================================================

@dataclass
class ShortCircuitConfig:
    """
    Unified configuration for PowerFactory short-circuit studies.

    Attributes:
        iopt_mde: Calculation method (IEC60909 or Complete)
        iopt_shc: Fault type code
        iopt_cur: Maximum (0) or Minimum (1) calculation
        iopt_cnf: Consider network configuration
        ildfinit: Load flow initialization
        cfac_full: Voltage factor (1.1 for max, 1.0 for min)
        iIgnLoad: Ignore loads in calculation
        iIgnLneCap: Ignore line capacitance
        iIgnShnt: Ignore shunts
        iopt_prot: Protection option
        iIksForProt: Tripping current mode
        Rf: Fault resistance (ohms)
        Xf: Fault reactance (ohms)
        i_p2psc: 2-phase short circuit option (for 3ph/2ph faults)
        i_pspgf: Single-phase ground fault option (for ground faults)
    """
    iopt_mde: int = SCMethod.Complete.value
    iopt_shc: str = FaultType.THREE_PHASE.value
    iopt_cur: int = CalculationBound.MAXIMUM.value
    iopt_cnf: bool = False
    ildfinit: bool = False
    cfac_full: float = 1.1
    iIgnLoad: bool = True
    iIgnLneCap: bool = True
    iIgnShnt: bool = True
    iIksForProt: int = ProtTrippingCurrent.TRANSIENT.value
    Rf: float = 0.0
    Xf: float = 0.0
    iopt_allbus: int = FaultLocation.BUSBARS_JUNCTIONS.value
    iopt_prot: int = ConsiderProt.ALL.value # Consider all protection devices
    # Optional fields for specific fault types
    i_p2psc: int = 0  # Used for 3-phase and 2-phase minimum faults
    i_pspgf: int = 0  # Used for ground faults
    iopt_dfr: int = 0  # Used for User Selection; Terminal i or Terminal j
    shcobj: Optional[pft.ElmLne] = None  # Used for User Selection
    ppro: int = 1  # Used for User Selection; Fault Distance from Terminal i

    def as_dict(self) -> Dict[str, Any]:
        """
        Return configuration as dictionary for applying to PowerFactory.

        Excludes fields not relevant to the current fault type.
        """
        # Base fields always included
        base_fields = [
            'iopt_mde', 'iopt_shc', 'iopt_cur', 'iopt_cnf', 'ildfinit',
            'cfac_full', 'iIgnLoad', 'iIgnLneCap', 'iIgnShnt',
            'iIksForProt', 'Rf', 'Xf', 'iopt_allbus', 'iopt_prot'
        ]

        result = {f: getattr(self, f) for f in base_fields}

        # Add fault-type specific fields
        if self.iopt_shc == FaultType.PHASE_TO_GROUND.value:
            result['i_pspgf'] = self.i_pspgf
        elif self.iopt_cur == CalculationBound.MINIMUM.value:
            # i_p2psc only for minimum 3-phase and 2-phase faults
            if self.iopt_shc in (FaultType.THREE_PHASE.value, FaultType.TWO_PHASE.value):
                result['i_p2psc'] = self.i_p2psc
        if self.iopt_allbus == FaultLocation.USER_SELECTION.value:
            result['iopt_dfr'] = self.iopt_dfr
            result['shcobj'] = self.shcobj
            result['ppro'] = self.ppro

        return result


# =============================================================================
# FACTORY FUNCTION TO CREATE CONFIGURATIONS
# =============================================================================

def create_short_circuit_config(bound: str, fault_type: str, consider_prot: str, location: Optional[pft.ElmLne] = None, relative: int = 0) -> ShortCircuitConfig:
    """
    Factory function to create short-circuit study configurations.

    Args:
        bound: 'Max' or 'Min' for maximum/minimum fault level
        fault_type: '3-Phase', '2-Phase', 'Ground', 'Ground Z10', or 'Ground Z50'
        consider_prot: 'None', 'All'
        location: pft.ElmLne
        relative: 0-99
    Returns:
        ShortCircuitConfig: Configured dataclass instance

    Raises:
        ValueError: If invalid bound or fault_type provided
    """
    # Determine if maximum or minimum study
    is_max = bound == 'Max'

    # Set base values depending on max/min
    cfac = 1.1 if is_max else 1.0
    calc_bound = CalculationBound.MAXIMUM.value if is_max else CalculationBound.MINIMUM.value

    # Determine fault type and resistance
    if fault_type == '3-Phase':
        shc_type = FaultType.THREE_PHASE.value
        rf = 0.0
    elif fault_type == '2-Phase':
        shc_type = FaultType.TWO_PHASE.value
        rf = 0.0
    elif fault_type == 'Ground':
        shc_type = FaultType.PHASE_TO_GROUND.value
        rf = 0.0
    elif fault_type == 'Ground Z10':
        shc_type = FaultType.PHASE_TO_GROUND.value
        rf = 10.0
    elif fault_type == 'Ground Z50':
        shc_type = FaultType.PHASE_TO_GROUND.value
        rf = 50.0
    else:
        raise ValueError(f"Unknown fault type: {fault_type}")

    if location is not None:
        iopt_allbus = 0
        shcobj = location
        ppro = relative
    else:
        iopt_allbus = 1
        shcobj = None
        ppro = 0

    if consider_prot == 'All':
        iopt_prot = ConsiderProt.ALL.value
    else: # consider_prot == 'None':
        iopt_prot = ConsiderProt.NONE.value

    return ShortCircuitConfig(
        iopt_shc=shc_type,
        iopt_cur=calc_bound,
        cfac_full=cfac,
        Rf=rf,
        iopt_allbus=iopt_allbus,
        shcobj=shcobj,
        ppro=ppro,
        iopt_prot=iopt_prot
    )


# =============================================================================
# APPLY FUNCTION
# =============================================================================

def apply_sc(comshc: pft.ComShc, bound: str, f_type: str, consider_prot: str, location: Optional[pft.ElmTerm] = None, relative: int = 0) -> None:
    """
    Configure PowerFactory short-circuit command with study parameters.

    Args:
        comshc: PowerFactory ComShc command object
        bound: 'Max' or 'Min' for maximum/minimum fault level
        f_type: '3-Phase', '2-Phase', 'Ground', 'Ground Z10', or 'Ground Z50'
        consider_prot: 'None', 'All'
        location: pft.ElmLne
        relative: 0-99

    Example:
        ComShc = app.GetFromStudyCase("Short_Circuit.ComShc")
        apply_sc(ComShc, 'Max', '3-Phase', 'All')
        ComShc.Execute()
    """
    config = create_short_circuit_config(bound, f_type, consider_prot, location, relative)

    for attr_name, attr_value in config.as_dict().items():
        comshc.SetAttribute(attr_name, attr_value)