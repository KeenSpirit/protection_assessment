"""
Domain models for PowerFactory protection assessment.

This package contains the core domain models (dataclasses) used throughout
the protection assessment system. Each model represents a distinct concept
in the electrical network domain.

Modules:
    enums: Element types, construction types, fault types
    fault_data: Immutable fault current containers
    feeder: Distribution feeder model
    device: Protection device model (relays, fuses)
    termination: Network terminal model
    line: Distribution line model
    transformer: Transformer/load model for fusing
    utils: Utility functions for domain operations

Usage:
    # Import the entire domain package (recommended for backward compatibility)
    import domain as dd

    # Or import specific items
    from domain import Device, Feeder, ElementType
    from domain.termination import Termination, initialise_term_dataclass

Backward Compatibility:
    This package provides the same interface as the original script_classes.py
    module. Existing code using `import script_classes as dd` can be migrated
    by changing to `import domain as dd` with minimal other changes.

Example:
    >>> import domain as dd
    >>>
    >>> # Create domain objects from PowerFactory elements
    >>> feeder = dd.initialise_fdr_dataclass(elm_feeder)
    >>> device = dd.initialise_dev_dataclass(elm_relay)
    >>>
    >>> # Check element types
    >>> if device.obj.GetClassName() == dd.ElementType.RELAY.value:
    ...     print("This is a relay")
    >>>
    >>> # After fault studies, populate immutable containers
    >>> dd.populate_fault_currents(terminal)
"""

# =============================================================================
# ENUMERATIONS
# =============================================================================

from domain.enums import (
    ElementType,
    ConstructionType,
    FaultType,
    ph_attr_lookup,
)

# =============================================================================
# FAULT DATA
# =============================================================================

from domain.fault_data import FaultCurrents

# =============================================================================
# DOMAIN MODELS
# =============================================================================

from domain.feeder import Feeder, initialise_fdr_dataclass
from domain.device import Device, initialise_dev_dataclass
from domain.termination import Termination, initialise_term_dataclass
from domain.line import Line, initialise_line_dataclass
from domain.transformer import Tfmr, initialise_load_dataclass

# =============================================================================
# UTILITIES
# =============================================================================

from domain.utils import populate_fault_currents

# =============================================================================
# PUBLIC API
# =============================================================================

__all__ = [
    # Enums
    "ElementType",
    "ConstructionType",
    "FaultType",
    "ph_attr_lookup",
    # Fault data
    "FaultCurrents",
    # Domain models
    "Feeder",
    "Device",
    "Termination",
    "Line",
    "Tfmr",
    # Initializers
    "initialise_fdr_dataclass",
    "initialise_dev_dataclass",
    "initialise_term_dataclass",
    "initialise_line_dataclass",
    "initialise_load_dataclass",
    # Utilities
    "populate_fault_currents",
]