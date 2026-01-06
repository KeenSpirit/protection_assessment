"""
Utility functions for domain model operations.

This module contains functions that operate on domain models but don't
belong to a specific model class.

Functions:
    populate_fault_currents: Create immutable FaultCurrents from terminal
"""

from typing import TYPE_CHECKING

from domain.fault_data import FaultCurrents

if TYPE_CHECKING:
    from domain.termination import Termination


def populate_fault_currents(terminal: "Termination") -> None:
    """
    Create immutable FaultCurrents objects from terminal fault values.

    Call this AFTER all fault studies have populated the individual
    fault current fields on the terminal. This function groups the
    fault values into immutable FaultCurrents containers for safer
    downstream access.

    The function creates two containers:
        - max_faults: Contains max_fl_3ph, max_fl_2ph, max_fl_pg
        - min_faults: Contains min_fl_3ph, min_fl_2ph, min_fl_pg

    Args:
        terminal: A Termination dataclass with fault values already
            populated from fault studies.

    Side Effects:
        Sets terminal.max_faults and terminal.min_faults attributes.

    Example:
        >>> # After fault studies complete:
        >>> for device in devices:
        ...     for terminal in device.sect_terms:
        ...         populate_fault_currents(terminal)
        >>>
        >>> # Now access fault currents either way:
        >>> old_way = terminal.max_fl_3ph
        >>> new_way = terminal.max_faults.three_phase  # Immutable
        >>>
        >>> # The immutable container prevents modification:
        >>> terminal.max_faults.three_phase = 999  # FrozenInstanceError

    Note:
        If the required fault current values are None, the corresponding
        FaultCurrents container will not be created (remains None).
    """
    # Create max faults container if values are present
    if terminal.max_fl_3ph is not None:
        terminal.max_faults = FaultCurrents(
            three_phase=terminal.max_fl_3ph or 0,
            two_phase=terminal.max_fl_2ph or 0,
            phase_ground=terminal.max_fl_pg or 0
        )

    # Create min faults container if values are present
    if terminal.min_fl_3ph is not None:
        terminal.min_faults = FaultCurrents(
            three_phase=terminal.min_fl_3ph or 0,
            two_phase=terminal.min_fl_2ph or 0,
            phase_ground=terminal.min_fl_pg or 0
        )