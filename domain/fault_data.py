"""
Fault current data containers for protection assessment.

This module provides immutable containers for storing fault current
values, ensuring data integrity throughout the analysis pipeline.

Classes:
    FaultCurrents: Immutable container for fault current values
"""

from dataclasses import dataclass


@dataclass(frozen=True)
class FaultCurrents:
    """
    Immutable container for fault current values at a location.

    Once created, values cannot be modified, preventing accidental
    corruption of fault study results. All values are in Amperes
    (primary).

    Attributes:
        three_phase: Three-phase symmetrical fault current (A).
        two_phase: Two-phase fault current (A).
        phase_ground: Phase-to-ground fault current (A).

    Properties:
        max_phase: Maximum of three-phase and two-phase currents.

    Example:
        >>> fc = FaultCurrents(
        ...     three_phase=1500,
        ...     two_phase=1300,
        ...     phase_ground=800
        ... )
        >>> print(fc.three_phase)
        1500
        >>> print(fc.max_phase)
        1500
        >>> fc.three_phase = 999  # Raises FrozenInstanceError
    """

    three_phase: float
    two_phase: float
    phase_ground: float

    @property
    def max_phase(self) -> float:
        """Return the maximum of three-phase and two-phase fault currents."""
        return max(self.three_phase, self.two_phase)

    def __repr__(self) -> str:
        """Return string representation with formatted current values."""
        return (
            f"FaultCurrents(3ph={self.three_phase:.0f}A, "
            f"2ph={self.two_phase:.0f}A, "
            f"pg={self.phase_ground:.0f}A)"
        )