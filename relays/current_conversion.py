"""
Current conversion utilities for relay analysis.

This module provides functions to convert fault currents to the values
seen by different relay measurement types (3I0, I2, etc.).

Functions:
    get_measured_current: Convert fault current to measurement type
    convert_to_i2: Convert to negative sequence current
    convert_to_i0: Convert to zero sequence current
"""


def get_measured_current(element, fault_level: float, fault_type: str) -> float:
    """
    Calculate the current seen by a relay element for a given fault.

    Different relay measurement types see different currents for the
    same fault. This function converts the fault current to what the
    specific element's measurement type would see.

    Args:
        element: Relay element (RelToc or RelIoc) with typ_id attribute.
        fault_level: Fault current magnitude in Amperes.
        fault_type: Type of fault ('3-Phase', '2-Phase', 'Phase-Ground').

    Returns:
        Current value in Amperes as seen by the element's measurement
        type.

    Measurement Type Mapping:
        - '3ph', 'd3m': 3-phase current (direct)
        - '3I0', 'S3I0': 3x zero sequence (earth) current
        - 'I0', '1ph': Zero sequence / single phase current
        - 'd1m': Single phase delta measurement
        - 'I2': Negative sequence current
        - '3I2': 3x negative sequence current

    Example:
        >>> current = get_measured_current(toc_element, 1000, '2-Phase')
    """
    measure_type = element.typ_id.atype

    if measure_type in ['3ph', 'd3m']:
        # 3-Phase current measurement
        return fault_level

    elif measure_type in ['3I0', 'S3I0']:
        # Earth current & sensitive earth current (3I0)
        return convert_to_i0(fault_level, threei0=False)

    elif measure_type in ['I0', '1ph']:
        # Zero sequence current & 1 phase current
        return fault_level

    elif measure_type in ['d1m']:
        # 1 phase delta measurement
        return fault_level

    elif measure_type in ['I2']:
        # Negative sequence current
        return convert_to_i2(fault_level, fault_type, threei2=False)

    elif measure_type in ['3I2']:
        # 3x negative sequence current
        return convert_to_i2(fault_level, fault_type, threei2=True)

    else:
        # Unknown measurement type - return 3-phase current as default
        return fault_level


def convert_to_i2(
    fault_current: float,
    fault_type: str,
    threei2: bool = False
) -> float:
    """
    Convert fault current to negative sequence current (I2).

    For unbalanced faults, calculates the negative sequence component
    using symmetrical component analysis. Assumes fault currents are
    approximately 180 degrees apart for 2-phase faults.

    Args:
        fault_current: Phase fault current magnitude in Amperes.
        fault_type: Type of fault ('3-Phase', '2-Phase', 'Phase-Ground').
        threei2: If True, return 3*I2 instead of I2.

    Returns:
        Negative sequence current (or 3*I2) in Amperes.
        Returns 0 for 3-phase faults (balanced, no negative sequence).

    Theory:
        For a 2-phase fault (B-C): Ia=0, Ib=-Ic
        I2 = (Ia + a²Ib + aIc) / 3
        where a = e^(j120°) = -0.5 + j0.866

    Example:
        >>> i2 = convert_to_i2(1000, '2-Phase')
        >>> print(f"Negative sequence: {i2:.1f}A")
    """
    if fault_type == '2-Phase':
        # 2-phase fault: Ia=0, Ib=+I, Ic=-I (approximately)
        ia = complex(0, fault_current)
        ib = complex(0, -fault_current)
    elif fault_type == 'Phase-Ground':
        # Phase-ground fault: Ia=I, Ib=0, Ic=0
        ia = complex(0, fault_current)
        ib = complex(0, 0)
    else:
        # 3-phase fault: balanced, no negative sequence
        return 0

    ic = complex(0, 0)

    # Symmetrical component transformation constants
    a = complex(-0.5, 0.866)   # 120° rotation operator
    a2 = complex(-0.5, -0.866)  # 240° rotation operator

    # Calculate sequence component products
    a2ib = ib * a2
    aic = ic * a

    # Negative sequence current: I2 = (Ia + a²Ib + aIc) / 3
    ia2 = (ia + a2ib + aic) / 3
    ib2 = ia2 * a
    ic2 = ia2 * a2

    # Get magnitudes
    ia2_mag = abs(ia2)
    ib2_mag = abs(ib2)
    ic2_mag = abs(ic2)

    if not threei2:
        # Return average (they should be equal in magnitude)
        result = (ia2_mag + ib2_mag + ic2_mag) / 3
    else:
        # Return 3*I2 (sum of magnitudes)
        result = ia2_mag + ib2_mag + ic2_mag

    return result


def convert_to_i0(fault_current: float, threei0: bool = False) -> float:
    """
    Convert fault current to zero sequence current (I0).

    For earth faults, the zero sequence current equals the fault current
    divided equally among the three phases.

    Args:
        fault_current: Earth fault current magnitude in Amperes.
        threei0: If True, return 3*I0 instead of I0.

    Returns:
        Zero sequence current (or 3*I0) in Amperes.

    Theory:
        For a single-phase-to-ground fault: Ia=If, Ib=0, Ic=0
        I0 = (Ia + Ib + Ic) / 3 = If / 3
        3I0 = If (the neutral/earth current)

    Example:
        >>> i0 = convert_to_i0(900, threei0=False)
        >>> print(f"Zero sequence: {i0:.1f}A")  # 300A
        >>> i0_3 = convert_to_i0(900, threei0=True)
        >>> print(f"3I0: {i0_3:.1f}A")  # 2700A
    """
    if threei0:
        return 3 * fault_current
    else:
        return fault_current