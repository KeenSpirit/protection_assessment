"""
Relay analysis package.

This package provides protection device analysis functionality split into
focused modules:

- elements: Relay/fuse retrieval and element filtering
- current_conversion: Fault current to measurement type conversion
- reclose: Auto-reclose sequence management
- reach_factors: relay reach factor calculations

For backward compatibility, all functions are re-exported at the package level.

Usage (new style - targeted imports):
    from relays.elements import get_all_relays, get_prot_elements
    from relays.reach_factors import device_reach_factors

Usage (backward compatible):
    from relay import get_all_relays, device_reach_factors
    # or
    import relays
    relays = relays.get_all_relays(app)
"""

# =============================================================================
# ELEMENT RETRIEVAL AND FILTERING
# =============================================================================

from relays.elements import (
    get_all_relays,
    get_prot_elements,
    get_active_elements,
)

# =============================================================================
# CURRENT CONVERSION
# =============================================================================

from relays.current_conversion import (
    get_measured_current,
    convert_to_i2,
    convert_to_i0,
)

# =============================================================================
# AUTO-RECLOSE SEQUENCE MANAGEMENT
# =============================================================================

from relays.reclose import (
    get_device_trips,
    reset_reclosing,
    trip_count,
    set_enabled_elements,
    reset_block_service_status,
)

# =============================================================================
# REACH FACTOR CALCULATIONS
# =============================================================================

from relays.reach_factors import (
    device_reach_factors,
    determine_pickup_values,
    swer_transform,
)

# =============================================================================
# PUBLIC API
# =============================================================================

__all__ = [
    # elements
    'get_all_relays',
    'get_prot_elements',
    'get_active_elements',
    # current_conversion
    'get_measured_current',
    'convert_to_i2',
    'convert_to_i0',
    # reclose
    'get_device_trips',
    'reset_reclosing',
    'trip_count',
    'set_enabled_elements',
    'reset_block_service_status',
    # reach_factors
    'device_reach_factors',
    'determine_pickup_values',
    'swer_transform',
]