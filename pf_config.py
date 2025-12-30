"""Centralized configuration for PowerFactory script environment.

This module handles path configuration and provides common imports used
throughout the protection assessment project. Import this module at the
top of any file that needs PowerFactory typing support.

Usage:
    from pf_config import pft

    def my_function(app: pft.Application) -> pft.ElmRelay:
        ...
"""

import sys
from pathlib import Path

# =============================================================================
# PATH CONFIGURATION
# =============================================================================

# PowerFactory typing stubs location
_PF_TYPING_PATH = Path(r"\\Ecasd01\WksMgmt\PowerFactory\ScriptsDEV\PowerFactoryTyping")


def _configure_paths() -> None:
    """Add required paths to sys.path if not already present.

    This function is idempotent - calling it multiple times has no effect
    after the first call.
    """
    pf_typing_str = str(_PF_TYPING_PATH)

    if pf_typing_str not in sys.path:
        sys.path.append(pf_typing_str)


# Configure paths on module import
_configure_paths()

# =============================================================================
# RE-EXPORTED IMPORTS
# =============================================================================

# Now that paths are configured, import and re-export the typing module
import powerfactorytyping as pft  # noqa: E402

# Export the typing module for use by other modules
__all__ = ['pft']