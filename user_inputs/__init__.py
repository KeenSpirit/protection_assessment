"""
User input collection package for protection assessment.

This package provides GUI dialogs for collecting user inputs required
to configure and execute protection fault level studies. It handles
study type selection, feeder selection, external grid parameter entry,
and protection device selection.

Modules:
    study_selection: Initial study type selection dialog.
    get_inputs: Feeder, grid, and device selection dialogs.

Example:
    >>> from user_inputs import study_selection, get_inputs
    >>>
    >>> # Get study type selection
    >>> selections = study_selection.get_study_selections(app)
    >>>
    >>> # Get feeder and device selections
    >>> feeders, bu_devices, selection, grid = get_inputs.get_input(
    ...     app, region, selections
    ... )
"""

from user_inputs.study_selection import get_study_selections
from user_inputs.get_inputs import get_input, FaultLevelStudy

__all__ = [
    'get_study_selections',
    'get_input',
    'FaultLevelStudy',
]