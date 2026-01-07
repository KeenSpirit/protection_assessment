"""
PowerFactory conditional diagram colouring for protection assessment.

This module configures colour formatting for line elements based on
protection assessment results. It maps reach factors and conductor
damage pass/fail conditions to visual colour attributes.

The colour scheme uses three DPL attributes on line elements. Each
attribute encodes a pair of assessment results using a 4x4 grid
(16 possible combinations):

    dpl1: Phase Pri Reach (row) × Phase BU Reach (column)
    dpl2: Earth Pri Reach (row) × Earth BU Reach (column)
    dpl3: Phase Cond Damage (row) × Earth Cond Damage (column)

Encoding formula: value = (row_condition - 1) * 4 + column_condition
where conditions are: Pass=1, Fail=2, No Data=3, SWER=4

Colour Mapping:
    - Green (3): Pass
    - Red (2): Fail
    - Grey (9): No Data
    - Yellow (6): SWER

Functions:
    colour_map: Main entry point for colour configuration
    set_up: Configure filters and clear existing attributes
    configure_quick_filters: Create IntFilt objects in settings
    configure_colour_conditions: Create SetColours condition sets
    clear_dpl_attributes: Reset all DPL values to zero
"""

from enum import IntEnum
from typing import Dict, List, Set, Tuple, TYPE_CHECKING, Union

if TYPE_CHECKING:
    from pf_config import pft

from relays.reach_factors import device_reach_factors


# =============================================================================
# Type Aliases
# =============================================================================

ReachFactorValue = Union[float, str]
DplLookup = Dict[str, Dict[str, str]]


# =============================================================================
# Enumerations
# =============================================================================

class Condition(IntEnum):
    """
    Assessment result conditions mapped to attribute values.

    Values are 1-indexed to match the DPL encoding scheme.
    """
    PASS = 1
    FAIL = 2
    NO_DATA = 3
    SWER = 4


class Colour(IntEnum):
    """
    PowerFactory colour codes for conditional formatting.

    Corresponds to PowerFactory's internal colour palette indices.
    """
    GREEN = 3   # Pass
    RED = 2     # Fail
    GREY = 9    # No Data
    YELLOW = 6  # SWER


# =============================================================================
# Lookup Tables
# =============================================================================

CONDITION_LOOKUP: Dict[str, Tuple[Condition, Colour]] = {
    "Pass": (Condition.PASS, Colour.GREEN),
    "Fail": (Condition.FAIL, Colour.RED),
    "No Data": (Condition.NO_DATA, Colour.GREY),
    "SWER": (Condition.SWER, Colour.YELLOW),
}

CONDITION_INDEX: Dict[str, int] = {
    "Pass": 0,
    "Fail": 1,
    "No Data": 2,
    "SWER": 3,
}

REACH_THRESHOLDS: Dict[str, Dict[str, float]] = {
    "SEQ": {"primary": 2.0, "backup": 1.3},
    "Regional Models": {"primary": 1.7, "backup": 1.3},
}

PRIMARY_REACH_TYPES: Set[str] = {"Phase Pri Reach", "Earth Pri Reach"}
BACKUP_REACH_TYPES: Set[str] = {"Phase BU Reach", "Earth BU Reach"}
CONDUCTOR_DAMAGE_TYPES: Set[str] = {"Phase Cond Damage", "Earth Cond Damage"}

REACH_FACTOR_KEYS: Dict[str, Tuple[str, str]] = {
    "Phase Pri Reach": ("ph_rf", "nps_ph_rf"),
    "Earth Pri Reach": ("ef_rf", "nps_ef_rf"),
    "Phase BU Reach": ("bu_ph_rf", "bu_nps_ph_rf"),
    "Earth BU Reach": ("bu_ef_rf", "bu_nps_ef_rf"),
}

DPL_CONFIG: Dict[str, Tuple[str, str]] = {
    "Phase Pri Reach": ("dpl1", "row"),
    "Phase BU Reach": ("dpl1", "col"),
    "Earth Pri Reach": ("dpl2", "row"),
    "Earth BU Reach": ("dpl2", "col"),
    "Phase Cond Damage": ("dpl3", "row"),
    "Earth Cond Damage": ("dpl3", "col"),
}

DPL_MAP_TYPE_PAIRS: Dict[str, Tuple[str, str]] = {
    "dpl1": ("Phase Pri Reach", "Phase BU Reach"),
    "dpl2": ("Earth Pri Reach", "Earth BU Reach"),
    "dpl3": ("Phase Cond Damage", "Earth Cond Damage"),
}

LINE_ELEMENT_FILTER: List[str] = ["*.ElmLne"]

ALL_MAP_TYPES: List[str] = [
    "Phase Pri Reach", "Earth Pri Reach",
    "Phase BU Reach", "Earth BU Reach",
    "Phase Cond Damage", "Earth Cond Damage"
]


# =============================================================================
# Main Entry Point
# =============================================================================

def colour_map(
    app: "pft.Application",
    region: str,
    feeders: List,
    study_selections: List[str]
) -> None:
    """
    Configure colour maps and write protection assessment results.

    Main entry point for colour map configuration. Creates filters for
    each feeder and map type, then evaluates protection assessments
    for all devices and their line sections.

    Args:
        app: PowerFactory application object.
        region: Network region ('SEQ' or 'Regional Models').
        feeders: List of Feeder dataclasses to process.
        study_selections: List of selected study type strings.

    Side Effects:
        - Creates IntFilt objects in project settings
        - Creates SetColours condition sets
        - Writes DPL attribute values to line elements

    Note:
        Only runs if 'Fault Level Study (all relays configured in model)'
        is in study_selections.

    Example:
        >>> colour_map(app, 'SEQ', feeders, study_selections)
        Protection reach results saved in PowerFactory...
    """
    study_type = "Fault Level Study (all relays configured in model)"
    if study_type not in study_selections:
        return

    # Determine map types to process
    map_types = [
        "Phase Pri Reach", "Earth Pri Reach",
        "Phase BU Reach", "Earth BU Reach"
    ]

    include_cond_damage = "Conductor Damage Assessment" in study_selections
    if include_cond_damage:
        map_types.extend(["Phase Cond Damage", "Earth Cond Damage"])

    project = app.GetActiveProject()
    set_up(app, project, feeders, map_types)

    app.SetGraphicUpdate(0)
    try:
        for feeder in feeders:
            _process_feeder(app, region, feeder, map_types)
            _print_completion_message(app, feeder, include_cond_damage)
    finally:
        app.SetGraphicUpdate(1)


# =============================================================================
# Setup Functions
# =============================================================================

def set_up(
    app: "pft.Application",
    project: "pft.DataObject",
    feeders: List,
    map_types: List[str]
) -> None:
    """
    Configure colour filters and conditional formatting.

    Args:
        app: PowerFactory application object.
        project: Active PowerFactory project.
        feeders: List of Feeder dataclasses.
        map_types: List of map type strings to create filters for.
    """
    filter_names = []
    for feeder in feeders:
        filter_names.extend([
            f"{feeder.obj.loc_name} {name}" for name in map_types
        ])

    settings_folder = configure_quick_filters(project, filter_names)
    clear_dpl_attributes(app, project)
    configure_colour_conditions(settings_folder, filter_names)


def configure_quick_filters(
    project: "pft.DataObject",
    filter_names: List[str]
) -> "pft.DataObject":
    """
    Create quick filters in the project settings folder.

    Args:
        project: Active PowerFactory project.
        filter_names: List of filter names to create.

    Returns:
        Settings folder DataObject for colour configuration.
    """
    settings_folder = project.GetContents("*.SetFold", True)[0]

    # Remove existing filters with matching names
    existing_filters = settings_folder.GetContents("*.IntFilt", True)
    for existing_filter in existing_filters:
        if any(name in existing_filter.loc_name for name in filter_names):
            existing_filter.Delete()

    # Create new filters
    for name in filter_names:
        general_filter = settings_folder.CreateObject("IntFilt", name)
        _create_condition_filters(general_filter, name)

    return settings_folder


def configure_colour_conditions(
    settings_folder: "pft.DataObject",
    filter_names: List[str]
) -> None:
    """
    Configure conditional colour formatting in project colour settings.

    Args:
        settings_folder: Project settings folder.
        filter_names: List of filter names to create condition sets for.
    """
    colour_folder = settings_folder.GetContents("*.SetColours", True)[0]

    for name in filter_names:
        # Remove existing condition sets
        existing_sets = colour_folder.GetContents("*.IntFiltSet", True)
        for existing_set in existing_sets:
            if name in existing_set.loc_name:
                existing_set.Delete()

        condition_set = colour_folder.CreateObject("IntFiltset", name)
        _create_condition_filters(condition_set, name)


def _create_condition_filters(
    parent_object: "pft.DataObject",
    filter_name: str
) -> None:
    """
    Create condition filter objects under a parent filter.

    Creates Pass/Fail/No Data/SWER child filters with appropriate
    DPL expressions and colour settings.

    Args:
        parent_object: IntFilt or IntFiltSet parent object.
        filter_name: Full filter name to determine map type.
    """
    map_type = _extract_map_type_from_name(filter_name)
    dpl_attr, position = DPL_CONFIG[map_type]

    for condition_name, (_, colour) in CONDITION_LOOKUP.items():
        parent_object.CreateObject("SetFilt", condition_name)
        condition_filter = parent_object.GetContents(
            f"{condition_name}.SetFilt"
        )[0]

        condition_filter.SetAttribute("objset", LINE_ELEMENT_FILTER)
        condition_filter.SetAttribute("icalcrel", 1)
        condition_filter.SetAttribute("icoups", 0)

        # Build expression for this condition
        values = get_dpl_values_for_condition(position, condition_name)
        expression = build_filter_expression(dpl_attr, values)

        condition_filter.SetAttribute("expr", [expression])
        condition_filter.SetAttribute("color", int(colour))


def clear_dpl_attributes(
    app: "pft.Application",
    project: "pft.DataObject"
) -> None:
    """
    Clear all DPL attributes from line elements.

    Args:
        app: PowerFactory application object.
        project: Active PowerFactory project.
    """
    app.SetGraphicUpdate(0)

    try:
        active_lines = get_active_lines(project)

        for element in active_lines:
            for dpl_index in range(1, 6):
                attr_name = f"e:dpl{dpl_index}"
                if element.GetAttribute(attr_name):
                    element.SetAttribute(attr_name, 0)
    finally:
        app.SetGraphicUpdate(1)


# =============================================================================
# Processing Functions
# =============================================================================

def _process_feeder(
    app: "pft.Application",
    region: str,
    feeder,
    map_types: List[str]
) -> None:
    """
    Process all devices for a single feeder.

    Args:
        app: PowerFactory application object.
        region: Network region identifier.
        feeder: Feeder dataclass containing devices.
        map_types: List of map types to assess.
    """
    for device in feeder.devices:
        _process_device(app, region, device, map_types)


def _process_device(
    app: "pft.Application",
    region: str,
    device,
    map_types: List[str]
) -> None:
    """
    Process all line sections for a single protection device.

    Args:
        app: PowerFactory application object.
        region: Network region identifier.
        device: Device dataclass with sect_lines.
        map_types: List of map types to assess.
    """
    dev_reach_factors = device_reach_factors(
        region, device, device.sect_lines
    )

    for i, line in enumerate(device.sect_lines):
        dpl_lookup = create_default_dpl_lookup()

        # Conductor damage assessments
        for map_type in map_types:
            if map_type in CONDUCTOR_DAMAGE_TYPES:
                _assess_conductor_damage(line, map_type, dpl_lookup)

        # Reach assessments
        for map_type, keys in REACH_FACTOR_KEYS.items():
            if map_type not in map_types:
                continue

            reach_factor = max_mixed_values(
                dev_reach_factors[keys[0]][i],
                dev_reach_factors[keys[1]][i]
            )
            _assess_reach(region, map_type, reach_factor, dpl_lookup)

        _write_result(dpl_lookup, line)


def _print_completion_message(
    app: "pft.Application",
    feeder,
    include_conductor_damage: bool
) -> None:
    """
    Print completion message with created colour scheme names.

    Args:
        app: PowerFactory application object.
        feeder: Processed Feeder dataclass.
        include_conductor_damage: Whether conductor damage was included.
    """
    name = feeder.obj.loc_name

    app.PrintPlain(
        f"Protection reach results saved in PowerFactory as user-defined "
        f"diagram colouring schemes:\n"
        f"'{name} Phase Pri Reach'\n"
        f"'{name} Earth Pri Reach'\n"
        f"'{name} Phase BU Reach'\n"
        f"'{name} Earth BU Reach'."
    )

    if include_conductor_damage:
        app.PrintPlain(
            f"Conductor damage results saved in PowerFactory as "
            f"user-defined diagram colouring schemes:\n"
            f"'{name} Phase Cond Damage'\n"
            f"'{name} Earth Cond Damage'."
        )


# =============================================================================
# Assessment Functions
# =============================================================================

def _assess_reach(
    region: str,
    map_type: str,
    reach_factor: ReachFactorValue,
    dpl_lookup: DplLookup
) -> None:
    """
    Assess reach factor against regional thresholds.

    Args:
        region: Network region ('SEQ' or 'Regional Models').
        map_type: Reach assessment type string.
        reach_factor: Calculated reach factor or 'NA' string.
        dpl_lookup: Dictionary to update with assessment result.
    """
    dpl_attr, _ = DPL_CONFIG[map_type]

    if isinstance(reach_factor, str):
        dpl_lookup[dpl_attr][map_type] = "No Data"
        return

    threshold = get_reach_threshold(region, map_type)
    condition = "Pass" if reach_factor >= threshold else "Fail"
    dpl_lookup[dpl_attr][map_type] = condition


def _assess_conductor_damage(
    line,
    map_type: str,
    dpl_lookup: DplLookup
) -> None:
    """
    Assess conductor damage based on fault energy vs thermal rating.

    Args:
        line: Line dataclass with thermal and energy attributes.
        map_type: 'Phase Cond Damage' or 'Earth Cond Damage'.
        dpl_lookup: Dictionary to update with assessment result.
    """
    dpl_attr, _ = DPL_CONFIG[map_type]

    if _is_swer_line(line):
        dpl_lookup[dpl_attr][map_type] = "SWER"
        return

    energy_attr = "ph_energy" if map_type == "Phase Cond Damage" else "pg_energy"

    try:
        line_energy = getattr(line, energy_attr, None)
        thermal_rating = getattr(line, "thermal_rating", None)

        if line_energy is None or thermal_rating is None:
            dpl_lookup[dpl_attr][map_type] = "No Data"
            return

        allowable_energy = thermal_rating ** 2
        condition = "Pass" if line_energy <= allowable_energy else "Fail"
        dpl_lookup[dpl_attr][map_type] = condition

    except (AttributeError, TypeError):
        dpl_lookup[dpl_attr][map_type] = "No Data"


def _is_swer_line(line) -> bool:
    """
    Check if a line is a SWER (Single Wire Earth Return) line.

    Args:
        line: Line dataclass with obj.typ_id attribute.

    Returns:
        True if line type contains 'SWER', False otherwise.
    """
    try:
        line_type = line.obj.typ_id
        return "SWER" in line_type.loc_name
    except AttributeError:
        return False


# =============================================================================
# Result Writing
# =============================================================================

def _write_result(dpl_lookup: DplLookup, line) -> None:
    """
    Write colour scheme results to a line element.

    Encodes each condition pair into a single DPL value using
    the 4x4 grid encoding scheme.

    Args:
        dpl_lookup: Dictionary mapping DPL attributes to conditions.
        line: Line dataclass with obj attribute for PowerFactory element.
    """
    for dpl_attr, (row_type, col_type) in DPL_MAP_TYPE_PAIRS.items():
        row_condition = dpl_lookup[dpl_attr][row_type]
        col_condition = dpl_lookup[dpl_attr][col_type]

        value = encode_dpl_value(row_condition, col_condition)
        line.obj.SetAttribute(f"e:{dpl_attr}", value)


# =============================================================================
# Helper Functions
# =============================================================================

def max_mixed_values(
    a: ReachFactorValue,
    b: ReachFactorValue
) -> ReachFactorValue:
    """
    Return maximum of two values that may be numbers or 'NA' strings.

    Args:
        a: First value (float or 'NA' string).
        b: Second value (float or 'NA' string).

    Returns:
        'NA' if both are strings, the numeric value if one is string,
        or the maximum if both are numeric.
    """
    a_is_string = isinstance(a, str)
    b_is_string = isinstance(b, str)

    if a_is_string and b_is_string:
        return "NA"
    if a_is_string:
        return b
    if b_is_string:
        return a
    return max(a, b)


def get_protection_category(map_type: str) -> str:
    """
    Determine protection category for a given map type.

    Args:
        map_type: Map type string.

    Returns:
        'primary' or 'backup'.

    Raises:
        ValueError: If map_type is not a reach assessment type.
    """
    if map_type in PRIMARY_REACH_TYPES:
        return "primary"
    if map_type in BACKUP_REACH_TYPES:
        return "backup"
    raise ValueError(f"Map type '{map_type}' is not a reach assessment type")


def get_reach_threshold(region: str, map_type: str) -> float:
    """
    Get reach factor threshold for a region and map type.

    Args:
        region: Network region ('SEQ' or 'Regional Models').
        map_type: Protection map type string.

    Returns:
        Minimum reach factor threshold for pass condition.
    """
    category = get_protection_category(map_type)
    return REACH_THRESHOLDS[region][category]


def get_active_lines(project: "pft.DataObject") -> List["pft.DataObject"]:
    """
    Retrieve all line elements connected to the grid.

    Args:
        project: Active PowerFactory project.

    Returns:
        List of ElmLne objects with grid connections.
    """
    return [
        line
        for line in project.GetContents("*.ElmLne", True)
        if line.GetAttribute("cpGrid")
    ]


def create_default_dpl_lookup() -> DplLookup:
    """
    Create DPL lookup dictionary with all values set to 'No Data'.

    Returns:
        Dictionary mapping DPL attributes to condition pair dicts.
    """
    return {
        "dpl1": {"Phase Pri Reach": "No Data", "Phase BU Reach": "No Data"},
        "dpl2": {"Earth Pri Reach": "No Data", "Earth BU Reach": "No Data"},
        "dpl3": {"Phase Cond Damage": "No Data", "Earth Cond Damage": "No Data"},
    }


def encode_dpl_value(row_condition: str, col_condition: str) -> int:
    """
    Encode two condition strings into a single DPL integer value.

    Uses 4x4 grid encoding: value = row_index * 4 + col_index + 1

    Args:
        row_condition: Row element condition string.
        col_condition: Column element condition string.

    Returns:
        Integer value 1-16 representing the condition pair.
    """
    row_idx = CONDITION_INDEX[row_condition]
    col_idx = CONDITION_INDEX[col_condition]
    return row_idx * 4 + col_idx + 1


def get_dpl_values_for_condition(position: str, condition: str) -> List[int]:
    """
    Get all DPL values matching a condition in a given position.

    For 'row' position: returns values where row matches condition.
    For 'col' position: returns values where column matches condition.

    Args:
        position: 'row' or 'col'.
        condition: Condition string ('Pass', 'Fail', etc.).

    Returns:
        List of 4 integer values matching the condition.
    """
    condition_idx = CONDITION_INDEX[condition]

    if position == "row":
        base = condition_idx * 4
        return [base + i for i in range(1, 5)]
    else:
        return [condition_idx + 1 + i * 4 for i in range(4)]


def build_filter_expression(dpl_attr: str, values: List[int]) -> str:
    """
    Build PowerFactory filter expression for matching DPL values.

    Args:
        dpl_attr: DPL attribute name (e.g., 'dpl1').
        values: List of integer values to match.

    Returns:
        PowerFactory expression string.

    Example:
        >>> build_filter_expression('dpl1', [1, 2, 3, 4])
        '{e:dpl1=1}.or.{e:dpl1=2}.or.{e:dpl1=3}.or.{e:dpl1=4}'
    """
    clauses = [f"{{e:{dpl_attr}={v}}}" for v in values]
    return ".or.".join(clauses)


def _extract_map_type_from_name(filter_name: str) -> str:
    """
    Extract map type from a filter name.

    Args:
        filter_name: Full filter name like 'Feeder1 Phase Pri Reach'.

    Returns:
        Map type string.

    Raises:
        ValueError: If no matching map type found.
    """
    for map_type in ALL_MAP_TYPES:
        if map_type in filter_name:
            return map_type
    raise ValueError(f"No matching map type found in name: {filter_name}")