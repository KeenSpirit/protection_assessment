"""
PowerFactory Colour Map Configuration Module

This module configures conditional colour formatting for line elements in DIgSILENT
PowerFactory based on protection assessment results. It maps protection study types
and pass/fail conditions to visual colour attributes on network line elements.

The colour scheme uses three DPL attributes on line elements. Each attribute encodes
a pair of assessment results using a 4x4 grid (16 possible combinations):

    dpl1: Phase Pri Reach (row) × Phase BU Reach (column)
    dpl2: Earth Pri Reach (row) × Earth BU Reach (column)
    dpl3: Phase Cond Damage (row) × Earth Cond Damage (column)

The encoding formula is: value = (row_condition - 1) * 4 + column_condition
where conditions are: Pass=1, Fail=2, No Data=3, SWER=4

These attributes are evaluated by PowerFactory's conditional formatting filters
to display lines in appropriate colours based on protection assessment results.

Assessment Types:
    - Reach Assessment: Compares relay reach factors against regional thresholds
    - Conductor Damage Assessment: Compares fault energy against thermal ratings
"""

from enum import IntEnum
from typing import TYPE_CHECKING, Union

if TYPE_CHECKING:
    from powerfactory import Application, DataObject

from devices import relays


# =============================================================================
# Enumerations
# =============================================================================

class Condition(IntEnum):
    """
    Assessment result conditions mapped to attribute values.

    These represent the outcome of each protection assessment check.
    Values are 1-indexed to match the DPL encoding scheme.
    """
    PASS = 1
    FAIL = 2
    NO_DATA = 3
    SWER = 4


class Colour(IntEnum):
    """
    PowerFactory colour codes for conditional formatting.

    These correspond to PowerFactory's internal colour palette indices.
    """
    GREEN = 3   # Pass
    RED = 2     # Fail
    GREY = 9    # No Data
    YELLOW = 6  # SWER


# =============================================================================
# Lookup Tables
# =============================================================================

# Condition string -> (Condition enum, Colour enum)
CONDITION_LOOKUP: dict[str, tuple[Condition, Colour]] = {
    "Pass": (Condition.PASS, Colour.GREEN),
    "Fail": (Condition.FAIL, Colour.RED),
    "No Data": (Condition.NO_DATA, Colour.GREY),
    "SWER": (Condition.SWER, Colour.YELLOW),
}

# Reverse lookup: condition string -> 0-based index for encoding
CONDITION_INDEX: dict[str, int] = {
    "Pass": 0,
    "Fail": 1,
    "No Data": 2,
    "SWER": 3,
}

# Reach factor thresholds by region and protection type
REACH_THRESHOLDS: dict[str, dict[str, float]] = {
    "SEQ": {
        "primary": 2.0,
        "backup": 1.3,
    },
    "Regional Models": {
        "primary": 1.7,
        "backup": 1.3,
    },
}

# Map types classified by protection category
PRIMARY_REACH_TYPES: set[str] = {"Phase Pri Reach", "Earth Pri Reach"}
BACKUP_REACH_TYPES: set[str] = {"Phase BU Reach", "Earth BU Reach"}
CONDUCTOR_DAMAGE_TYPES: set[str] = {"Phase Cond Damage", "Earth Cond Damage"}

# Mapping from map_type to device_reach_factors dictionary keys
REACH_FACTOR_KEYS: dict[str, tuple[str, str]] = {
    "Phase Pri Reach": ("ph_rf", "nps_ph_rf"),
    "Earth Pri Reach": ("ef_rf", "nps_ef_rf"),
    "Phase BU Reach": ("bu_ph_rf", "bu_nps_ph_rf"),
    "Earth BU Reach": ("bu_ef_rf", "bu_nps_ef_rf"),
}

# DPL attribute configuration: maps each map_type to its DPL attribute and position
# Position "row" means it's the first element in the pair (encoded as row in 4x4 grid)
# Position "col" means it's the second element (encoded as column)
DPL_CONFIG: dict[str, tuple[str, str]] = {
    "Phase Pri Reach": ("dpl1", "row"),
    "Phase BU Reach": ("dpl1", "col"),
    "Earth Pri Reach": ("dpl2", "row"),
    "Earth BU Reach": ("dpl2", "col"),
    "Phase Cond Damage": ("dpl3", "row"),
    "Earth Cond Damage": ("dpl3", "col"),
}

# Maps DPL attribute to the pair of map_types it encodes (row_type, col_type)
DPL_MAP_TYPE_PAIRS: dict[str, tuple[str, str]] = {
    "dpl1": ("Phase Pri Reach", "Phase BU Reach"),
    "dpl2": ("Earth Pri Reach", "Earth BU Reach"),
    "dpl3": ("Phase Cond Damage", "Earth Cond Damage"),
}

# PowerFactory element filter for line elements
LINE_ELEMENT_FILTER: list[str] = ["*.ElmLne"]

# All available map types in display order
ALL_MAP_TYPES: list[str] = [
    "Phase Pri Reach", "Earth Pri Reach",
    "Phase BU Reach", "Earth BU Reach",
    "Phase Cond Damage", "Earth Cond Damage"
]


# =============================================================================
# Type Aliases
# =============================================================================

ReachFactorValue = Union[float, str]
DplLookup = dict[str, dict[str, str]]


# =============================================================================
# Helper Functions
# =============================================================================

def max_mixed_values(a: ReachFactorValue, b: ReachFactorValue) -> ReachFactorValue:
    """
    Return the maximum of two values that may be numbers or "NA" strings.

    Args:
        a: First value (float or string like "NA").
        b: Second value (float or string like "NA").

    Returns:
        - "NA" if both inputs are strings
        - The numeric value if one input is a string
        - The maximum if both are numeric
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
    Determine the protection category for a given map type.

    Args:
        map_type: The map type string.

    Returns:
        Either "primary" or "backup".

    Raises:
        ValueError: If map_type is not a reach-type assessment.
    """
    if map_type in PRIMARY_REACH_TYPES:
        return "primary"
    if map_type in BACKUP_REACH_TYPES:
        return "backup"
    raise ValueError(f"Map type '{map_type}' is not a reach assessment type")


def get_reach_threshold(region: str, map_type: str) -> float:
    """
    Get the reach factor threshold for a region and map type combination.

    Args:
        region: The network region ("SEQ" or "Regional Models").
        map_type: The protection map type.

    Returns:
        The minimum reach factor threshold for pass condition.
    """
    category = get_protection_category(map_type)
    return REACH_THRESHOLDS[region][category]


def get_active_lines(project: "DataObject") -> list["DataObject"]:
    """
    Retrieve all line elements connected to the grid.

    Args:
        project: The active PowerFactory project.

    Returns:
        List of ElmLne objects that have a grid connection.
    """
    return [
        line
        for line in project.GetContents("*.ElmLne", True)
        if line.GetAttribute("cpGrid")
    ]


def create_default_dpl_lookup() -> DplLookup:
    """
    Create a fresh DPL lookup dictionary with all values set to "No Data".

    Returns:
        Dictionary mapping DPL attributes to their map_type condition pairs.
    """
    return {
        "dpl1": {"Phase Pri Reach": "No Data", "Phase BU Reach": "No Data"},
        "dpl2": {"Earth Pri Reach": "No Data", "Earth BU Reach": "No Data"},
        "dpl3": {"Phase Cond Damage": "No Data", "Earth Cond Damage": "No Data"},
    }


def encode_dpl_value(row_condition: str, col_condition: str) -> int:
    """
    Encode two condition strings into a single DPL integer value.

    Uses a 4x4 grid encoding where:
        value = row_index * 4 + col_index + 1

    Args:
        row_condition: Condition for the row element ("Pass", "Fail", etc.)
        col_condition: Condition for the column element.

    Returns:
        Integer value 1-16 representing the condition pair.
    """
    row_idx = CONDITION_INDEX[row_condition]
    col_idx = CONDITION_INDEX[col_condition]
    return row_idx * 4 + col_idx + 1


def get_dpl_values_for_condition(position: str, condition: str) -> list[int]:
    """
    Get all DPL values that match a condition in a given position.

    For "row" position: returns all values where the row matches the condition.
    For "col" position: returns all values where the column matches the condition.

    Args:
        position: Either "row" or "col".
        condition: The condition string ("Pass", "Fail", etc.)

    Returns:
        List of 4 integer values that match the condition in that position.
    """
    condition_idx = CONDITION_INDEX[condition]

    if position == "row":
        # Row determines the "tens" place: values are condition_idx*4 + 1,2,3,4
        base = condition_idx * 4
        return [base + i for i in range(1, 5)]
    else:  # col
        # Column determines the "ones" place: values are condition_idx+1 + 0,4,8,12
        return [condition_idx + 1 + i * 4 for i in range(4)]


def build_filter_expression(dpl_attr: str, values: list[int]) -> str:
    """
    Build a PowerFactory filter expression for matching DPL values.

    Args:
        dpl_attr: The DPL attribute name (e.g., "dpl1").
        values: List of integer values to match.

    Returns:
        PowerFactory expression string like "{e:dpl1=1}.or.{e:dpl1=2}..."
    """
    clauses = [f"{{e:{dpl_attr}={v}}}" for v in values]
    return ".or.".join(clauses)


# =============================================================================
# Main Entry Point
# =============================================================================

def colour_map(
    app: "Application",
    region: str,
    feeders: list,
    study_selections: list[str]
) -> None:
    """
    Main entry point for colour map configuration and result writing.

    Configures the colour filters for feeders and evaluates protection
    assessments for all devices and their associated line sections.

    Args:
        app: PowerFactory application object.
        region: Network region identifier ("SEQ" or "Regional Models").
        feeders: List of feeder objects to process.
        study_selections: List of selected study types.
    """
    if "Fault Level Study (all relays configured in model)" not in study_selections:
        return

    # Determine which map types to process
    map_types = ["Phase Pri Reach", "Earth Pri Reach", "Phase BU Reach", "Earth BU Reach"]
    include_conductor_damage = "Conductor Damage Assessment" in study_selections
    if include_conductor_damage:
        map_types.extend(["Phase Cond Damage", "Earth Cond Damage"])

    project = app.GetActiveProject()
    set_up(app, project, feeders, map_types)

    app.SetGraphicUpdate(0)
    try:
        for feeder in feeders:
            _process_feeder(app, region, feeder, map_types)
            _print_completion_message(app, feeder, include_conductor_damage)
    finally:
        app.SetGraphicUpdate(1)


def _process_feeder(
    app: "Application",
    region: str,
    feeder,
    map_types: list[str]
) -> None:
    """
    Process all devices for a single feeder.

    Args:
        app: PowerFactory application object.
        region: Network region identifier.
        feeder: Feeder object containing devices.
        map_types: List of map types to assess.
    """
    for device in feeder.devices:
        _process_device(app, region, device, map_types)


def _process_device(
    app: "Application",
    region: str,
    device,
    map_types: list[str]
) -> None:
    """
    Process all line sections for a single protection device.

    Args:
        app: PowerFactory application object.
        region: Network region identifier.
        device: Protection device object.
        map_types: List of map types to assess.
    """
    device_reach_factors = relays.device_reach_factors(region, device, device.sect_lines)

    for i, line in enumerate(device.sect_lines):
        # Create fresh lookup for each line
        dpl_lookup = create_default_dpl_lookup()

        # Process conductor damage assessments
        for map_type in map_types:
            if map_type in CONDUCTOR_DAMAGE_TYPES:
                _assess_conductor_damage(line, map_type, dpl_lookup)

        # Process reach assessments
        for map_type, keys in REACH_FACTOR_KEYS.items():
            if map_type not in map_types:
                continue

            reach_factor = max_mixed_values(
                device_reach_factors[keys[0]][i],
                device_reach_factors[keys[1]][i]
            )
            _assess_reach(region, map_type, reach_factor, dpl_lookup)

        _write_result(dpl_lookup, line)


def _print_completion_message(
    app: "Application",
    feeder,
    include_conductor_damage: bool
) -> None:
    """
    Print completion message with created colour scheme names.

    Args:
        app: PowerFactory application object.
        feeder: Feeder object that was processed.
        include_conductor_damage: Whether conductor damage was included.
    """
    feeder_name = feeder.obj.loc_name
    app.PrintPlain(
        f"Protection reach results saved in PowerFactory as user-defined "
        f"diagram colouring schemes:\n"
        f"'{feeder_name} Phase Pri Reach'\n"
        f"'{feeder_name} Earth Pri Reach'\n"
        f"'{feeder_name} Phase BU Reach'\n"
        f"'{feeder_name} Earth BU Reach'."
    )
    if include_conductor_damage:
        app.PrintPlain(
            f"Conductor damage results saved in PowerFactory as user-defined "
            f"diagram colouring schemes:\n"
            f"'{feeder_name} Phase Cond Damage'\n"
            f"'{feeder_name} Earth Cond Damage'."
        )


# =============================================================================
# Setup Functions
# =============================================================================

def set_up(
    app: "Application",
    project: "DataObject",
    feeders: list,
    map_types: list[str]
) -> None:
    """
    Configure colour filters and conditional formatting for feeders.

    Args:
        app: PowerFactory application object.
        project: The active PowerFactory project.
        feeders: List of feeder objects.
        map_types: List of map types to create filters for.
    """
    filter_names = []
    for feeder in feeders:
        filter_names.extend([f"{feeder.obj.loc_name} {name}" for name in map_types])

    settings_folder = configure_quick_filters(project, filter_names)
    clear_dpl_attributes(app, project)
    configure_colour_conditions(settings_folder, filter_names)


def configure_quick_filters(
    project: "DataObject",
    filter_names: list[str]
) -> "DataObject":
    """
    Configure quick filters in the Project Colour Settings folder.

    Args:
        project: The active PowerFactory project.
        filter_names: List of filter names to create.

    Returns:
        The settings folder DataObject for use in colour configuration.
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


def _create_condition_filters(parent_object: "DataObject", filter_name: str) -> None:
    """
    Create condition filter objects (Pass/Fail/No Data/SWER) under a parent filter.

    Args:
        parent_object: The IntFilt or IntFiltSet parent object.
        filter_name: The full filter name to determine map type.
    """
    # Determine which map type this filter is for
    map_type = _extract_map_type_from_name(filter_name)
    dpl_attr, position = DPL_CONFIG[map_type]

    for condition_name, (_, colour) in CONDITION_LOOKUP.items():
        parent_object.CreateObject("SetFilt", condition_name)
        condition_filter = parent_object.GetContents(f"{condition_name}.SetFilt")[0]

        condition_filter.SetAttribute("objset", LINE_ELEMENT_FILTER)
        condition_filter.SetAttribute("icalcrel", 1)
        condition_filter.SetAttribute("icoups", 0)

        # Build expression from the encoding pattern
        values = get_dpl_values_for_condition(position, condition_name)
        expression = build_filter_expression(dpl_attr, values)

        condition_filter.SetAttribute("expr", [expression])
        condition_filter.SetAttribute("color", int(colour))


def _extract_map_type_from_name(filter_name: str) -> str:
    """
    Extract the map type from a filter name.

    Args:
        filter_name: Full filter name like "Feeder1 Phase Pri Reach".

    Returns:
        The map type string.

    Raises:
        ValueError: If no matching map type is found.
    """
    for map_type in ALL_MAP_TYPES:
        if map_type in filter_name:
            return map_type
    raise ValueError(f"No matching map type found in name: {filter_name}")


def clear_dpl_attributes(app: "Application", project: "DataObject") -> None:
    """
    Clear all DPL attributes from line elements in the model.

    Args:
        app: PowerFactory application object.
        project: The active PowerFactory project.
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


def configure_colour_conditions(
    settings_folder: "DataObject",
    filter_names: list[str]
) -> None:
    """
    Configure conditional colour formatting in the project colour settings.

    Args:
        settings_folder: The project settings folder.
        filter_names: List of filter names to create condition sets for.
    """
    colour_folder = settings_folder.GetContents("*.SetColours", True)[0]

    for name in filter_names:
        # Remove existing condition sets with this name
        existing_sets = colour_folder.GetContents("*.IntFiltSet", True)
        for existing_set in existing_sets:
            if name in existing_set.loc_name:
                existing_set.Delete()

        condition_set = colour_folder.CreateObject("IntFiltset", name)
        _create_condition_filters(condition_set, name)


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
    Assess reach factor against regional thresholds and update lookup.

    Args:
        region: Network region identifier ("SEQ" or "Regional Models").
        map_type: The reach assessment type.
        reach_factor: Calculated reach factor (numeric) or "NA" string.
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
        line: Line object with thermal and energy attributes.
        map_type: Either "Phase Cond Damage" or "Earth Cond Damage".
        dpl_lookup: Dictionary to update with assessment result.
    """
    dpl_attr, _ = DPL_CONFIG[map_type]

    # Check for SWER line
    if _is_swer_line(line):
        dpl_lookup[dpl_attr][map_type] = "SWER"
        return

    # Get the appropriate energy attribute
    energy_attr = "ph_energy" if map_type == "Phase Cond Damage" else "pg_energy"

    try:
        line_energy = getattr(line, energy_attr, None)
        thermal_rating = getattr(line, "thermal_rating", None)

        if line_energy is None or thermal_rating is None:
            dpl_lookup[dpl_attr][map_type] = "No Data"
            return

        # Calculate allowable energy (I²t with t=1 second)
        allowable_energy = thermal_rating ** 2
        condition = "Pass" if line_energy <= allowable_energy else "Fail"
        dpl_lookup[dpl_attr][map_type] = condition

    except (AttributeError, TypeError):
        dpl_lookup[dpl_attr][map_type] = "No Data"


def _is_swer_line(line) -> bool:
    """
    Check if a line is a SWER (Single Wire Earth Return) line.

    Args:
        line: Line object with .obj.typ_id attribute.

    Returns:
        True if the line type contains "SWER", False otherwise.
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

    Encodes each pair of conditions into a single DPL value using
    the 4x4 grid encoding scheme.

    Args:
        dpl_lookup: Dictionary mapping DPL attributes to condition pairs.
        line: Line object with .obj attribute for PowerFactory element.
    """
    for dpl_attr, (row_type, col_type) in DPL_MAP_TYPE_PAIRS.items():
        row_condition = dpl_lookup[dpl_attr][row_type]
        col_condition = dpl_lookup[dpl_attr][col_type]

        value = encode_dpl_value(row_condition, col_condition)
        line.obj.SetAttribute(f"e:{dpl_attr}", value)