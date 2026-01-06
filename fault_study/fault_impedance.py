"""
Fault impedance and construction type handling.

This module determines appropriate fault current values based on network
region, line construction type, and study type. Different regions use
different fault impedance assumptions for earth faults:

- SEQ: Uses a single fault impedance value (0 ohms)
- Regional: Uses construction-dependent values (50Ω OH, 10Ω UG)

Functions:
    update_node_construction: Set construction type for all terminals
    get_terminal_pg_fault: Get appropriate PG fault current for terminal
"""

from typing import List

from pf_config import pft
import domain as dd


def update_node_construction(devices: List[dd.Device]) -> None:
    """
    Determine and set the construction type for all terminal nodes.

    Examines connected line elements to determine if each terminal is
    connected to overhead (OH), underground (UG), or SWER construction.
    This information is used to select appropriate fault impedance values.

    Args:
        devices: List of Device dataclasses with sect_terms populated.

    Side Effects:
        Sets the 'constr' attribute on each Termination in device.sect_terms
        to one of: 'OH', 'UG', or 'SWER'.
    """
    all_nodes = _get_all_terms(devices)
    _update_construction(all_nodes)


def _get_all_terms(devices: List[dd.Device]) -> List[dd.Termination]:
    """
    Extract all terminations from device list.

    Args:
        devices: List of Device dataclasses.

    Returns:
        Flat list of all Termination objects from all devices.
    """
    all_nodes = []
    for device in devices:
        terms = device.sect_terms
        all_nodes.extend(terms)
    return all_nodes


def _update_construction(all_nodes: List[dd.Termination]) -> None:
    """
    Set construction type for each node based on connected lines.

    Examines lines connected to each terminal and sets the construction
    type. Priority order: SWER > UG > OH (default).

    Args:
        all_nodes: List of Termination dataclasses to update.

    Side Effects:
        Sets constr attribute on each Termination.
    """
    for node in all_nodes:
        # Skip if already determined
        if node.constr is not None:
            continue

        # Get all lines connected to the node
        line_elements = [
            ele for ele in node.obj.GetConnectedElements()
            if ele.GetClassName() == dd.ElementType.LINE.value
        ]

        # Handle case where upstream connection is not a line (e.g., ElmCoup)
        if not line_elements:
            try:
                substation = node.cpSubstat
                proxy_node = substation.pBusbar
                line_elements = [
                    ele for ele in proxy_node.GetConnectedElements()
                    if ele.GetClassName() == dd.ElementType.LINE.value
                ]
            except (AttributeError, IndexError):
                line_elements = []

        # Determine construction type from connected lines
        for line in line_elements:
            try:
                line_type = line.typ_id
                if 'SWER' in line_type.loc_name:
                    node.constr = "SWER"
                    break
            except AttributeError:
                pass

            if line.IsCable() and node.constr != "OH":
                node.constr = "UG"
            else:
                node.constr = "OH"

        # Default to overhead if no lines found
        if node.constr is None:
            node.constr = "OH"


def get_terminal_pg_fault(
    region: str,
    term: dd.Termination,
    system_normal: bool = False
) -> float:
    """
    Get the appropriate minimum phase-ground fault current for a terminal.

    Selects the correct fault current value based on region, construction
    type, and study type. This accounts for different fault impedance
    assumptions used in different network regions.

    Args:
        region: Network region identifier ('SEQ' or 'Regional Models').
        term: Termination dataclass with fault current attributes.
        system_normal: If True, returns system normal minimum values.
            If False (default), returns minimum values.

    Returns:
        Minimum phase-ground fault current in Amperes.
        Returns 0 if the required attribute is None.

    Fault Impedance by Region:
        SEQ:
            - All constructions: 0Ω fault impedance
            - Uses: min_fl_pg or min_sn_fl_pg

        Regional Models:
            - Overhead (OH): 50Ω fault impedance
            - Underground (UG): 10Ω fault impedance
            - Uses: min_fl_pg50/min_sn_fl_pg50 for OH
                    min_fl_pg10/min_sn_fl_pg10 for UG

    Example:
        >>> # Get minimum fault current for reach factor calculation
        >>> min_fl = get_terminal_pg_fault('SEQ', terminal)
        >>>
        >>> # Get system normal minimum for reporting
        >>> sn_min_fl = get_terminal_pg_fault('SEQ', terminal, True)
    """
    if region == 'SEQ':
        # SEQ: Single fault impedance value (0 ohms)
        if system_normal:
            fault_level = term.min_sn_fl_pg
        else:
            fault_level = term.min_fl_pg
    else:
        # Regional: Construction-dependent fault impedance
        if term.constr == 'OH':
            # Overhead: 50 ohm fault impedance
            if system_normal:
                fault_level = term.min_sn_fl_pg50
            else:
                fault_level = term.min_fl_pg50
        else:
            # Underground/Cable: 10 ohm fault impedance
            if system_normal:
                fault_level = term.min_sn_fl_pg10
            else:
                fault_level = term.min_fl_pg10

    return fault_level if fault_level is not None else 0