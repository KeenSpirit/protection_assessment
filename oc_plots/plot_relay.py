"""
Time-overcurrent coordination plot generation for PowerFactory.

This module creates time-overcurrent (TOC) characteristic plots in the
PowerFactory graphics board. Plots show relay curves, fault level markers,
and downstream transformer fuse curves for protection coordination review.

Plot Features:
    - Primary device curves with unique colours
    - Upstream and downstream device curves
    - Minimum and maximum fault level markers
    - Downstream transformer fuse curves (if applicable)
    - Separate plots for phase and earth faults

Functions:
    plot_all_relays: Main entry point for plot generation
    update_ds_tr_data: Collect RMU fuse specifications for SEQ
    create_colour_dic: Assign unique colours to devices
    create_plot: Generate a single coordination plot
    create_plot_folder: Create timestamped folder for plots
    _add_fault_markers: Configure fault level marker lines
    _add_transformer_fuse: Add fuse curve and marker to plot
"""

import time
from typing import Dict, List

from pf_config import pft
from devices import fuses as ds
from oc_plots import get_rmu_fuses as grf
from pf_protection_helper import create_obj, obtain_region
import domain as dd
from oc_plots import plot_settings
from importlib import reload
reload (plot_settings)


def plot_all_relays(
    app: pft.Application,
    feeder: dd.Feeder,
    selected_devices: List[dd.Device]
) -> None:
    """
    Generate time-overcurrent coordination plots for selected devices.

    Creates separate phase and earth fault plots for each terminal
    location with relay devices. Plots are saved to a timestamped
    folder in the PowerFactory graphics board.

    Args:
        app: PowerFactory application instance.
        feeder: Feeder dataclass containing device topology.
        selected_devices: List of Device dataclasses to plot.

    Side Effects:
        - Creates page format in project settings
        - Creates plot folder in graphics board
        - Generates VisOcplot objects for each device group
        - For SEQ region, prompts for RMU fuse specifications

    Note:
        Devices sharing a parent terminal are plotted together.
        Phase plots are only created if minimum 2-phase fault > 0.

    Example:
        >>> plot_all_relays(app, feeder, selected_devices)
        Time overcurrent plots saved in PowerFactory to...
    """
    app.PrintPlain(
        f"Generating device time overcurrent plots for "
        f"{feeder.obj.loc_name}..."
    )

    study_case = app.GetActiveStudyCase()
    graphics_board = app.GetFromStudyCase("Graphics Board.SetDesktop")
    project = app.GetActiveProject()

    study_case.Deactivate()
    plot_folder = create_plot_folder(feeder, project)

    # Collect RMU fuse data for SEQ transformers
    update_ds_tr_data(app, selected_devices)

    # Create colour mapping for all devices
    colour_dic = create_colour_dic(feeder.devices)

    # Group devices by parent terminal
    terminals = set(device.term.loc_name for device in selected_devices)
    device_dic = {
        terminal: [
            device for device in selected_devices
            if device.term.loc_name == terminal
        ]
        for terminal in terminals
    }

    # Generate plots for each terminal group
    for devices in device_dic.values():
        has_relays = any(
            device.obj.GetClassName() == dd.ElementType.RELAY.value
            for device in devices
        )

        if has_relays:

            # Earth fault plot
            create_plot(
                app, plot_folder, colour_dic, devices,
                feeder.sys_volts, f_type='Ground'
            )

            # Phase fault plot (if applicable)
            if any(device.min_fl_2ph > 0 for device in devices):
                create_plot(
                    app, plot_folder, colour_dic, devices,
                    feeder.sys_volts, f_type='Phase'
                )

        else:
            app.PrintPlain(
                f'No relays in feeder {feeder.obj.loc_name} to plot.'
            )

    study_case.Activate()
    directory = plot_folder.GetFullName()
    app.PrintPlain(
        f"Time overcurrent plots saved in PowerFactory to:\n{directory}"
    )

# =============================================================================
# TRANSFORMER DATA COLLECTION
# =============================================================================

def update_ds_tr_data(
    app: pft.Application,
    device_list: List[dd.Device]
) -> None:
    """
    Collect RMU fuse specifications for SEQ downstream transformers.

    For SEQ region, prompts user via GUI to enter insulation type
    (air/oil) and impedance class (high/low) for RMU-connected
    transformers ≥750kVA.

    Args:
        app: PowerFactory application instance.
        device_list: List of Device dataclasses to process.

    Side Effects:
        Updates max_ds_tr.insulation and max_ds_tr.impedance
        attributes on Device objects.

    Note:
        Only applies to SEQ region. Regional models use standard
        fuse selection without user input.
    """
    region = obtain_region(app)

    if region != "SEQ":
        return

    tr_strings_dic = {}

    for device in device_list:
        if device.obj.GetClassName() != dd.ElementType.RELAY.value:
            continue

        max_ds_tr = device.max_ds_tr
        if max_ds_tr is None:
            continue

        tr_name = max_ds_tr.term.cpSubstat.loc_name

        # Skip substations (SP prefix)
        if tr_name[:2] == "SP":
            continue

        # Encode transformer size for GUI lookup
        if max_ds_tr.load_kva in [1000, 1500]:
            max_tr_string = tr_name + "_1"
        else:
            max_tr_string = tr_name + "_0"

        tr_strings_dic[device.obj] = max_tr_string

    if not tr_strings_dic:
        return

    # Prompt user for specifications
    max_tr_strings = list(tr_strings_dic.values())
    app.PrintPlain("Please enter distribution transformer fuse specification")
    results = grf.get_transformer_specifications(max_tr_strings)

    # Apply results to device objects
    for device_object, string in tr_strings_dic.items():
        device = [d for d in device_list if d.obj == device_object][0]
        string_result = results[string]
        max_ds_tr = device.max_ds_tr
        max_ds_tr.insulation = string_result['insulation']
        max_ds_tr.impedance = string_result['impedance']


# =============================================================================
# COLOUR ASSIGNMENT
# =============================================================================

def create_colour_dic(devices: List[dd.Device]) -> Dict:
    """
    Assign unique colours to each device for plot curves.

    Ensures all devices (primary, upstream, downstream) have distinct
    colours when plotted together.

    Args:
        devices: List of all Device dataclasses in feeder.

    Returns:
        Dictionary mapping device objects to colour indices.

    Note:
        Colour index 0 is white, so assignments start at 1.
    """
    all_devices = []

    for device in devices:
        all_devices.append(device)
        all_devices.extend(device.us_devices)
        all_devices.extend(device.ds_devices)

    # Remove duplicates while preserving order
    unique_devices = []
    for device in all_devices:
        if device not in unique_devices:
            unique_devices.append(device)

    # Colour 0 is white, 1 is black, start with 2 (Red)
    colour_dic = {device.obj: i + 2 for i, device in enumerate(unique_devices)}

    return colour_dic


def create_plot(
    app: pft.Application,
    plot_folder: pft.IntFolder,
    colour_dic: Dict,
    devices: List[dd.Device],
    sys_volts: str,
    f_type: str
) -> None:
    """
    Create a single time-overcurrent coordination plot.

    Generates a VisOcplot containing:
    - Primary device characteristic curves
    - Upstream and downstream device curves
    - Minimum and maximum fault level markers
    - Downstream transformer fuse curve (if applicable)

    Args:
        app: PowerFactory application instance.
        plot_folder: Target folder for plot.
        colour_dic: Device to colour index mapping.
        devices: List of devices at same terminal to plot.
        sys_volts: System voltage string for fuse selection.
        f_type: 'Ground' or 'Phase' fault type.

    Returns:
        SetVipage object containing the plot.
    """
    date_string = time.strftime("%Y%m%d")

    # Build device name string
    devices_name = ""
    for device in devices:
        devices_name = devices_name + "_" + device.obj.loc_name

    # Create plot page
    folder_name = f"{devices_name} {f_type} Coord Plot {date_string}"
    grppage = create_obj(plot_folder, folder_name, "GrpPage")
    drawing_format = create_obj(grppage, "Drawing Format", "SetGrfpage")
    plot_settings.setup_drawing_format(drawing_format)

    # Configure plot settings
    pltlinebarplot = grppage.GetOrInsertPlot(folder_name, 10, 1)
    plot_settings.axis_settings(pltlinebarplot, f_type, devices)
    pltlegend = pltlinebarplot.GetLegend()
    plot_name = f"{devices_name} {f_type} Coord Plot {date_string}"
    plot_settings.title_settings(pltlinebarplot, plot_name)
    pltovercurrent = pltlinebarplot.GetDataSource()
    plot_settings.setup_toc_plot(pltovercurrent, f_type)

    # Add fault level markers
    _add_fault_markers(pltlinebarplot, devices, devices_name, f_type)

    # Add downstream transformer fuse
    ds_fuse = _add_transformer_fuse(
        app, pltlinebarplot, devices, sys_volts, f_type
    )

    # Add all device curves to plot
    us_devices = [d.obj for d in devices[0].us_devices]
    ds_devices = [d.obj for d in devices[0].ds_devices]
    all_devices = [d.obj for d in devices] + us_devices + ds_devices + ds_fuse

    for device in all_devices:
        if device.GetClassName() == 'ElmRelay':
            colour = colour_dic[device]
        else:
            colour = 10  # Fuse colour

        # Relays are added to the pltovercurrent object
        pltovercurrent.AddCurve(device, colour, 1, 80.0)


def create_plot_folder(
    feeder: dd.Feeder,
    graphics_board: pft.SetDesktop
) -> pft.IntFolder:
    """
    Create timestamped folder for coordination plots.

    Args:
        feeder: Feeder dataclass for naming.
        graphics_board: Graphics board to create folder in.

    Returns:
        IntFolder object for storing plots.
    """
    date_string = time.strftime("%Y%m%d-%H%M%S")
    title = f'{date_string} {feeder.obj.loc_name} Prot Coord Plots'

    return graphics_board.CreateObject("IntFolder", title)


def _add_fault_markers(
    plot,
    devices: List[dd.Device],
    devices_name: str,
    f_type: str
) -> None:
    """
    Add minimum and maximum fault level marker lines to plot.

    Args:
        plot: Target VisOcplot object.
        devices: List of devices for fault level extraction.
        devices_name: Combined device name string.
        f_type: 'Ground' or 'Phase' fault type.
    """
    min_fl = plot.CreateObject("VisXvalue", f'{devices_name} min fl')
    max_fl = plot.CreateObject("VisXvalue", f'{devices_name} max fl')

    if f_type == 'Ground':
        min_fl_value = min(d.min_fl_pg for d in devices)
        max_fl_value = max(d.max_fl_pg for d in devices)
        plot_settings.xvalue_settings(min_fl, 'PG Min FL', min_fl_value)
        plot_settings.xvalue_settings(max_fl, 'PG Max FL', max_fl_value)
    else:
        min_fl_value = min(d.min_fl_2ph for d in devices)
        max_fl_value = max(
            max(d.max_fl_3ph, d.max_fl_2ph) for d in devices
        )
        plot_settings.xvalue_settings(min_fl, "Ph Min FL", min_fl_value)
        plot_settings.xvalue_settings(max_fl, "Ph Max FL", max_fl_value)


def _add_transformer_fuse(
    app: pft.Application,
    plot,
    devices: List[dd.Device],
    sys_volts: str,
    f_type: str
) -> List:
    """
    Add downstream transformer fuse curve and marker to plot.

    Args:
        app: PowerFactory application instance.
        plot: Target VisOcplot object.
        devices: List of devices to find max downstream transformer.
        sys_volts: System voltage string for fuse selection.
        f_type: 'Ground' or 'Phase' fault type.

    Returns:
        List containing the fuse element, or empty list if none.
    """
    # Find device with downstream transformer
    device = next(
        (d for d in devices if d.max_ds_tr is not None),
        None
    )

    if device is None:
        return []

    max_ds_tr = device.max_ds_tr
    if max_ds_tr.term is None:
        return []

    # Add transformer fault level marker
    tr_term = max_ds_tr.term
    tr_name = tr_term.cpSubstat.loc_name
    max_tr_fl = plot.CreateObject("VisXvalue", f'{tr_name} max fl')

    if f_type == 'Ground':
        plot_settings.xvalue_settings(max_tr_fl, 'DS TR PG Max FL', max_ds_tr.max_pg)
    else:
        plot_settings.xvalue_settings(max_tr_fl, "DS TR Ph Max FL", max_ds_tr.max_ph)

    # Create/get fuse element
    tr_term_dataclass = next(
        (t for t in device.sect_terms if t.obj == max_ds_tr.term),
        None
    )

    if tr_term_dataclass is None:
        return []

    ds_fuse = ds.create_fuse(app, max_ds_tr, tr_term_dataclass, sys_volts)

    if not ds_fuse:
        app.PrintPlain(
            f'Could not find fuse element for {tr_name} in PowerFactory'
        )
        return []

    return ds_fuse
