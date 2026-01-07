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
    new_page_format: Create A4 landscape page format
    drawing_format: Configure graphics board drawing format
    create_plot_folder: Create timestamped folder for plots
    update_ds_tr_data: Collect RMU fuse specifications for SEQ
    create_colour_dic: Assign unique colours to devices
    create_plot: Generate a single coordination plot
    create_draw_format: Apply drawing format to plot page
    plot_settings: Configure plot axis and display settings
    setocplt: Apply overcurrent plot settings object
    xvalue_settings: Configure fault level marker lines
"""

import math
import time
from typing import Dict, List

from pf_config import pft
from devices import fuses as ds
from oc_plots import get_rmu_fuses as grf
from pf_protection_helper import create_obj, obtain_region
import domain as dd
from importlib import reload


# =============================================================================
# MAIN ENTRY POINT
# =============================================================================

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

    new_format = new_page_format(app)
    study_case = app.GetActiveStudyCase()
    graphics_board = app.GetFromStudyCase("Graphics Board.SetDesktop")

    study_case.Deactivate()
    drawing_format(graphics_board, new_format)
    plot_folder = create_plot_folder(feeder, graphics_board)

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
            vipage = create_plot(
                app, plot_folder, colour_dic, devices,
                feeder.sys_volts, f_type='Ground'
            )
            create_draw_format(vipage)

            # Phase fault plot (if applicable)
            if any(device.min_fl_2ph > 0 for device in devices):
                vipage = create_plot(
                    app, plot_folder, colour_dic, devices,
                    feeder.sys_volts, f_type='Phase'
                )
                create_draw_format(vipage)
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
# PAGE FORMAT SETUP
# =============================================================================

def new_page_format(app: pft.Application) -> pft.SetFormat:
    """
    Create A4 landscape page format in project settings.

    Args:
        app: PowerFactory application instance.

    Returns:
        SetFormat object configured for A4 landscape (297x210mm).
    """
    prjt = app.GetActiveProject()
    settings = prjt.GetContents('Settings.SetFold')[0]
    formats = create_obj(settings, "Page Formats", "SetFoldpage")

    new_format_name = 'A4'
    new_format = create_obj(formats, new_format_name, "SetFormat")
    new_format.iSizeX = 297
    new_format.iSizeY = 210
    new_format.iLeft = 0
    new_format.iRight = 0
    new_format.iTop = 0
    new_format.iBottom = 0

    return new_format


def drawing_format(
    graphics_board: pft.SetDesktop,
    format_graph: pft.SetFormat
) -> None:
    """
    Configure drawing format for the graphics board.

    Args:
        graphics_board: PowerFactory SetDesktop object.
        format_graph: Page format to apply.
    """
    draw_form = create_obj(graphics_board, "Drawing Format", "SetGrfpage")
    draw_form.iDrwFrm = 1  # Landscape
    draw_form.aDrwFrm = format_graph.loc_name


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


def create_draw_format(vipage: pft.SetVipage) -> None:
    """
    Apply drawing format to a plot page.

    Args:
        vipage: Plot page object.
    """
    draw_format = create_obj(vipage, "Drawing Format", "SetGrfpage")
    draw_format.iDrwFrm = 1  # Landscape
    draw_format.aDrwFrm = 'A4'


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

    # Colour 0 is white, start with 1
    colour_dic = {device.obj: i + 1 for i, device in enumerate(unique_devices)}

    return colour_dic


# =============================================================================
# PLOT CREATION
# =============================================================================

def create_plot(
    app: pft.Application,
    graphics_board: pft.SetDesktop,
    colour_dic: Dict,
    devices: List[dd.Device],
    sys_volts: str,
    f_type: str
) -> pft.SetVipage:
    """
    Create a single time-overcurrent coordination plot.

    Generates a VisOcplot containing:
    - Primary device characteristic curves
    - Upstream and downstream device curves
    - Minimum and maximum fault level markers
    - Downstream transformer fuse curve (if applicable)

    Args:
        app: PowerFactory application instance.
        graphics_board: Target folder for plot.
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
    vipage = create_obj(graphics_board, folder_name, "SetVipage")

    plot_name = f"{devices_name} {f_type} Coord Plot {date_string}"
    plot = vipage.CreateObject("VisOcplot", plot_name)
    plot.Clear()

    # Add fault level markers
    _add_fault_markers(plot, devices, devices_name, f_type)

    # Add downstream transformer fuse
    ds_fuse = _add_transformer_fuse(
        app, plot, devices, sys_volts, f_type
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

        plot.AddRelay(device, colour)

    plot_settings(plot, devices[0], f_type)

    return vipage


def _add_fault_markers(
    plot: pft.VisOcplot,
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
        xvalue_settings(min_fl, 'PG Min FL', min_fl_value)
        xvalue_settings(max_fl, 'PG Max FL', max_fl_value)
    else:
        min_fl_value = min(d.min_fl_2ph for d in devices)
        max_fl_value = max(
            max(d.max_fl_3ph, d.max_fl_2ph) for d in devices
        )
        xvalue_settings(min_fl, "Ph Min FL", min_fl_value)
        xvalue_settings(max_fl, "Ph Max FL", max_fl_value)


def _add_transformer_fuse(
    app: pft.Application,
    plot: pft.VisOcplot,
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
        xvalue_settings(max_tr_fl, 'DS TR PG Max FL', max_ds_tr.max_pg)
    else:
        xvalue_settings(max_tr_fl, "DS TR Ph Max FL", max_ds_tr.max_ph)

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


# =============================================================================
# PLOT SETTINGS
# =============================================================================

def plot_settings(
    plot: pft.VisOcplot,
    relay: dd.Device,
    f_type: str
) -> None:
    """
    Configure plot axis ranges and display settings.

    Args:
        plot: VisOcplot object to configure.
        relay: Reference device for fault level bounds.
        f_type: 'Ground' or 'Phase' fault type.

    Settings Applied:
        - X-axis: Log scale, auto-fit to fault range
        - Y-axis: Log scale, 0.01 to 10 seconds
        - Earth fault curves: Dashed line style
    """
    if f_type == 'Ground':
        x_min = _get_bound(relay.min_fl_pg, bound='Min')
        x_max = _get_bound(relay.max_fl_pg, bound='Max')
        # Dashed line style for earth fault curves
        plot.gStyle = [10 for _ in range(len(plot.gStyle))]
    else:
        max_fl_ph = max(relay.max_fl_2ph, relay.max_fl_3ph)
        x_min = _get_bound(relay.min_fl_2ph, bound='Min')
        x_max = _get_bound(max_fl_ph, bound='Max')

    setocplt(plot, f_type)

    plot.x_max = x_max
    plot.x_min = x_min
    plot.x_map = 1  # Log scale
    plot.y_max = 10
    plot.y_min = 0.01
    plot.y_map = 1  # Log scale
    plot.y_fit = 0  # Fixed scale


def _get_bound(num: float, bound: str) -> float:
    """
    Round fault level to nearest order of magnitude.

    Args:
        num: Fault level value.
        bound: 'Min' to round down, 'Max' to round up.

    Returns:
        Rounded value for axis limit.
    """
    order_of_mag = 10 ** int(math.log10(num))

    if bound == 'Min':
        return math.floor(num / order_of_mag) * order_of_mag
    else:
        return math.ceil(num / order_of_mag) * order_of_mag


def setocplt(plot: pft.VisOcplot, f_type: str) -> None:
    """
    Create and configure overcurrent plot settings object.

    Args:
        plot: VisOcplot object to configure.
        f_type: 'Ground' or 'Phase' fault type.
    """
    settings = create_obj(plot, "Overcurrent Plot Settings", "SetOcplt")

    settings.unit = 0  # Primary current
    settings.ishow = 2 if f_type == 'Ground' else 1  # Phase and/or Earth
    settings.ishowminmax = 0  # All characteristics
    settings.ishowdir = 0  # All directions
    settings.ishowtframe = 0  # All recloser operations
    settings.ishowcalc = 1  # Display current automatically
    settings.iTbrk = 0  # No breaker time consideration
    settings.iushow = 0  # All voltage references
    settings.imarg = 1  # Show grading margins


def xvalue_settings(
    constant: pft.VisXvalue,
    name: str,
    value: float
) -> None:
    """
    Configure fault current marker line settings.

    Args:
        constant: VisXvalue object to configure.
        name: Label text for the marker.
        value: Fault current value in Amperes.
    """
    constant.loc_name = name
    constant.label = 1
    constant.lab_text = [name]
    constant.show = 1  # Show with intersections
    constant.iopt_lab = 3  # Label position
    constant.value = value
    constant.color = 1
    constant.width = 5
    constant.xis = 0  # Current axis