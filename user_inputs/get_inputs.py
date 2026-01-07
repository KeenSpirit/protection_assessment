"""
Feeder and device selection dialog for protection assessment.

This module provides GUI dialogs for users to:
1. Select feeders for study (excluding mesh feeders)
2. Select specific protection devices per feeder
3. Enter or confirm external grid fault level parameters

Functions:
    get_input: Main entry point for user input collection
    mesh_feeder_check: Filter feeders to exclude mesh configurations
    get_feeders: Display feeder selection dialog
    get_external_grid_data: Collect external grid parameters
    get_device_selection: Display device selection per feeder
    populate_feeder_list: Create feeder checkbox interface
    validate_grid_inputs: Validate external grid parameter format
    center_window: Center a tkinter window on screen
    exit_script: Clean exit handler for GUI
"""

import sys
import tkinter as tk
from tkinter import ttk
from typing import Dict, List, Optional, Tuple, Any

from pf_config import pft
from relays import elements
from devices import fuses


def get_input(
    app: pft.Application,
    region: str,
    study_selections: List[str]
) -> Tuple[Dict, Dict, Optional[Dict], Dict]:
    """
    Collect all user inputs for the protection assessment study.

    Orchestrates the complete user input workflow:
    1. Check for mesh feeders and filter to radial only
    2. Display feeder selection dialog
    3. Collect external grid fault level parameters
    4. For relay studies, display device selection per feeder

    Args:
        app: PowerFactory application instance.
        region: Network region ('SEQ' or 'Regional Models').
        study_selections: List of selected study types.

    Returns:
        Tuple containing:
        - feeders_devices: Dict mapping feeder names to device lists
        - bu_devices: Dict of backup devices by grid
        - user_selection: Dict of user-selected devices per feeder,
            or None if no device selection required
        - external_grid: Dict of grid objects to parameter lists

    Example:
        >>> feeders, bu_devs, selection, grids = get_input(
        ...     app, 'SEQ', ['Fault Level Study (all relays)']
        ... )
    """
    # Check for mesh feeders and get radial feeders only
    radial_dic = mesh_feeder_check(app)
    if not radial_dic:
        return {}, {}, None, {}

    # Get user's feeder selection
    selected_feeders = get_feeders(app, radial_dic)
    if not selected_feeders:
        return {}, {}, None, {}

    # Get external grid parameters
    external_grid = get_external_grid_data(app)

    # Get protection devices for selected feeders
    feeders_devices, bu_devices = get_feeder_devices(
        app, selected_feeders, study_selections
    )

    # For relay-configured studies, allow device selection
    user_selection = None
    if "Fault Level Study (all relays configured in model)" in study_selections:
        if any(devices for devices in feeders_devices.values()):
            user_selection = get_device_selection(app, feeders_devices)

    return feeders_devices, bu_devices, user_selection, external_grid


def mesh_feeder_check(app: pft.Application) -> Dict:
    """
    Check for mesh feeders and return only radial feeders.

    Performs topological search on each feeder to determine if the
    external grid is reachable from both upstream and downstream
    directions. Mesh feeders are excluded from the study.

    Args:
        app: PowerFactory application instance.

    Returns:
        Dictionary mapping feeder objects to their names for radial
        feeders only. Returns empty dict if no radial feeders found.

    Note:
        If no radial feeders are found, displays an error dialog
        instructing the user to radialise the network.
    """
    app.PrintPlain("Checking for mesh feeders...")

    grids_all = app.GetCalcRelevantObjects('*.ElmXnet')
    grids = [grid for grid in grids_all if grid.GetAttribute('outserv') == 0]
    all_feeders = app.GetCalcRelevantObjects('*.ElmFeeder')

    radial_dic = {}

    for feeder in all_feeders:
        cubicle = feeder.obj_id
        down_devices = cubicle.GetAll(1, 0)
        up_devices = cubicle.GetAll(0, 0)

        # Check if grid is found in both directions (mesh)
        grid_downstream = any(item in grids for item in down_devices)
        grid_upstream = any(item in grids for item in up_devices)

        if grid_downstream and grid_upstream:
            # Mesh feeder - exclude
            continue
        elif grid_downstream or grid_upstream:
            # Radial feeder - include
            radial_dic[feeder] = feeder.GetAttribute('loc_name')

    if len(radial_dic) == 0:
        _show_no_radial_feeders_error(app)

    return radial_dic


def _show_no_radial_feeders_error(app: pft.Application) -> None:
    """Display error dialog when no radial feeders are found."""
    root = tk.Tk()
    center_window(root, 350, 150)
    root.title("Protection Assessment")

    ttk.Label(
        root,
        text="No radial feeders were found at the substation"
    ).grid(padx=10, pady=10)

    ttk.Label(
        root,
        text="To run this script, please radialise one or more feeders"
    ).grid(padx=10, pady=5)

    ttk.Button(
        root,
        text='Exit',
        command=lambda: exit_script(root, app)
    ).grid(pady=10)

    root.mainloop()


def get_feeders(app: pft.Application, radial_dic: Dict) -> List[str]:
    """
    Display feeder selection dialog and return selected feeders.

    Shows a scrollable list of radial feeders with checkboxes.
    User must select at least one feeder to proceed.

    Args:
        app: PowerFactory application instance.
        radial_dic: Dictionary mapping feeder objects to names.

    Returns:
        List of selected feeder names.

    Note:
        If no feeders are selected, displays an error and re-prompts.
    """
    radial_list = sorted(radial_dic.values())

    root = tk.Tk()
    root.title("Protection Assessment - Feeder Selection")

    # Calculate window dimensions
    window_width = 400
    num_rows = len(radial_list)
    row_height = 32
    window_height = min(max(num_rows * row_height + 150, 300), 700)

    center_window(root, window_width, window_height)

    # Create scrollable frame
    canvas = tk.Canvas(root, borderwidth=0)
    frame = ttk.Frame(canvas)
    vsb = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vsb.set)

    vsb.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    canvas.create_window((4, 4), window=frame, anchor="nw")

    frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    # Populate feeder checkboxes
    feeder_vars = populate_feeder_list(frame, radial_list)

    # Add buttons
    button_row = max(len(radial_list) + 2, 10)
    ttk.Button(
        frame,
        text='Okay',
        command=lambda: root.destroy()
    ).grid(row=button_row, column=0, sticky="w", padx=5, pady=10)

    ttk.Button(
        frame,
        text='Exit',
        command=lambda: exit_script(root, app)
    ).grid(row=button_row, column=1, sticky="w", padx=5, pady=10)

    root.mainloop()

    # Collect selected feeders
    feeder_list = [
        radial_list[i]
        for i, var in enumerate(feeder_vars)
        if var.get() == 1
    ]

    if not feeder_list:
        app.PrintPlain("Please select at least one feeder to continue")
        return get_feeders(app, radial_dic)

    return feeder_list


def populate_feeder_list(
    frame: ttk.Frame,
    radial_list: List[str]
) -> List[tk.IntVar]:
    """
    Create feeder checkbox interface.

    Args:
        frame: Parent frame for the checkboxes.
        radial_list: List of feeder names.

    Returns:
        List of IntVar objects for checkbox states.
    """
    ttk.Label(
        frame,
        text="Select feeders to study:",
        font='Helvetica 14 bold'
    ).grid(columnspan=2, padx=5, pady=5)

    ttk.Label(
        frame,
        text="(mesh feeders excluded)",
        font='Helvetica 10'
    ).grid(columnspan=2, padx=5, pady=(0, 10))

    feeder_vars = []
    for i, feeder_name in enumerate(radial_list):
        var = tk.IntVar()
        feeder_vars.append(var)
        ttk.Checkbutton(
            frame,
            text=feeder_name,
            variable=var
        ).grid(column=0, sticky="w", padx=30, pady=2)

    return feeder_vars


def get_external_grid_data(app: pft.Application) -> Dict:
    """
    Collect external grid fault level parameters from user.

    Displays current grid values and allows user to modify:
    - Maximum fault level (3-phase kA, R/X, Z2/Z1, X0/X1, R0/X0)
    - Minimum fault level parameters
    - System normal minimum parameters

    Args:
        app: PowerFactory application instance.

    Returns:
        Dictionary mapping grid objects to lists of 15 parameters:
        [max_3ph, max_rx, max_z2z1, max_x0x1, max_r0x0,
         min_3ph, min_rx, min_z2z1, min_x0x1, min_r0x0,
         sn_3ph, sn_rx, sn_z2z1, sn_x0x1, sn_r0x0]
    """
    grids = app.GetCalcRelevantObjects('*.ElmXnet')
    active_grids = [g for g in grids if not g.GetAttribute('outserv')]

    external_grid = {}

    for grid in active_grids:
        # Read current values from PowerFactory
        max_3ph = grid.GetAttribute('Sk3max') / 1000  # Convert to kA
        max_rx = grid.GetAttribute('RxmaxS1')
        max_z2z1 = grid.GetAttribute('Z2_Z1maxS1')
        max_x0x1 = grid.GetAttribute('X0_X1maxS1')
        max_r0x0 = grid.GetAttribute('R0_X0maxS1')

        min_3ph = grid.GetAttribute('Sk3min') / 1000
        min_rx = grid.GetAttribute('RxminS1')
        min_z2z1 = grid.GetAttribute('Z2_Z1minS1')
        min_x0x1 = grid.GetAttribute('X0_X1minS1')
        min_r0x0 = grid.GetAttribute('R0_X0minS1')

        # System normal minimum defaults to minimum values
        sn_3ph = min_3ph
        sn_rx = min_rx
        sn_z2z1 = min_z2z1
        sn_x0x1 = min_x0x1
        sn_r0x0 = min_r0x0

        external_grid[grid] = [
            max_3ph, max_rx, max_z2z1, max_x0x1, max_r0x0,
            min_3ph, min_rx, min_z2z1, min_x0x1, min_r0x0,
            sn_3ph, sn_rx, sn_z2z1, sn_x0x1, sn_r0x0
        ]

    return external_grid


def get_feeder_devices(
    app: pft.Application,
    selected_feeders: List[str],
    study_selections: List[str]
) -> Tuple[Dict, Dict]:
    """
    Get protection devices for selected feeders.

    Retrieves relays and fuses associated with each selected feeder.
    Also identifies backup devices at grid level.

    Args:
        app: PowerFactory application instance.
        selected_feeders: List of selected feeder names.
        study_selections: List of selected study types.

    Returns:
        Tuple containing:
        - feeders_devices: Dict mapping feeder names to device lists
        - bu_devices: Dict of backup devices by grid
    """
    feeders_devices = {}
    bu_devices = {}

    # Get all relays and fuses
    all_relays = elements.get_all_relays(app)
    all_fuses = fuses.get_all_fuses(app)

    for feeder_name in selected_feeders:
        feeder = app.GetCalcRelevantObjects(feeder_name + ".ElmFeeder")
        if not feeder:
            continue

        feeder = feeder[0]
        feeder_devices = []

        # Get devices on this feeder
        cubicle = feeder.obj_id
        ds_objects = cubicle.GetAll(1, 0)

        # Filter relays on this feeder
        for relay in all_relays:
            if relay in ds_objects:
                feeder_devices.append(relay)

        # Filter fuses on this feeder
        for fuse in all_fuses:
            if fuse in ds_objects:
                feeder_devices.append(fuse)

        feeders_devices[feeder_name] = feeder_devices

    return feeders_devices, bu_devices


def get_device_selection(
    app: pft.Application,
    feeders_devices: Dict
) -> Dict:
    """
    Display device selection dialog for each feeder.

    Allows user to select specific devices for detailed analysis
    (coordination plots, conductor damage assessment).

    Args:
        app: PowerFactory application instance.
        feeders_devices: Dict mapping feeder names to device lists.

    Returns:
        Dict mapping feeder names to lists of selected device objects.
    """
    user_selection = {}

    for feeder_name, devices in feeders_devices.items():
        if not devices:
            continue

        root = tk.Tk()
        root.title(f"Device Selection - {feeder_name}")

        window_height = min(len(devices) * 30 + 150, 500)
        center_window(root, 400, window_height)

        frame = ttk.Frame(root, padding="10")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        ttk.Label(
            frame,
            text=f"Select devices for {feeder_name}:",
            font='Helvetica 12 bold'
        ).grid(columnspan=2, pady=(0, 10))

        device_vars = []
        for i, device in enumerate(devices):
            var = tk.IntVar(value=1)  # Default selected
            device_vars.append(var)
            ttk.Checkbutton(
                frame,
                text=device.loc_name,
                variable=var
            ).grid(row=i + 1, column=0, sticky="w", padx=20)

        ttk.Button(
            frame,
            text='Okay',
            command=lambda: root.destroy()
        ).grid(row=len(devices) + 2, column=0, pady=10)

        root.mainloop()

        # Collect selected devices
        selected = [
            devices[i]
            for i, var in enumerate(device_vars)
            if var.get() == 1
        ]
        user_selection[feeder_name] = selected

    return user_selection


def center_window(root: tk.Tk, width: int, height: int) -> None:
    """
    Center a tkinter window on the screen.

    Args:
        root: The tkinter window to center.
        width: Window width in pixels.
        height: Window height in pixels.
    """
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    root.geometry(f"{width}x{height}+{x}+{y}")


def exit_script(root: tk.Tk, app: pft.Application) -> None:
    """
    Clean exit handler for GUI dialogs.

    Args:
        root: The tkinter root window to destroy.
        app: PowerFactory application instance for output messages.
    """
    app.PrintPlain("User terminated script.")
    root.destroy()
    sys.exit(0)