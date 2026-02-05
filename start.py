"""
Main entry point for PowerFactory protection assessment automation.

This script orchestrates protection assessment studies for distribution
feeders in DIgSILENT PowerFactory. It provides a unified interface for
multiple study types and coordinates the analysis workflow.

Available Studies:
    - Find PowerFactory Project: Locate project by substation acronym
    - Find Feeder Open Points: Identify normally-open switches
    - Fault Level Study (Legacy): SEQ-specific legacy output format
    - Fault Level Study (No Relays): Fault calculations at switches
    - Fault Level Study (All Relays): Full protection assessment

Optional Assessments (with full relay study):
    - Conductor Damage Assessment: I²t thermal withstand evaluation
    - Protection Coordination Plots: Time-overcurrent curves

Workflow:
    1. User selects study type via GUI dialog
    2. Model validation checks run on all relays
    3. User selects feeders and devices to study
    4. Fault studies execute for each feeder
    5. Optional assessments run on selected devices
    6. Results exported to Excel and PowerFactory graphics

Usage:
    Map this script to a PowerFactory script object and execute from
    within PowerFactory with an active study case selected.

Functions:
    main: Primary orchestration function
    switch_study_case: Activate study case with grid configuration
    cvrt_fdr_to_dataclass: Convert selection to domain dataclasses
"""

import logging
import logging.config
import time
from typing import Dict, List

import powerfactory as pf

from pf_config import pft
import pf_protection_helper as helper
import model_checks
from fault_study import fault_level_study as fs
from fdr_open_points import get_open_points as gop
from find_substation import find_sub
import domain as dd
from user_inputs import get_inputs as gi, study_selection as ss
from legacy_script import script_bridge as sb
from relays import elements
from cond_damage import conductor_damage as cd
from save_results import save_result as sr
from config_logging import configure_logging as cl
from colour_maps import colour_maps as cm
from oc_plots import plot_relay
from importlib import reload

reload(model_checks)
reload(gi)
reload(gop)
reload(fs)
reload(elements)
reload(cd)
reload(sr)
reload(ss)
reload(sb)
reload(find_sub)
reload(cm)
reload(plot_relay)


def main(app: pft.Application) -> None:
    """
    Primary orchestration function for protection assessment.

    Coordinates the complete protection assessment workflow based on
    user selections. Handles study type routing, model validation,
    fault analysis, and result generation.

    Args:
        app: PowerFactory application instance.

    Workflow:
        1. Display study selection dialog
        2. Route to appropriate study handler based on selection
        3. For fault studies:
           a. Run model validation checks
           b. Get user input for feeders and devices
           c. Convert selections to domain dataclasses
           d. Execute fault studies per feeder
           e. Run optional assessments
           f. Generate outputs (Excel, plots, colour maps)

    Side Effects:
        - Modifies PowerFactory study case and calculation state
        - Creates Excel output files
        - Creates graphics board plots (if selected)
        - Creates colour scheme definitions (if applicable)

    Example:
        >>> with helper.app_manager(app, gui=True) as app:
        ...     main(app)
    """


    study_case = app.GetActiveStudyCase()
    graphics_board = app.GetFromStudyCase("Graphics Board.SetDesktop")

    # Get study type selection from user
    study_selections = ss.get_study_selections(app)

    # Route: Find PowerFactory Project
    if "Find PowerFactory Project" in study_selections:
        find_sub.get_project(app)
        return

    # Verify active project
    prjt = app.GetActiveProject()
    if prjt is None:
        app.PrintWarn("No Active Project, Ending Script")
        return

    # Route: Find Feeder Open Points
    if "Find Feeder Open Points" in study_selections:
        gop.main(app)
        return

    # Run model validation checks
    all_relays = elements.get_all_relays(app)
    model_checks.relay_checks(app, all_relays)

    # Get region and user inputs
    region = helper.obtain_region(app)
    feeders_devices, bu_devices, user_selection, external_grid = gi.get_input(
        app, region, study_selections
    )

    # Convert to domain dataclasses
    feeders = cvrt_fdr_to_dataclass(app, feeders_devices, bu_devices)

    # Activate all grids to load floating terminals
    user_selected_study_case = app.GetActiveStudyCase()
    switch_study_case(app, user_selected_study_case, all_grids=True)
    switch_study_case(app, user_selected_study_case, all_grids=False)

    # Process each feeder
    for feeder in feeders:
        gop.get_open_points(app, feeder)
        fs.fault_study(
            app, external_grid, region, feeder, study_selections
        )

        # Run optional assessments on selected devices
        if user_selection:
            selected_devices = [
                device for device in feeder.devices
                if device.obj in user_selection[feeder.obj.loc_name]
            ]

            if 'Conductor Damage Assessment' in study_selections:
                cd.cond_damage(app, selected_devices)

            if 'Protection Relay Coordination Plot' in study_selections:
                plot_relay.plot_all_relays(app, feeder, selected_devices)

    # Generate outputs
    if "Fault Level Study (legacy)" in study_selections:
        sb.bridge_results(app, external_grid, feeders)
    else:
        cm.colour_map(app, region, feeders, study_selections)
        sr.save_dataframe(
            app, region, study_selections, external_grid, feeders
        )


def switch_study_case(
    app: pft.Application,
    user_selected_study_case: pft.IntCase,
    all_grids: bool = False
) -> None:
    """
    Activate a study case with optional all-grids configuration.

    Used to temporarily activate the 'All Active Grids Study Case' to
    ensure floating terminals are properly loaded into line connected
    elements before reverting to the user's selected study case.

    Args:
        app: PowerFactory application instance.
        user_selected_study_case: The user's originally selected IntCase.
        all_grids: If True, activate 'All Active Grids Study Case'.
            If False, activate the user's selected study case.

    Raises:
        SystemExit: If no study case is selected (user_selected_study_case
            is None).

    Note:
        The 'All Active Grids Study Case' must exist in the project's
        study folder for all_grids=True to work.
    """
    if user_selected_study_case is None:
        app.PrintError('Please select a study case and re-run the script.')
        exit(1)

    if all_grids:
        study_folder = app.GetProjectFolder("study")
        int_case = study_folder.GetContents("All Active Grids Study Case")[0]
    else:
        int_case = user_selected_study_case

    int_case.Activate()


def cvrt_fdr_to_dataclass(
    app: pft.Application,
    feeders_devices: Dict,
    bu_devices: Dict
) -> List[dd.Feeder]:
    """
    Convert PowerFactory element selections to domain dataclasses.

    Transforms the dictionaries of PowerFactory objects returned by
    user input collection into structured Feeder and Device dataclasses
    for use in the analysis workflow.

    Args:
        app: PowerFactory application instance.
        feeders_devices: Dictionary mapping feeder names to lists of
            protection device PowerFactory objects.
        bu_devices: Dictionary mapping external grid objects to lists
            of backup device PowerFactory objects.

    Returns:
        List of Feeder dataclasses with devices and bu_devices populated.

    Example:
        >>> feeders = cvrt_fdr_to_dataclass(app, fdr_devs, bu_devs)
        >>> for feeder in feeders:
        ...     print(f"{feeder.obj.loc_name}: {len(feeder.devices)} devices")
    """
    # Convert backup devices to dataclasses
    if bu_devices:
        for grid, grid_devices in bu_devices.items():
            bu_devices[grid] = [
                dd.initialise_dev_dataclass(device)
                for device in grid_devices
            ]

    # Convert feeders and their devices to dataclasses
    feeders = []

    for fdr, devs in feeders_devices.items():
        feeder_obj = app.GetCalcRelevantObjects(fdr + ".ElmFeeder")[0]
        feeder = dd.initialise_fdr_dataclass(feeder_obj)

        devices = [dd.initialise_dev_dataclass(dev) for dev in devs]
        feeder.devices = devices
        feeder.bu_devices = bu_devices

        feeders.append(feeder)

    return feeders


# =============================================================================
# SCRIPT EXECUTION
# =============================================================================

if __name__ == '__main__':
    start = time.time()

    # Configure logging
    logging.basicConfig(
        filename=cl.getpath() / 'prot_assess_log.txt',
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )

    app = pf.GetApplication()

    with helper.app_manager(app, gui=True) as app:
        main(app)

    end = time.time()
    run_time = round(end - start, 6)
    app.PrintPlain(f"Script run time: {run_time} seconds")