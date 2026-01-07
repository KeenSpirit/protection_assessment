"""
Substation-to-project mapping utility for PowerFactory.

This module provides functionality to locate which PowerFactory master
project contains a given substation based on its acronym. It searches
through the project database to find study cases matching the provided
substation identifier.

Supported Regions:
    - SEQ: South East Queensland models
    - Regional North: Northern regional models
    - Regional South: Southern regional models

Functions:
    get_project: Main entry point for project lookup
    sub_selection: Display GUI for region and substation input
    all_substations: Build mapping of projects to substations
    find_project: Search for project containing substation
"""

import sys
import tkinter as tk
from tkinter import ttk
from typing import Dict, List, Optional

from pf_config import pft


def get_project(app: pft.Application) -> None:
    """
    Find the PowerFactory project containing a given substation.

    Main entry point for the substation lookup utility. Displays
    a GUI for user input, searches the project database, and prints
    the result to the PowerFactory output window.

    Args:
        app: PowerFactory application instance.

    Side Effects:
        Prints result message to PowerFactory output window.

    Example:
        >>> get_project(app)
        Substation ABC belongs to PowerFactory project Abermain.
    """
    region, substation = sub_selection(app)
    projects_subs = all_substations(app, region)
    project = find_project(projects_subs, substation)

    if project is not None:
        app.PrintPlain(
            f"Substation {substation} belongs to PowerFactory "
            f"project {project}."
        )
    else:
        app.PrintPlain(
            f"A matching project for substation {substation} "
            f"could not be found."
        )


def sub_selection(app: pft.Application) -> tuple:
    """
    Display GUI dialog for region and substation input.

    Creates a dialog with radio buttons for region selection and
    a text entry for the substation acronym. Validates input
    before returning.

    Args:
        app: PowerFactory application instance.

    Returns:
        Tuple of (region_string, substation_acronym).

    Validation Rules:
        - Substation must be non-empty
        - Only alphabetical characters allowed
        - SEQ: Maximum 3 characters
        - Regional: Maximum 4 characters

    Example:
        >>> region, sub = sub_selection(app)
        >>> print(f"Region: {region}, Substation: {sub}")
    """
    root = tk.Tk()
    root.title("Find Project")
    root.geometry("+800+200")

    # Region selection
    ttk.Label(
        root,
        text="Select Region:",
        font='Helvetica 14 bold'
    ).grid(row=0, column=0, columnspan=3, sticky="w", padx=5, pady=5)

    selection = tk.StringVar(value="0")

    ttk.Radiobutton(
        root, text="SEQ", variable=selection, value="0"
    ).grid(row=1, column=0, sticky="w", padx=5, pady=2)

    ttk.Radiobutton(
        root, text="Regional Northern", variable=selection, value="1"
    ).grid(row=2, column=0, sticky="w", padx=5, pady=2)

    ttk.Radiobutton(
        root, text="Regional Southern", variable=selection, value="2"
    ).grid(row=3, column=0, sticky="w", padx=5, pady=2)

    # Substation entry
    name_var = tk.StringVar()
    ttk.Label(
        root,
        text="Enter the Substation Acronym:",
        font='Helvetica 14 bold'
    ).grid(row=4, columnspan=3, sticky="w", padx=5, pady=(15, 5))

    ttk.Entry(
        root, textvariable=name_var, font=('calibre', 10, 'normal')
    ).grid(row=5, columnspan=3, sticky="w", padx=5, pady=5)

    # Input criteria frame
    criteria_frame = ttk.LabelFrame(root, text="Input Criteria", padding=(10, 5))
    criteria_frame.grid(row=6, column=0, columnspan=3, sticky="ew", padx=5, pady=10)

    criteria_text = [
        "• The substation must have a PowerFactory study case assigned",
        "• Only alphabetical characters are allowed (case insensitive)",
        "• The input for SEQ must be three characters long",
        "• The input for Regional must be four characters long"
    ]

    for i, text in enumerate(criteria_text):
        ttk.Label(
            criteria_frame, text=text, font=('calibre', 10, 'normal')
        ).grid(row=i, column=0, sticky="w", pady=2)

    ttk.Label(
        root,
        text="The result will be displayed in the Output Window.",
        font=('calibre', 10, 'normal')
    ).grid(row=7, column=0, columnspan=3, sticky="w", padx=5, pady=5)

    # Buttons
    ttk.Button(
        root, text='Okay', command=lambda: root.destroy()
    ).grid(row=8, column=0, sticky="w", padx=5, pady=5)

    ttk.Button(
        root, text='Exit', command=lambda: exit_script(root, app)
    ).grid(row=8, column=1, sticky="w", padx=5, pady=5)

    root.mainloop()

    # Map selection to region string
    region_mapping = {
        "0": "SEQ",
        "1": "Regional North",
        "2": "Regional South"
    }
    region = region_mapping[selection.get()]

    substation = name_var.get()

    # Validate input
    if len(substation) < 1:
        return error_message(app, 2)

    if not substation.isalpha():
        return error_message(app, 3)

    is_seq = region == "SEQ"
    is_regional = region in ("Regional North", "Regional South")

    if (is_seq and len(substation) > 3) or (is_regional and len(substation) > 4):
        return error_message(app, 1)

    return region, substation.upper()


def all_substations(app: pft.Application, region: str) -> Dict[str, List[str]]:
    """
    Build mapping of projects to their substation acronyms.

    Traverses the project database for the specified region and
    extracts substation acronyms from study case names.

    Args:
        app: PowerFactory application instance.
        region: Region string ('SEQ', 'Regional North', 'Regional South').

    Returns:
        Dictionary mapping project names to lists of substation acronyms.
        Example: {"Abermain": ["ABY", "BDB"]}

    Note:
        Study case names are expected to have the substation acronym
        as the first space-separated token.
    """
    projects_subs = {}

    user = app.GetCurrentUser()
    database = user.GetAttribute('fold_id')
    models = None

    # Navigate to appropriate region folder
    if region == "Regional North":
        path = "Publisher\\MasterProjects\\Regional Models\\Northern"
        models = database.GetContents(path)[0]
    elif region == "Regional South":
        path = "Publisher\\MasterProjects\\Regional Models\\Southern"
        models = database.GetContents(path)[0]
    elif region == "SEQ":
        path = "Publisher\\MasterProjects\\SEQ Models"
        models = database.GetContents(path)[0]

    app.PrintPlain(
        f"Obtaining a list of all {region} Master Project models..."
    )

    if models is None:
        app.PrintPlain(
            f"Could not obtain a list of all {region} Master Project models."
        )
        sys.exit(0)

    projects = models.GetContents("*.IntPrj")

    # Extract substations from each project's study cases
    for project in projects:
        study_cases = project.GetContents("Study Cases")
        intcases = []

        for study_case in study_cases:
            sub_study_cases = study_case.GetContents("Substation Study Cases")

            for sub_study_case in sub_study_cases:
                intcases.extend([
                    intcase.GetAttribute('loc_name').split()[0]
                    for intcase in sub_study_case.GetContents("*.IntCase")
                ])

        projects_subs[project.GetAttribute('loc_name')] = intcases

    return projects_subs


def find_project(
    projects_subs: Dict[str, List[str]],
    substation: str
) -> Optional[str]:
    """
    Find the project containing a given substation acronym.

    Searches through the project-to-substations mapping to find
    which project contains the specified substation.

    Args:
        projects_subs: Dictionary mapping project names to substation lists.
        substation: Substation acronym to search for (case insensitive).

    Returns:
        Project name if found, None otherwise.

    Example:
        >>> project = find_project(mapping, "ABY")
        >>> print(project)
        'Abermain'
    """
    substation = substation.upper()

    for project, sub_list in projects_subs.items():
        if substation in sub_list:
            return project

    return None


def error_message(app: pft.Application, error_code: int) -> tuple:
    """
    Display error message and re-prompt for input.

    Args:
        app: PowerFactory application instance.
        error_code: Error type indicator:
            1 = Input too long
            2 = Empty input
            3 = Non-alphabetical characters

    Returns:
        Tuple of (region, substation) from re-prompted dialog.
    """
    if error_code == 1:
        app.PrintPlain(
            "Please input 3 characters or less for SEQ, "
            "or 4 characters or less for Regional"
        )
    elif error_code == 2:
        app.PrintPlain("Please input a substation acronym to proceed")
    else:
        app.PrintPlain("Please input alphabetical characters only")

    return sub_selection(app)


def exit_script(root: tk.Tk, app: pft.Application) -> None:
    """
    Clean exit handler for GUI dialogs.

    Args:
        root: Tkinter root window to destroy.
        app: PowerFactory application instance.
    """
    app.PrintPlain("User terminated script.")
    root.destroy()
    sys.exit(0)