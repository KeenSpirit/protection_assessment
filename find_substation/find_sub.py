import tkinter as tk
from tkinter import ttk
import powerfactory as pf
from typing import List, Dict, Union
import sys

from pf_config import pft


def get_project(app: pft.Application):
    """
    Returns the Powerfactory project belonging to a given Energex substation acronym.
    """

    region, substation = sub_selection(app)
    projects_subs = all_substations(app, region)
    projects = find_project(projects_subs, substation)
    if projects is not None:
        app.PrintPlain(f"Substation {substation} belongs to PowerFactory project {projects}.")
    else:
        app.PrintPlain(f"A matching project for substation {substation} could not be found.")


def sub_selection(app: pft.Application):
    """User inputs a substation acronym and selects a region"""

    root = tk.Tk()
    root.title("Find Project")
    root.geometry("+800+200")

    # Region selection label
    ttk.Label(root, text="Select Region:", font='Helvetica 14 bold').grid(
        row=0, column=0, columnspan=3, sticky="w", padx=5, pady=5
    )

    # Radio button variable and options
    selection = tk.StringVar(value="0")  # Default to SEQ

    ttk.Radiobutton(root, text="SEQ", variable=selection, value="0").grid(
        row=1, column=0, sticky="w", padx=5, pady=2
    )
    ttk.Radiobutton(root, text="Regional Northern", variable=selection, value="1").grid(
        row=2, column=0, sticky="w", padx=5, pady=2
    )
    ttk.Radiobutton(root, text="Regional Southern", variable=selection, value="2").grid(
        row=3, column=0, sticky="w", padx=5, pady=2
    )

    # Substation entry
    name_var = tk.StringVar()
    ttk.Label(root, text="Enter the Substation Acronym:", font='Helvetica 14 bold').grid(
        row=4, columnspan=3, sticky="w", padx=5, pady=(15, 5)
    )
    ttk.Entry(root, textvariable=name_var, font=('calibre', 10, 'normal')).grid(
        row=5, columnspan=3, sticky="w", padx=5, pady=5
    )

    # Input Criteria frame with border
    criteria_frame = ttk.LabelFrame(root, text="Input Criteria", padding=(10, 5))
    criteria_frame.grid(row=6, column=0, columnspan=3, sticky="ew", padx=5, pady=10)

    ttk.Label(criteria_frame, text="• The substation must have a PowerFactory study case assigned",
              font=('calibre', 10, 'normal')).grid(row=0, column=0, sticky="w", pady=2)
    ttk.Label(criteria_frame, text="• Only alphabetical characters are allowed (case insensitive)",
              font=('calibre', 10, 'normal')).grid(row=1, column=0, sticky="w", pady=2)
    ttk.Label(criteria_frame, text="• The input for SEQ must be three characters long",
              font=('calibre', 10, 'normal')).grid(row=2, column=0, sticky="w", pady=2)
    ttk.Label(criteria_frame, text="• The input for Regional must be four characters long",
              font=('calibre', 10, 'normal')).grid(row=3, column=0, sticky="w", pady=2)

    ttk.Label(root, text="The result will be displayed in the Output Window.",
              font=('calibre', 10, 'normal')).grid(row=7, column=0, columnspan=3, sticky="w", padx=5, pady=5)

    # Buttons
    ttk.Button(root, text='Okay', command=lambda: root.destroy()).grid(
        row=8, column=0, sticky="w", padx=5, pady=5
    )
    ttk.Button(root, text='Exit', command=lambda: exit_script(root, app)).grid(
        row=8, column=1, sticky="w", padx=5, pady=5
    )

    # Run the interface
    root.mainloop()

    # Map radio selection to region string
    region_mapping = {
        "0": "SEQ",
        "1": "Regional North",
        "2": "Regional South"
    }
    region = region_mapping[selection.get()]

    # Collect results and validate input
    substation = name_var.get()

    if len(substation) < 1:
        region, substation = error_message(app, 2)
        return region, substation
    elif not substation.isalpha():
        region, substation = error_message(app, 3)
        return region, substation
    elif (region == "SEQ" and len(substation) > 3) or (
            region in ("Regional North", "Regional South") and len(substation) > 4):
        region, substation = error_message(app, 1)
        return region, substation
    else:
        return region, substation.upper()


def exit_script(root, app):
    """Exits script"""

    app.PrintPlain("User terminated script.")
    root.destroy()
    sys.exit(0)


def error_message(app: pft.Application, x: int):
    """Print an error message if input is wrong format"""

    if x == 1:
        app.PrintPlain("Please input 3 characters or less for SEQ, or 4 characters or less for Regional")
    elif x == 2:
        app.PrintPlain("Please input a substation acronym to proceed")
    else:
        app.PrintPlain("Please input alphabetical characters only")
    region, study = sub_selection(app)
    return region, study


def all_substations(app: pft.Application, region: str):
    """
    Returns a dictionary mapping project names to their substation acronyms.
    Example: {"Abermain": ["ABY", "BDB"]}
    """
    projects_subs = {}

    user = app.GetCurrentUser()
    database = user.GetAttribute('fold_id')
    models = None

    if region == "Regional North":
        models = database.GetContents("Publisher\\MasterProjects\\Regional Models\\Northern")[0]
    elif region == "Regional South":
        models = database.GetContents("Publisher\\MasterProjects\\Regional Models\\Southern")[0]
    elif region == "SEQ":
        models = database.GetContents("Publisher\\MasterProjects\\SEQ Models")[0]

    app.PrintPlain(f"Obtaining a list of all {region} Master Project models...")

    if models is not None:
        projects = models.GetContents("*.IntPrj")
    else:
        app.PrintPlain(f"Could not obtain a list of all {region} Master Project models.")
        sys.exit(0)

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


def find_project(projects_subs: Dict, substation: str):
    """Find the project containing the given substation acronym"""

    substation = substation.upper()
    for project, sub_list in projects_subs.items():
        if substation in sub_list:
            return project
    return None