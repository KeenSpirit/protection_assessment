"""
Study type selection dialog for protection assessment.

This module provides a GUI dialog for users to select which study types
to execute. The available studies include fault level analysis, protection
coordination, and conductor damage assessment.

Functions:
    get_study_selections: Display study selection dialog and return choices
    exit_script: Clean exit handler for the GUI
"""

import sys
import tkinter as tk
from tkinter import ttk
from typing import List

from pf_config import pft


def get_study_selections(app: pft.Application) -> List[str]:
    """
    Display a dialog for selecting study types to execute.

    Presents a GUI with checkboxes for each available study type.
    The user can select one or more studies to run. Some studies
    are mutually exclusive or have dependencies.

    Available Study Types:
        - Find PowerFactory Project: Locate project by substation
        - Find Feeder Open Points: Identify normally-open points
        - Fault Level Study (legacy): SEQ-specific legacy format
        - Fault Level Study (no relays): Fault study without relay data
        - Fault Level Study (all relays): Full protection assessment
        - Conductor Damage Assessment: Thermal withstand evaluation
        - Protection Relay Coordination Plot: Time-overcurrent curves

    Args:
        app: PowerFactory application instance.

    Returns:
        List of selected study type strings.

    Note:
        Conductor Damage Assessment and Protection Relay Coordination
        Plot are only available when running with relay configuration.

    Example:
        >>> selections = get_study_selections(app)
        >>> if "Fault Level Study (all relays)" in selections:
        ...     # Run full protection assessment
    """
    root = tk.Tk()
    root.title("Protection Assessment - Study Selection")

    # Center the window on screen
    window_width = 450
    window_height = 400
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")

    # Main frame with padding
    main_frame = ttk.Frame(root, padding="10")
    main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    # Header
    header = ttk.Label(
        main_frame,
        text="Select Study Type(s):",
        font='Helvetica 14 bold'
    )
    header.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

    # Study type definitions with descriptions
    study_types = [
        ("Find PowerFactory Project", "Locate project by substation acronym"),
        ("Find Feeder Open Points", "Identify normally-open points on feeders"),
        ("Fault Level Study (legacy)", "SEQ fault study with legacy output"),
        (
            "Fault Level Study (no relays configured in model)",
            "Fault calculations at switch locations"
        ),
        (
            "Fault Level Study (all relays configured in model)",
            "Full protection assessment with relay data"
        ),
        ("Conductor Damage Assessment", "Thermal withstand evaluation"),
        ("Protection Relay Coordination Plot", "Time-overcurrent curves"),
    ]

    # Create checkboxes for each study type
    study_vars = {}
    for i, (study_name, description) in enumerate(study_types):
        var = tk.IntVar()
        study_vars[study_name] = var

        cb = ttk.Checkbutton(
            main_frame,
            text=study_name,
            variable=var
        )
        cb.grid(row=i + 1, column=0, sticky="w", padx=(20, 0), pady=2)

    # Separator
    separator = ttk.Separator(main_frame, orient='horizontal')
    separator.grid(
        row=len(study_types) + 1,
        column=0,
        columnspan=2,
        sticky="ew",
        pady=10
    )

    # Note about optional assessments
    note_text = (
        "Note: Conductor Damage Assessment and Coordination Plots\n"
        "require 'Fault Level Study (all relays configured)' to be selected."
    )
    note_label = ttk.Label(
        main_frame,
        text=note_text,
        font=('TkDefaultFont', 9),
        foreground='gray'
    )
    note_label.grid(
        row=len(study_types) + 2,
        column=0,
        columnspan=2,
        sticky="w",
        pady=(0, 10)
    )

    # Button frame
    button_frame = ttk.Frame(main_frame)
    button_frame.grid(
        row=len(study_types) + 3,
        column=0,
        columnspan=2,
        sticky="e",
        pady=(10, 0)
    )

    ttk.Button(
        button_frame,
        text='Okay',
        command=lambda: root.destroy()
    ).grid(row=0, column=0, padx=5)

    ttk.Button(
        button_frame,
        text='Exit',
        command=lambda: exit_script(root, app)
    ).grid(row=0, column=1, padx=5)

    # Run the dialog
    root.mainloop()

    # Collect selected studies
    selected_studies = [
        study_name
        for study_name, var in study_vars.items()
        if var.get() == 1
    ]

    return selected_studies


def exit_script(root: tk.Tk, app: pft.Application) -> None:
    """
    Clean exit handler for the study selection dialog.

    Prints a termination message to PowerFactory output and exits
    the script cleanly.

    Args:
        root: The tkinter root window to destroy.
        app: PowerFactory application instance for output messages.
    """
    app.PrintPlain("User terminated script.")
    root.destroy()
    sys.exit(0)