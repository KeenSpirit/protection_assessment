"""
Study type selection dialog for protection assessment.

This module provides the initial GUI dialog for users to select
which type of protection study to perform. It presents radio button
options for different study types and conditional checkboxes for
additional analyses.

Available Study Types:
    - Find PowerFactory Project: Locate project by substation acronym
    - Find Feeder Open Points: Identify normally open points
    - Fault Level Study (legacy): SEQ legacy output format
    - Fault Level Study (no relays): Assessment without relay models
    - Fault Level Study (all relays): Full protection assessment

Additional Options (with 'all relays' only):
    - Conductor Damage Assessment: Evaluate conductor thermal limits
    - Protection Relay Coordination Plot: Generate time-overcurrent
      coordination plots

Functions:
    get_study_selections: Main entry point for study selection dialog.
    exit_script: Clean exit handler for the GUI.
    no_study_selected: Re-prompt handler for empty selections.

Example:
    >>> from user_inputs import study_selection
    >>> selections = study_selection.get_study_selections(app)
    >>> if 'Conductor Damage Assessment' in selections:
    ...     run_conductor_damage_study()
"""

import sys
import tkinter as tk
from tkinter import ttk
from typing import List

from pf_config import pft


def get_study_selections(app: pft.Application) -> List[str]:
    """
    Display study selection dialog and collect user choices.

    Creates a tkinter dialog with radio buttons for study type
    selection and conditional checkboxes for additional analyses.
    The checkboxes are only enabled when 'All relays configured'
    study type is selected.

    Args:
        app: PowerFactory application instance for output messaging.

    Returns:
        List of selected study options as strings. The first element
        is always the study type, followed by any additional options.

        Possible study types:
            - "Find PowerFactory Project"
            - "Find Feeder Open Points"
            - "Fault Level Study (legacy)"
            - "Fault Level Study (no relays configured in model)"
            - "Fault Level Study (all relays configured in model)"

        Possible additional options (only with 'all relays'):
            - "Conductor Damage Assessment"
            - "Protection Relay Coordination Plot"

        Returns empty list via recursive call if no selection made.

    Example:
        >>> selections = get_study_selections(app)
        >>> print(selections)
        ['Fault Level Study (all relays configured in model)',
         'Conductor Damage Assessment']
    """
    root = tk.Tk()
    root.title("Protection Assessment")
    root.geometry("+800+200")

    selection = tk.StringVar()
    selection.set('var')

    conductor_damage_var = tk.BooleanVar(value=False)
    coordination_plot_var = tk.BooleanVar(value=False)

    ttk.Label(
        root,
        text="Select study to undertake:",
        font='Helvetica 14 bold'
    ).grid(row=0, columnspan=3, sticky="w", padx=5, pady=5)

    tk.Radiobutton(
        root,
        text="Find PowerFactory Project",
        value="0",
        variable=selection,
        command=lambda: on_radio_change()
    ).grid(row=1, column=0, sticky="w", padx=30, pady=5)

    tk.Radiobutton(
        root,
        text="Find Feeder Open Points",
        value="1",
        variable=selection,
        command=lambda: on_radio_change()
    ).grid(row=2, column=0, sticky="w", padx=30, pady=5)

    tk.Radiobutton(
        root,
        text="Fault Level Study (SEQ legacy script)",
        value="2",
        variable=selection,
        command=lambda: on_radio_change()
    ).grid(row=3, column=0, sticky="w", padx=30, pady=5)

    tk.Radiobutton(
        root,
        text="Fault Level Study (No relays configured in model)",
        value="3",
        variable=selection,
        command=lambda: on_radio_change()
    ).grid(row=4, column=0, sticky="w", padx=30, pady=5)

    tk.Radiobutton(
        root,
        text="Fault Level Study (All relays configured in model)",
        value="4",
        variable=selection,
        command=lambda: on_radio_change()
    ).grid(row=5, column=0, sticky="w", padx=30, pady=5)

    conductor_damage_cb = tk.Checkbutton(
        root,
        text="Conductor Damage Assessment",
        variable=conductor_damage_var,
        state='disabled'
    )
    conductor_damage_cb.grid(row=6, column=0, sticky="w", padx=50, pady=2)

    coordination_plot_cb = tk.Checkbutton(
        root,
        text="Protection Relay Coordination Plot",
        variable=coordination_plot_var,
        state='disabled'
    )
    coordination_plot_cb.grid(row=7, column=0, sticky="w", padx=50, pady=2)

    def on_radio_change():
        """Enable/disable checkboxes based on radio button selection."""
        if selection.get() == "4":
            conductor_damage_cb.config(state='normal')
            coordination_plot_cb.config(state='normal')
        else:
            conductor_damage_var.set(False)
            coordination_plot_var.set(False)
            conductor_damage_cb.config(state='disabled')
            coordination_plot_cb.config(state='disabled')

    info_frame = ttk.LabelFrame(root, text="Information", padding=(10, 5))
    info_frame.grid(row=8, column=0, columnspan=3, sticky="ew", padx=10, pady=10)

    info_text = (
        "At script completion, the location of saved study results files\n"
        "will be displayed in the PowerFactory output window."
    )
    ttk.Label(info_frame, text=info_text, font=('Helvetica', 9)).pack(
        anchor="w"
    )

    ttk.Button(
        root, text='Okay', command=lambda: root.destroy()
    ).grid(row=9, column=0, sticky="w", padx=5, pady=5)

    ttk.Button(
        root, text='Exit', command=lambda: exit_script(root, app)
    ).grid(row=9, column=1, sticky="w", padx=5, pady=5)

    root.mainloop()

    study_mapping = {
        "0": "Find PowerFactory Project",
        "1": "Find Feeder Open Points",
        "2": "Fault Level Study (legacy)",
        "3": "Fault Level Study (no relays configured in model)",
        "4": "Fault Level Study (all relays configured in model)"
    }

    result = []

    if selection.get() != 'var':
        study_name = study_mapping.get(selection.get())
        if study_name:
            result.append(study_name)

        if selection.get() == "4":
            if conductor_damage_var.get():
                result.append("Conductor Damage Assessment")
            if coordination_plot_var.get():
                result.append("Protection Relay Coordination Plot")
    else:
        result = no_study_selected(app)

    return result


def exit_script(root: tk.Tk, app: pft.Application) -> None:
    """
    Clean exit handler for the study selection dialog.

    Prints termination message to PowerFactory output, destroys
    the dialog window, and exits the script.

    Args:
        root: The tkinter root window to destroy.
        app: PowerFactory application instance for output messaging.
    """
    app.PrintPlain("User terminated script.")
    root.destroy()
    sys.exit(0)


def no_study_selected(app: pft.Application) -> List[str]:
    """
    Handle case when user closes dialog without selecting a study.

    Prints a prompt message and recursively calls the selection
    dialog to ensure the user makes a selection.

    Args:
        app: PowerFactory application instance for output messaging.

    Returns:
        List of selected study options from the re-prompted dialog.
    """
    app.PrintPlain("Please select a study to continue")
    study = get_study_selections(app)
    return study