import powerfactory as pf
import tkinter as tk
from tkinter import ttk
from typing import List
import sys
sys.path.append(r"\\Ecasd01\WksMgmt\PowerFactory\ScriptsDEV\PowerFactoryTyping")
import powerfactorytyping as pft



def get_study_selections(app: pft.Application) -> List:
    """
    User selects a study to undertake.

    Args:
        app: PowerFactory application object

    Returns:
        list: A list of strings representing the selected study type and any
              additional options. Returns empty list if no selection made.

    Possible return values include:
        - "Find PowerFactory Project"
        - "Find Feeder Open Points"
        - "Protection Assessment (Legacy)" (only available if region == "SEQ")
        - "Protection Assessment (No Relays Configured In Model)"
        - "Protection Assessment (All Relays Configured In Model)"
        - "Conductor Damage Assessment" (only with All Relays option)
        - "Protection Relay Coordination Plot" (only with All Relays option)
    """

    root = tk.Tk()
    root.title("Protection Assessment")
    root.geometry("+800+200")

    # Variables for radio button and checkboxes
    selection = tk.StringVar()
    selection.set('var')

    conductor_damage_var = tk.BooleanVar(value=False)
    coordination_plot_var = tk.BooleanVar(value=False)

    # Header
    ttk.Label(root, text="Select study to undertake:", font='Helvetica 14 bold'). \
        grid(row=0, columnspan=3, sticky="w", padx=5, pady=5)

    # Radio buttons for study selection
    tk.Radiobutton(root, text="Find PowerFactory Project", value="0", variable=selection,
                   command=lambda: on_radio_change()) \
        .grid(row=1, column=0, sticky="w", padx=30, pady=5)
    tk.Radiobutton(root, text="Find Feeder Open Points", value="1", variable=selection,
                   command=lambda: on_radio_change()) \
        .grid(row=2, column=0, sticky="w", padx=30, pady=5)

    tk.Radiobutton(root, text="Fault Level Study (SEQ legacy script)", value="2", variable=selection,
                   command=lambda: on_radio_change()) \
        .grid(row=3, column=0, sticky="w", padx=30, pady=5)

    tk.Radiobutton(root, text="Fault Level Study (No relays configured in model)", value="3", variable=selection,
                   command=lambda: on_radio_change()) \
        .grid(row=4, column=0, sticky="w", padx=30, pady=5)

    tk.Radiobutton(root, text="Fault Level Study (All relays configured in model)", value="4", variable=selection,
                   command=lambda: on_radio_change()) \
        .grid(row=5, column=0, sticky="w", padx=30, pady=5)

    # Conditional checkboxes (disabled by default)
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
        """Enable/disable checkboxes based on radio button selection"""
        if selection.get() == "4":  # "Protection Assessment (All Relays Configured In Model)"
            conductor_damage_cb.config(state='normal')
            coordination_plot_cb.config(state='normal')
        else:
            # Disable and untick checkboxes
            conductor_damage_var.set(False)
            coordination_plot_var.set(False)
            conductor_damage_cb.config(state='disabled')
            coordination_plot_cb.config(state='disabled')

    # Information frame with light border
    info_frame = ttk.LabelFrame(root, text="Information", padding=(10, 5))
    info_frame.grid(row=8, column=0, columnspan=3, sticky="ew", padx=10, pady=10)

    info_text = "At script completion, the location of saved study results files\nwill be displayed in the PowerFactory output window."
    ttk.Label(info_frame, text=info_text, font=('Helvetica', 9)).pack(anchor="w")

    # Buttons
    ttk.Button(root, text='Okay', command=lambda: root.destroy()) \
        .grid(row=9, column=0, sticky="w", padx=5, pady=5)
    ttk.Button(root, text='Exit', command=lambda: exit_script(root, app)) \
        .grid(row=9, column=1, sticky="w", padx=5, pady=5)

    # Run the interface
    root.mainloop()

    # Mapping of radio button values to study names
    study_mapping = {
        "0": "Find PowerFactory Project",
        "1": "Find Feeder Open Points",
        "2": "Fault Level Study (legacy)",
        "3": "Fault Level Study (no relays configured in model)",
        "4": "Fault Level Study (all relays configured in model)"
    }

    # Build the return list based on selections
    result = []

    if selection.get() != 'var':
        # Add the selected study type
        study_name = study_mapping.get(selection.get())
        if study_name:
            result.append(study_name)

        # Add checkbox selections (only relevant for "All Relays" option)
        if selection.get() == "4":
            if conductor_damage_var.get():
                result.append("Conductor Damage Assessment")
            if coordination_plot_var.get():
                result.append("Protection Relay Coordination Plot")
    else:
        # No selection made, prompt user
        result = no_study_selected(app)

    return result


def exit_script(root, app: pft.Application):
    """Exits script"""

    app.PrintPlain("User terminated script.")
    root.destroy()
    sys.exit(0)


def no_study_selected(app: pft.Application) -> List:
    """Ensure the user selects a study before continuing"""

    app.PrintPlain("Please select a study to continue")
    study = get_study_selections(app)
    return study
