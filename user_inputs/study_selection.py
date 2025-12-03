import tkinter as tk
from tkinter import ttk
import sys
import os


def get_study_selections(app, region):
    """
    From the user, gather a study selection list. Also prompt the user to set external grid parameters.

    Args:
        app: The application instance
        region: String indicating region type ("Regional" or "SEQ")

    Returns:
        list: Selected study options
    """
    # Create and configure the root window
    root = tk.Tk()
    root.title("Distribution Protection Assessment")

    # Configure window dimensions and position
    width = 615
    height = 600  # Increased height to accommodate new elements
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    root.geometry(f"{width}x{height}+{x}+{y}")
    root.resizable(False, True)

    # Apply a modern theme if available
    try:
        root.tk.call("source", "azure.tcl")
        root.tk.call("set_theme", "light")
    except tk.TclError:
        pass  # Continue with default theme if custom theme isn't available

    # Create main frame
    main_frame = ttk.Frame(root)
    main_frame.pack(fill="both", expand=True, padx=10, pady=10)

    # Study choices available for selection
    studies = [
        "Distribution Fault Level Study",
        "Conductor Damage Assessment",
        "Protection Relay Coordination Plot",
    ]

    # Study variables and GUI elements
    param_vars, study_vars = create_ui_elements(main_frame, studies, region)

    # Add buttons frame at the bottom
    button_frame = ttk.Frame(root)
    button_frame.pack(fill="x", padx=10, pady=(0, 10))

    # Create styled buttons
    style = ttk.Style()
    style.configure("Accent.TButton", font=("Helvetica", 10, "bold"))

    ok_button = ttk.Button(
        button_frame,
        text="OK",
        command=root.destroy,
        style="Accent.TButton",
        width=10
    )
    exit_button = ttk.Button(
        button_frame,
        text="Exit",
        command=lambda: exit_script(root, app),
        width=10
    )

    # Position buttons
    ok_button.pack(side="left", padx=(0, 10))
    exit_button.pack(side="left")

    # Center the window on screen
    root.update_idletasks()
    root.eval(f'tk::PlaceWindow . center')

    # Start the main loop
    root.mainloop()

    # Process selections after window is closed
    selections = get_selected_studies(studies, study_vars, param_vars)
    return selections


def create_ui_elements(parent, studies, region):
    """
    Create the UI elements for the study selection interface.

    Args:
        parent: Parent frame to place elements in
        studies: List of available studies
        region: String indicating region type ("Regional" or "SEQ")

    Returns:
        tuple: (param_vars, study_vars) - Variables tracking checkbox states
    """
    # Header
    header_frame = ttk.Frame(parent)
    header_frame.pack(fill="x", pady=(0, 10))

    header_label = ttk.Label(
        header_frame,
        text="Distribution Protection Assessment",
        font=("Helvetica", 14, "bold")
    )
    header_label.pack(anchor="w")

    # Assessment Parameters Section
    param_frame = ttk.LabelFrame(parent, text="Assessment Parameters")
    param_frame.pack(fill="x", pady=(0, 15))

    # Create parameter checkboxes
    param_vars = []

    param1_var = tk.IntVar()
    param1_checkbox = ttk.Checkbutton(
        param_frame,
        text="All relevant protection devices have been correctly configured",
        variable=param1_var
    )
    param1_checkbox.pack(anchor="w", padx=10, pady=5)
    param_vars.append(param1_var)

    # Only show Check box 2 if region is "SEQ"
    if region == "SEQ":
        param2_var = tk.IntVar()
        param2_checkbox = ttk.Checkbutton(
            param_frame,
            text="Legacy script output",
            variable=param2_var
        )
        param2_checkbox.pack(anchor="w", padx=10, pady=5)
        param_vars.append(param2_var)
    else:
        # Add a placeholder None for Regional mode
        param_vars.append(None)

    # Study Selection Section
    study_frame = ttk.LabelFrame(parent, text="Select All Studies To Undertake:")
    study_frame.pack(fill="x", pady=(0, 15))

    study_vars = []

    # First checkbox is always checked and disabled
    first_var = tk.IntVar(value=1)
    first_checkbox = ttk.Checkbutton(
        study_frame,
        text=studies[0],
        variable=first_var,
        state="disabled"
    )
    first_checkbox.pack(anchor="w", padx=10, pady=5)
    study_vars.append(first_var)

    # Create remaining checkboxes - initially disabled
    second_var = tk.IntVar()
    second_checkbox = ttk.Checkbutton(
        study_frame,
        text=studies[1],
        variable=second_var,
        state="disabled"  # Start disabled
    )
    second_checkbox.pack(anchor="w", padx=10, pady=5)
    study_vars.append(second_var)

    third_var = tk.IntVar()
    third_checkbox = ttk.Checkbutton(
        study_frame,
        text=studies[2],
        variable=third_var,
        state="disabled"  # Start disabled
    )
    third_checkbox.pack(anchor="w", padx=10, pady=5)
    study_vars.append(third_var)

    # Configure the interaction logic
    def on_param_change():
        """Handle parameter checkbox changes"""
        # Check if checkbox 1 is ticked
        checkbox1_ticked = param1_var.get() == 1

        # Check if checkbox 2 is ticked (only if it exists in SEQ region)
        checkbox2_ticked = False
        if region == "SEQ" and param_vars[1] is not None:
            checkbox2_ticked = param_vars[1].get() == 1

        # Enable/disable Conductor Damage and Protection Relay checkboxes based on checkbox 1
        if checkbox1_ticked:
            second_checkbox.config(state="normal")
            third_checkbox.config(state="normal")
        else:
            # Disable and untick if checkbox 1 is not ticked
            second_var.set(0)
            third_var.set(0)
            second_checkbox.config(state="disabled")
            third_checkbox.config(state="disabled")

        # If checkbox 2 is ticked (in SEQ region), also untick and disable
        if checkbox2_ticked:
            second_var.set(0)
            third_var.set(0)
            second_checkbox.config(state="disabled")
            third_checkbox.config(state="disabled")

    # Bind the parameter checkboxes to the change handler
    param1_checkbox.config(command=on_param_change)
    if region == "SEQ" and param_vars[1] is not None:
        param2_checkbox.config(command=on_param_change)

    # Information section
    info_frame = ttk.LabelFrame(parent, text="Information")
    info_frame.pack(fill="x", pady=(0, 15))

    # Get current user ID for the path
    user_id = os.environ.get("USERNAME", "user")

    info_text = f"""If all protection devices are not correctly configured, only a fault level study may be undertaken.

Fault Study results and Conductor Damage Assessment results are stored at the following location:

C:\\LocalData\\{user_id}\\

For script documentation, refer to Job Aid XXXX.
    """

    info_label = ttk.Label(info_frame, text=info_text, justify="left")
    info_label.pack(anchor="w", padx=10, pady=10)

    return param_vars, study_vars


def get_selected_studies(studies, study_vars, param_vars):
    """
    Get the list of selected studies based on checkbox states.

    Args:
        studies: List of available studies
        study_vars: Variables tracking study checkbox states
        param_vars: Variables tracking parameter checkbox states

    Returns:
        list: Selected study names with parameter flags
    """
    selections = []

    # Add parameter flags if checked
    if param_vars[0].get() == 1:
        selections.append("All Relays Configured")

    # Check if checkbox 2 exists (only in SEQ region) and is ticked
    if param_vars[1] is not None and param_vars[1].get() == 1:
        selections.append("Legacy Script")

    # First study is always included (it's disabled but checked)
    selections.append(studies[0])

    # Add other selected studies only if checkbox 1 is ticked
    # and checkbox 2 is either not present or not ticked
    checkbox1_ticked = param_vars[0].get() == 1
    checkbox2_ticked = param_vars[1] is not None and param_vars[1].get() == 1

    if checkbox1_ticked and not checkbox2_ticked:
        for i in range(1, len(studies)):
            if study_vars[i].get() == 1:
                selections.append(studies[i])

    return selections


def exit_script(root, app):
    """
    Exit the script and destroy the UI.

    Args:
        root: The Tkinter root window
        app: The application instance
    """
    app.PrintPlain("User terminated script.")
    root.destroy()
    sys.exit(0)