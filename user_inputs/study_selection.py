import tkinter as tk
from tkinter import ttk
import sys
import os


def get_study_selections(app):
    """
    From the user, gather a study selection list. Also prompt the user to set external grid parameters.

    Args:
        app: The application instance

    Returns:
        list: Selected study options
    """
    # Create and configure the root window
    root = tk.Tk()
    root.title("Distribution Protection Assessment")

    # Configure window dimensions and position
    window_width = 615
    window_height = 450
    screen_width = root.winfo_screenwidth()
    screen_x_position = (screen_width - window_width) // 2
    root.geometry(f"{window_width}x{window_height}+{screen_x_position}+300")
    root.resizable(False, True)

    # Apply a modern theme if available
    try:
        root.tk.call("source", "azure.tcl")
        root.tk.call("set_theme", "light")
    except tk.TclError:
        pass  # Continue with default theme if custom theme isn't available

    # Create main frame with scrolling capability
    main_frame = ttk.Frame(root)
    main_frame.pack(fill="both", expand=True, padx=10, pady=10)

    # Create canvas for scrolling
    canvas = tk.Canvas(main_frame)
    scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    # Pack scrolling components
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Study choices available for selection
    studies = [
        "Distribution Fault Level Study",
        "Conductor Damage Assessment",
        "Protection Relay Coordination Plot",
        "Protection Audit"
    ]

    # Study variables and GUI elements
    study_vars = create_ui_elements(scrollable_frame, studies)

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
    selections = get_selected_studies(studies, study_vars)
    return selections


def create_ui_elements(parent, studies):
    """
    Create the UI elements for the study selection interface.

    Args:
        parent: Parent frame to place elements in
        studies: List of available studies

    Returns:
        list: Variables tracking checkbox states
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

    # Subheader
    subheader = ttk.Label(
        parent,
        text="Select all studies to undertake:",
        font=("Helvetica", 12)
    )
    subheader.pack(anchor="w", pady=(0, 10))

    # Checkboxes for study selection
    checkbox_frame = ttk.Frame(parent)
    checkbox_frame.pack(fill="x")

    study_vars = []

    # First checkbox is always checked and disabled
    first_var = tk.IntVar(value=1)
    first_checkbox = ttk.Checkbutton(
        checkbox_frame,
        text=studies[0],
        variable=first_var,
        state="disabled"
    )
    first_checkbox.pack(anchor="w", padx=5, pady=5)
    study_vars.append(first_var)

    # Create remaining checkboxes
    for i in range(1, len(studies)):
        var = tk.IntVar()
        checkbox = ttk.Checkbutton(
            checkbox_frame,
            text=studies[i],
            variable=var
        )
        checkbox.pack(anchor="w", padx=5, pady=5)
        study_vars.append(var)

    # Information section
    info_frame = ttk.LabelFrame(parent, text="Information")
    info_frame.pack(fill="x", pady=15)

    # Get current user ID for the path
    user_id = os.environ.get("USERNAME", "user")

    info_text = f"""
Fault Study results and Conductor Damage Assessment results 
are stored at the following location:

C:\\LocalData\\{user_id}\\

For documentation, refer to Job Aid XXXX.
    """

    info_label = ttk.Label(info_frame, text=info_text, justify="left")
    info_label.pack(anchor="w", padx=10, pady=10)

    return study_vars


def get_selected_studies(studies, study_vars):
    """
    Get the list of selected studies based on checkbox states.

    Args:
        studies: List of available studies
        study_vars: Variables tracking checkbox states

    Returns:
        list: Selected study names
    """
    selections = []

    # First study is always included (it's disabled but checked)
    selections.append(studies[0])

    # Add other selected studies
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
