import powerfactory as pf
import tkinter as tk
from tkinter import ttk
import sys


def get_study_selections(app):
    """From the user, gather a feeder list for study. Also prompt the user to set external grid parameters"""

    # Setup the root window
    root = tk.Tk()

    def _window_dim():
        horiz_offset = 600
        selection_col = 285
        col_padding = 330
        window_width = selection_col + col_padding
        num_rows = 8
        row_height = 32
        row_padding = 120
        window_height = min((num_rows * row_height + row_padding), 480)
        if window_height > 900:
            window_height = 900
        return window_width, window_height, horiz_offset

    window_width, window_height, horiz_offset = _window_dim()
    root.geometry(f"{window_width}x{window_height}+{horiz_offset}+300")
    root.title("Distribution Protection Assessment")

    canvas = tk.Canvas(root, borderwidth=0)
    frame = tk.Frame(canvas)

    canvas.pack(side="left", fill="both", expand=True)
    canvas.create_window((4, 4), window=frame, anchor="nw")

    frame.bind("<Configure>", lambda event, canvas=canvas: onFrameConfigure(canvas))

    choices = [
        "Conductor Damage Assessment",
        "Protection Relay Coordination Plot",
        "Protection Audit"
    ]
    var = populate_feeders(root, frame, choices)

    # Run the interface
    root.mainloop()
    list_length = 3
    # Collect the feeder list and give prompt if no feeders selected
    selections = []
    for i in range(list_length):
        if var[i].get() == 1:
            selections.append(choices[i])

    return selections


def populate_feeders(root, frame, choices):

    ttk.Label(frame, text="Select all studies to undertake:", font='Helvetica 14 bold').grid(columnspan=3, padx=5, pady=5)
    ttk.Label(frame, text="", font='Helvetica 12 bold').grid(columnspan=3, padx=5, pady=5)

    # Create the list interface
    first_var = tk.IntVar(value=1)
    ttk.Checkbutton(frame, text='Distribution Fault Level Study', variable=first_var, state="disabled").grid(column=0, sticky="w", padx=30, pady=5)
    list_length = 3
    var = []
    for i in range(list_length):
        var.append(tk.IntVar())
        ttk.Checkbutton(frame, text=choices[i], variable=var[i]).grid(column=0, sticky="w", padx=30, pady=5)


    ttk.Label(frame, text="Note:").grid(sticky="w", padx=5, pady=5)
    ttk.Label(frame, text="Fault Study results and Conductor Damage Assessment results are stored at the following location:").grid(columnspan=3, sticky="w", padx=5, pady=5)
    ttk.Label(frame, text="\\C\\LocalData\\{user_id}\\").grid(sticky="w", padx=5, pady=5)
    ttk.Label(frame,
              text="For the documentation for this script, Refer to Job Aid XXXX.").grid(
        columnspan=3, sticky="w", padx=5, pady=5)

    frame.columnconfigure(4, minsize=100)

    # Calculate bottom row for placement of Okay and Exit buttons
    row_index = 10

    ttk.Button(frame, text='Okay', command=lambda: root.destroy()).grid(row=row_index, column=0, sticky="w", padx=5,
                                                                        pady=5)
    ttk.Button(frame, text='Exit', command=lambda: exit_script(root)).grid(row=row_index, column=1, sticky="w",
                                                                           padx=5, pady=5)
    return var


def onFrameConfigure(canvas):
    """Reset the scroll region to encompass the inner frame"""
    canvas.configure(scrollregion=canvas.bbox("all"))


def exit_script(root):

    """Exits script"""
    app = pf.GetApplication()
    app.PrintPlain("User terminated script.")
    root.destroy()
    sys.exit(0)