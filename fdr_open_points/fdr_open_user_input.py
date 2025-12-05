import powerfactory as pf
import sys
import tkinter as tk
from tkinter import ttk
from typing import List, Dict
sys.path.append(r"\\Ecasd01\WksMgmt\PowerFactory\ScriptsDEV\PowerFactoryTyping")
import powerfactorytyping as pft


def mesh_feeder_check(app: pft.Application):
    """Checks if any of the substation feeders are meshes and excludes them from the feeder select list. If no radial
    feeders are found, a message is displayed to the user """

    app.PrintPlain("Checking for mesh feeders...")
    grids_0 = app.GetCalcRelevantObjects('*.ElmXnet')
    grids = [grid for grid in grids_0 if grid.GetAttribute('outserv') == 0]
    all_feeders = app.GetCalcRelevantObjects('*.ElmFeeder')

    radial_dic = {}
    # For each feeder, do a topological search up- and downstream. If the external grid is found in both searches
    # this is a meshed feeder and will be excluded.
    for feeder in all_feeders:
        cubicle = feeder.obj_id
        down_devices = cubicle.GetAll(1, 0)
        up_devices = cubicle.GetAll(0, 0)
        if any(item in grids for item in down_devices) and any(item in grids for item in up_devices):
            continue
        elif any(item in grids for item in down_devices) or any(item in grids for item in up_devices):
            radial_dic[feeder] = feeder.GetAttribute('loc_name')

    if len(radial_dic) == 0:
        root = tk.Tk()
        root.geometry("+200+200")

        root.title("11kV fault study")
        ttk.Label(root, text="No radial feeders were found at the substation").grid(padx=5, pady=5)
        ttk.Label(root, text="To run this script, please radialise one or more feeders").grid(padx=5, pady=5)

        ttk.Button(root, text='Exit', command=lambda: exit_script(root, app)).grid(sticky="s", padx=5, pady=5)

        # Run the interface
        root.mainloop()

    return radial_dic

def get_feeders(app: pft.Application, radial_dic: Dict):
    """From the user, gather a feeder list for study. Also prompt the user to set external grid parameters"""

    radial_list = list(radial_dic.values())
    radial_list.sort()

    # Setup the root window
    root = tk.Tk()

    def _window_dim(radial_list: List[str]):
        horiz_offset = 400
        feeder_col = 350
        window_width = feeder_col
        if window_width > 1300:
            horiz_offset = 200
            window_width = 1500
        num_rows = len(radial_list)
        row_height = 32
        row_padding = 100
        window_height = max((num_rows * row_height + row_padding), 480)
        if window_height > 900:
            window_height = 900
        return window_width, window_height, horiz_offset

    window_width, window_height, horiz_offset = _window_dim(radial_list)
    root.geometry(f"{window_width}x{window_height}+{horiz_offset}+100")
    root.title("11kV Fault Study")

    canvas = tk.Canvas(root, borderwidth=0)
    frame = tk.Frame(canvas)
    vsb = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vsb.set)

    vsb.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    canvas.create_window((4, 4), window=frame, anchor="nw")

    frame.bind("<Configure>", lambda event, canvas=canvas: onFrameConfigure(canvas))

    list_length, var = populate_feeders(app, root, frame, radial_list)

    # Run the interface
    root.mainloop()

    # Collect the feeder list and give prompt if no feeders selected
    feeder_list = []
    for i in range(list_length):
        if var[i].get() == 1:
            feeder_list.append(radial_list[i])
    if not feeder_list:
        feeder_list = window_error(app, radial_dic, 3)
        return feeder_list

    return feeder_list


def onFrameConfigure(canvas):
    """Reset the scroll region to encompass the inner frame"""
    canvas.configure(scrollregion=canvas.bbox("all"))


def populate_feeders(app, root, frame, radial_list):

    ttk.Label(frame, text="Select all feeders to study:", font='Helvetica 14 bold').grid(columnspan=3, padx=5, pady=5)
    ttk.Label(frame, text="(mesh feeders excluded)", font='Helvetica 12 bold').grid(columnspan=3, padx=5, pady=5)

    grids = app.GetCalcRelevantObjects('*.ElmXnet')

    # Get existing external grid values from PowerFactory
    grid_data = {}

    # Create the feeder list interface
    list_length = len(radial_list)
    var = []
    for i in range(list_length):
        var.append(tk.IntVar())
        ttk.Checkbutton(frame, text=radial_list[i], variable=var[i]).grid(column=0, sticky="w", padx=30, pady=5)

    # Calculate bottom row for placement of Okay and Exit buttons
    if list_length + 2 > 12:
        row_index = list_length + 2
    else:
        row_index = 12

    ttk.Button(frame, text='Okay', command=lambda: root.destroy()).grid(row=row_index, column=0, sticky="w", padx=5,
                                                                        pady=5)
    ttk.Button(frame, text='Exit', command=lambda: exit_script(root, app)).grid(row=row_index, column=1, sticky="w",
                                                                           padx=5, pady=5)
    return list_length, var


def window_error(app: pft.Application, radial_dic: Dict, a: int):
    """Check user inputs are formatted correctly"""

    if a == 1:
        app.PrintPlain("Please enter fault level values in kA, not A")
    elif a == 2:
        app.PrintPlain("Please enter numerical values only")
    else:
        app.PrintPlain("Please select at least one feeder to continue")

    feeder_list = get_feeders(app, radial_dic)
    return feeder_list


def exit_script(root, app):
    """Exits script"""

    app.PrintPlain("User terminated script.")
    root.destroy()
    sys.exit(0)