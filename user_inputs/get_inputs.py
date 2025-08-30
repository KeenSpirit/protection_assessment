from tkinter import *  # noqa [F403]
from tkinter import ttk
import sys
from typing import Dict, List, Tuple, Any


def loc_name(device):

    if device.HasAttribute("r:cpSubstat:e:loc_name"):
        sub_code = device.GetAttribute("r:cpSubstat:e:loc_name")
        if (
                sub_code == device.loc_name[: len(sub_code)]
                or device.loc_name[:2] == "RC"
        ):
            key = device.loc_name
        else:
            key = f"{sub_code}_{device.loc_name}"
    else:
        key = device.loc_name
    return key


import tkinter as tk
from tkinter import ttk
import sys
from typing import Dict, List, Tuple, Any


def get_grid_data(app) -> list[dict[str:float]]:
    """
    PowerFactory model external grid data read to put in a pd.DataFrame
    :param app:
    :return:
    """

    all_grids = app.GetCalcRelevantObjects('*.ElmXnet')
    grids = [grid for grid in all_grids if grid.outserv == 0]

    grid_data = {}
    for grid in grids:
        grid_data['Parameter'] = ['3-P fault level (A):', 'R/X:', 'Z2/Z1:', 'X0/X1:', 'R0/X0:']
        grid_data[f'{grid.loc_name} Maximum'] = [
            round(grid.GetAttribute('ikss'), 3),
            round(grid.GetAttribute('rntxn'), 8),
            round(grid.GetAttribute('z2tz1'), 8),
            round(grid.GetAttribute('x0tx1'), 8),
            round(grid.GetAttribute('r0tx0'), 8)
        ]
        grid_data[f'{grid.loc_name} Minimum'] = [
            round(grid.GetAttribute('ikssmin'), 3),
            round(grid.GetAttribute('rntxnmin'), 8),
            round(grid.GetAttribute('z2tz1min'), 8),
            round(grid.GetAttribute('x0tx1min'), 8),
            round(grid.GetAttribute('r0tx0min'), 8)
        ]
    return grid_data


def reset_min_source_imp(app, new_grid_data, sys_norm_min=False):
    """"""

    grids = app.GetCalcRelevantObjects('*.ElmXnet')

    for grid in grids:
        name = grid.loc_name
        if sys_norm_min:
            grid.ikssmin = new_grid_data[name][0]
            grid.rntxnmin = new_grid_data[name][1]
            grid.z2tz1min = new_grid_data[name][2]
            grid.x0tx1min = new_grid_data[name][3]
            grid.r0tx0min = new_grid_data[name][4]
        else:
            grid.ikssmin = new_grid_data[name][5]
            grid.rntxnmin = new_grid_data[name][6]
            grid.z2tz1min = new_grid_data[name][7]
            grid.x0tx1min = new_grid_data[name][8]
            grid.r0tx0min = new_grid_data[name][9]
    return


def determine_fuse_type(fuse):
    """This function will observe the fuse location and determine if it is
    a Distribution transformer fuse, SWER isolating fuse or a line fuse"""
    # First check is that if the fuse exists in a terminal that is in the
    # System Overiew then it will be a line fuse.
    fuse_active = fuse.HasAttribute("r:fold_id:r:obj_id:e:loc_name")
    if not fuse_active:
        return True
    fuse_grid = fuse.cpGrid
    if (
            fuse.GetAttribute("r:fold_id:r:cterm:r:fold_id:e:loc_name")
            == fuse_grid.loc_name
    ):
        # This would indicate it is in a line cubical
        return True
    if fuse.loc_name not in fuse.GetAttribute("r:fold_id:r:obj_id:e:loc_name"):
        # This indicates that the fuse is not in a switch object
        return True
    secondary_sub = fuse.fold_id.cterm.fold_id
    contents = secondary_sub.GetContents()
    for content in contents:
        if content.GetClassName() == "ElmTr2":
            return False
    else:
        return True


class FaultLevelStudy:
    def __init__(self, app):
        self.app = app

    def main(self, region: str, study_selections: int) -> Tuple[Dict, Dict, Dict, Dict]:
        radial_list = self.mesh_feeder_check()
        feeder_list, external_grid = self.feeders_external_grid(radial_list)
        feeders_devices, bu_devices = self.get_feeders_devices(feeder_list)
        user_selection = self.run_window(feeder_list, feeders_devices, region)
        return feeders_devices, bu_devices, user_selection, external_grid

    def center_window(self, root, width, height):
        """Center the window on the user's screen"""
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        root.geometry(f"{width}x{height}+{x}+{y}")

    def mesh_feeder_check(self) -> List[str]:
        self.app.PrintPlain("Checking for radial feeders...")
        grids = self.app.GetCalcRelevantObjects('*.ElmXnet')
        all_feeders = self.app.GetCalcRelevantObjects('*.ElmFeeder')

        radial_list = [
            feeder.loc_name for feeder in all_feeders
            if not (set(feeder.obj_id.GetAll(1, 0)) & set(grids) and
                    set(feeder.obj_id.GetAll(0, 0)) & set(grids) and
                    not feeder.IsOutOfService())
        ]

        if radial_list:
            self.app.PrintPlain("Radial feeders detected.")
            self.app.PrintPlain("Please enter the requested inputs.")
        else:
            self.show_no_radial_feeders_message()

        return sorted(radial_list)

    def show_no_radial_feeders_message(self):
        root = tk.Tk()
        root.title("Distribution fault study")

        # Calculate window size and center it
        window_width = 400
        window_height = 150
        self.center_window(root, window_width, window_height)

        ttk.Label(root, text="No radial feeders were found at the substation").grid(padx=5, pady=5)
        ttk.Label(root, text="To run this script, please radialise one or more feeders").grid(padx=5, pady=5)
        ttk.Button(root, text='Exit', command=lambda: self.exit_script(root)).grid(sticky="s", padx=5, pady=5)
        root.mainloop()

    def exit_script(self, root):
        self.app.PrintPlain("User terminated script.")
        root.destroy()
        sys.exit(0)

    def feeders_external_grid(self, radial_list: List[str]) -> Tuple[List[str], Dict[str, List[float]]]:
        root = tk.Tk()

        def _window_dim(radial_list):
            grids = [grid for grid in self.app.GetCalcRelevantObjects('*.ElmXnet') if grid.outserv == 0]
            grid_cols = len(grids)
            column_width = 360
            feeder_col = 285
            window_width = max((grid_cols * column_width + feeder_col), 700)
            if window_width > 1300:
                window_width = 1500
            num_rows = len(radial_list)
            row_height = 32
            row_padding = 100
            window_height = max((num_rows * row_height + row_padding), 495)
            if window_height > 900:
                window_height = 900
            return window_width, window_height

        window_width, window_height = _window_dim(radial_list)
        self.center_window(root, window_width, window_height)
        root.title("Distribution Fault Study")

        # Create main container frame
        main_frame = tk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True)

        canvas, frame = self.setup_scrollable_frame(main_frame)

        # Create button frame at the bottom, outside the scroll area
        button_frame = tk.Frame(root)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=5)

        list_length, var, load, grids = self.populate_feeders(root, frame, radial_list, button_frame)

        root.mainloop()

        feeder_list = [radial_list[i] for i in range(list_length) if var[i].get() == 1]

        if not feeder_list:
            return self.window_error(radial_list, 3)

        new_grid_data = self.collect_grid_data(load, radial_list)

        if not self.validate_grid_data(new_grid_data):
            return self.window_error(radial_list, 1)

        self.update_grid_data(grids, new_grid_data)

        return feeder_list, new_grid_data

    def setup_scrollable_frame(self, parent):
        canvas = tk.Canvas(parent, borderwidth=0)
        frame = tk.Frame(canvas)
        vsb = tk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        hsb = tk.Scrollbar(parent, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Initially hide scrollbars
        canvas.pack(fill="both", expand=True)
        canvas.create_window((4, 4), window=frame, anchor="nw")

        # Bind events to show/hide scrollbars as needed
        frame.bind("<Configure>", lambda event: self.onFrameConfigure(canvas, vsb, hsb))
        canvas.bind("<Configure>", lambda event: self.onCanvasConfigure(canvas, vsb, hsb))

        return canvas, frame

    def populate_feeders(self, root, frame, radial_list, button_frame):
        ttk.Label(frame, text="Select all feeders to study:", font='Helvetica 14 bold').grid(columnspan=3, padx=5,
                                                                                             pady=5)
        ttk.Label(frame, text="(mesh feeders excluded)", font='Helvetica 10 bold').grid(columnspan=3, padx=5, pady=5)

        grids = [grid for grid in self.app.GetCalcRelevantObjects('*.ElmXnet') if grid.outserv == 0]

        grid_data = self.get_grid_data(grids)

        # Create a separate frame for feeder checkboxes to avoid row height interference
        feeder_frame = tk.Frame(frame)
        feeder_frame.grid(row=2, column=0, columnspan=4, sticky="nw", padx=5, pady=5)

        var = self.create_feeder_checkboxes(feeder_frame, radial_list)

        frame.columnconfigure(4, minsize=100)

        load = self.create_external_grid_interface(frame, grids, grid_data)

        # Place buttons in the button_frame instead of the scrollable frame
        ttk.Button(button_frame, text='Okay', command=root.destroy).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text='Exit', command=lambda: self.exit_script(root)).pack(side=tk.LEFT, padx=5)

        return len(radial_list), var, load, grids

    def get_grid_data(self, grids):
        grid_data = {}
        attributes = ['ikss', 'rntxn', 'z2tz1', 'x0tx1', 'r0tx0', 'ikssmin', 'rntxnmin', 'z2tz1min', 'x0tx1min',
                      'r0tx0min']
        for grid in grids:
            grid_data[grid] = [grid.GetAttribute(attr) for attr in attributes]
        return grid_data

    def create_feeder_checkboxes(self, feeder_frame, radial_list):
        var = []
        for i, feeder in enumerate(radial_list):
            var.append(tk.IntVar())
            ttk.Checkbutton(feeder_frame, text=feeder, variable=var[-1]).grid(row=i, column=0, sticky="w", padx=25,
                                                                              pady=5)
        return var

    def create_external_grid_interface(self, frame, grids, grid_data):
        ttk.Label(frame, text="Enter external grid data:", font='Helvetica 14 bold').grid(column=5, columnspan=3,
                                                                                          sticky="w", row=0, padx=5,
                                                                                          pady=5)

        # Create a dedicated frame for grid entries to avoid height conflicts
        grid_entries_frame = tk.Frame(frame)
        grid_entries_frame.grid(row=2, column=5, columnspan=10, sticky="nw", padx=5, pady=5)

        load = {}
        for i, grid in enumerate(grids):
            load[grid] = self.create_grid_entries(grid_entries_frame, grid, grid_data[grid], i * 3)
        return load

    def create_grid_entries(self, grid_frame, grid, data, column):
        grid_entries = []
        labels = ["P-P-P fault max.", "R/X max.", "Z2/Z1 max.", "X0/X1 max.", "R0/X0 max.", "P-P-P fault min.",
                  "R/X min.", "Z2/Z1 min.", "X0/X1 min.", "R0/X0 min."]

        # Create bordered frames for max and min parameters
        max_frame = tk.LabelFrame(grid_frame, text=f"{grid.loc_name} Maximum Values", relief="solid", bd=1, padx=5,
                                  pady=5)
        max_frame.grid(row=0, column=column, columnspan=3, padx=5, pady=5, sticky="ew")

        min_frame = tk.LabelFrame(grid_frame, text=f"{grid.loc_name} Minimum Values", relief="solid", bd=1, padx=5,
                                  pady=5)
        min_frame.grid(row=1, column=column, columnspan=3, padx=5, pady=5, sticky="ew")

        # Create entries for maximum values (first 5 parameters)
        for i in range(5):
            label, value = labels[i], data[i]
            tk.Label(max_frame, text=label).grid(row=i, column=0, pady=2, sticky="w")
            var = tk.DoubleVar(value=value)
            tk.Entry(max_frame, textvariable=var).grid(row=i, column=1, pady=2, padx=(5, 0))
            if i in [0]:  # Only first entry (P-P-P fault max.) gets kA unit
                tk.Label(max_frame, text='kA').grid(row=i, column=2, ipadx=5, pady=2)
            grid_entries.append(var)

        # Create entries for minimum values (last 5 parameters)
        for i in range(5, 10):
            label, value = labels[i], data[i]
            tk.Label(min_frame, text=label).grid(row=i - 5, column=0, pady=2, sticky="w")
            var = tk.DoubleVar(value=value)
            tk.Entry(min_frame, textvariable=var).grid(row=i - 5, column=1, pady=2, padx=(5, 0))
            if i == 5:  # Only P-P-P fault min. gets kA unit
                tk.Label(min_frame, text='kA').grid(row=i - 5, column=2, ipadx=5, pady=2)
            grid_entries.append(var)

        return grid_entries

    def collect_grid_data(self, load, radial_list):
        new_grid_data = {}
        for grid in load:
            try:
                new_grid_data[grid] = [item.get() for item in load[grid]]
            except Exception:
                return self.window_error(radial_list, 2)
        return new_grid_data

    def validate_grid_data(self, new_grid_data):
        return all(new_grid_data[grid][0] <= 100 and new_grid_data[grid][5] <= 100 for grid in new_grid_data)

    def update_grid_data(self, grids, new_grid_data):
        for grid in grids:
            for attr, value in zip(
                    ['ikss', 'rntxn', 'z2tz1', 'x0tx1', 'r0tx0', 'ikssmin', 'rntxnmin', 'z2tz1min', 'x0tx1min',
                     'r0tx0min'], new_grid_data[grid]):
                setattr(grid, attr, value)
            grid.cmax = 1.1
            grid.cmin = 1
            name = grid.loc_name
            new_grid_data[name] = new_grid_data[grid]
            del new_grid_data[grid]

    def window_error(self, radial_list, error_code):
        error_messages = {
            1: "Please enter fault level values in kA, not A",
            2: "Please enter numerical values only",
            3: "Please select at least one feeder to continue"
        }
        self.app.PrintPlain(error_messages[error_code])
        return self.feeders_external_grid(radial_list)

    def get_feeders_devices(self, radial_list: List[str]) -> Tuple[Dict[str,list],Dict[Any,list]]:
        """Get active relays and fuses.
        Map them to corresponding external grid or feeder using a dictionary"""
        # Relays belong in the network model project folder.
        net_mod = self.app.GetProjectFolder("netmod")
        # Filter for relays under network model recursively.
        all_relays = net_mod.GetContents("*.ElmRelay", True)
        relays = [
            relay
            for relay in all_relays
            if relay.cpGrid
            if relay.cpGrid.IsCalcRelevant()
            if relay.GetParent().GetClassName() == "StaCubic"
            if relay.fold_id.cterm.IsEnergized()
            if not relay.IsOutOfService()
        ]
        # Create a list of active fuses
        all_fuses = net_mod.GetContents("*.RelFuse", True)
        fuses = [
            fuse
            for fuse in all_fuses
            if fuse.cpGrid
            if fuse.cpGrid.IsCalcRelevant()
            if fuse.fold_id.HasAttribute("cterm")
            if fuse.fold_id.cterm.IsEnergized()
            if not fuse.IsOutOfService()
            if determine_fuse_type(fuse)
        ]
        devices = relays + fuses

        feeder_device_dict = {feeder: [] for feeder in radial_list}
        grid_device_dict = {grid: [] for grid in self.app.GetCalcRelevantObjects('*.ElmXnet') if grid.bus1 is not None}
        for device in devices:
            term = device.cbranch
            feeder = [feeder for feeder in radial_list if
                      term in self.app.GetCalcRelevantObjects(feeder + ".ElmFeeder")[0].GetAll()]
            if feeder:
                feeder_device_dict[feeder[0]].append(device)
                continue
            for grid in grid_device_dict:
                try:
                    grid_term = grid.bus1.cterm
                    grid_term.SetAttribute("iUsage", 0)
                    if grid_term == device.cn_bus:
                        grid_device_dict[grid].append(device)
                        break
                except AttributeError:
                    self.app.PrintPlain(grid)
                    exit(0)
        return feeder_device_dict, grid_device_dict

    def populate(self, frame, feeder_list, feeders_switches, region, button_frame):
        if region == 'Regional Models':
            fdr_sw_locname = {feeder: [switch.loc_name for switch in switches]
                              for feeder, switches in feeders_switches.items()}
        else:  # region == 'SEQ Models'
            fdr_sw_locname = {feeder: [] for feeder in feeders_switches}
            for feeder, switches in feeders_switches.items():
                for switch in switches:
                    switch_term = switch.fold_id.cterm.loc_name
                    fdr_sw_locname[feeder].append(switch_term[:-5] if switch_term.endswith("_Term") else switch_term)

        ttk.Label(frame, text="Select all protection devices to study:",
                  font='Helvetica 12 bold').grid(columnspan=8, padx=5, pady=5)

        for idx, fid in enumerate(feeder_list):
            ttk.Label(frame, text=fid).grid(row=1, column=idx, sticky='W', padx=10, pady=5)

        var = []
        max_list_length = max(len(switches) for switches in fdr_sw_locname.values())

        for feeder, switch_list in fdr_sw_locname.items():
            col = list(fdr_sw_locname).index(feeder)
            for i, switch in enumerate(switch_list):
                var.append(tk.IntVar())
                ttk.Checkbutton(frame, text=switch, variable=var[-1]).grid(
                    row=i + 4, column=col, sticky='W', padx=10, pady=5
                )

        # Define functions for select all and unselect all
        def select_all():
            for checkbox_var in var:
                checkbox_var.set(1)

        def unselect_all():
            for checkbox_var in var:
                checkbox_var.set(0)

        # Place buttons in the button_frame
        ttk.Button(button_frame, text='Select All', command=select_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text='Unselect All', command=unselect_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text='Okay', command=lambda: button_frame.master.destroy()).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text='Exit', command=lambda: self.exit_script(button_frame.master)).pack(side=tk.LEFT,
                                                                                                          padx=5)

        return var, fdr_sw_locname

    def run_window(self, feeder_list: List[str], feeders_devices: Dict[str, List[Any]], region: str) -> Dict[
        str, List[Any]]:

        def _window_dim(feeder_list, feeders_switches, region):
            num_columns = len(feeder_list)
            if region == 'Regional Models':
                column_width = 230
            else:
                column_width = 120
            col_padding = 0
            window_width = max((num_columns * column_width + col_padding), 600)
            if window_width > 1300:
                window_width = 1500
            num_rows = max(len(list) for list in feeders_switches.values())
            row_height = 32
            row_padding = 100
            window_height = max((num_rows * row_height + row_padding), 350)
            if window_height > 900:
                window_height = 900
            return window_width, window_height

        root = tk.Tk()
        root.title("Distribution Fault Study")
        window_width, window_height = _window_dim(feeder_list, feeders_devices, region)
        self.center_window(root, window_width, window_height)

        # Create main container frame
        main_frame = tk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True)

        canvas, frame = self.setup_scrollable_frame(main_frame)

        # Create button frame at the bottom, outside the scroll area
        button_frame = tk.Frame(root)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=5)

        var, fdr_dev_locname = self.populate(frame, feeder_list, feeders_devices, region, button_frame)

        canvas.pack(expand=True, fill=tk.BOTH)

        root.mainloop()

        acr_fuse_list = [
            dev for feeder in fdr_dev_locname
            for i, dev in enumerate(feeders_devices[feeder])
            if var[sum(len(fdr_dev_locname[f]) for f in
                       list(fdr_dev_locname)[:list(fdr_dev_locname).index(feeder)]) + i].get() == 1
        ]

        feeders_relays = {feeder: [switch for switch in switches if switch in acr_fuse_list]
                          for feeder, switches in feeders_devices.items()}

        for feeder, relays in feeders_relays.items():
            feeder_elm = self.app.GetCalcRelevantObjects(feeder + ".ElmFeeder")
            relays.append(feeder_elm[0])

        return feeders_relays

    @staticmethod
    def onFrameConfigure(canvas, vsb, hsb):
        """Update scroll region and show/hide scrollbars as needed"""
        canvas.configure(scrollregion=canvas.bbox("all"))
        FaultLevelStudy.update_scrollbars(canvas, vsb, hsb)

    @staticmethod
    def onCanvasConfigure(canvas, vsb, hsb):
        """Handle canvas resize events"""
        FaultLevelStudy.update_scrollbars(canvas, vsb, hsb)

    @staticmethod
    def update_scrollbars(canvas, vsb, hsb):
        """Show or hide scrollbars based on content size vs canvas size"""
        # Get the scroll region
        scrollregion = canvas.cget("scrollregion")
        if not scrollregion:
            return

        # Parse scroll region (x1, y1, x2, y2)
        x1, y1, x2, y2 = map(float, scrollregion.split())
        content_width = x2 - x1
        content_height = y2 - y1

        # Get canvas dimensions
        canvas_width = canvas.winfo_width()
        canvas_height = canvas.winfo_height()

        # Show/hide vertical scrollbar
        if content_height > canvas_height:
            vsb.pack(side="right", fill="y", before=canvas)
        else:
            vsb.pack_forget()

        # Show/hide horizontal scrollbar
        if content_width > canvas_width:
            hsb.pack(side="bottom", fill="x", before=canvas)
        else:
            hsb.pack_forget()

        # Repack canvas to fill remaining space
        canvas.pack(fill="both", expand=True)


def get_input(app, region, study_selections):
    study = FaultLevelStudy(app)
    return study.main(region, study_selections)
