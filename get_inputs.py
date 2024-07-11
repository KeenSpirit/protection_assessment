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

    def main(self, region: str) -> Tuple[Dict, Dict, Dict]:
        radial_list = self.mesh_feeder_check()
        feeder_list, external_grid = self.feeders_external_grid(radial_list)
        feeders_devices = self.get_feeders_devices(feeder_list)
        user_selection = self.run_window(feeder_list, feeders_devices, region)
        return feeders_devices, user_selection, external_grid

    def mesh_feeder_check(self) -> List[str]:
        self.app.PrintPlain("Checking for mesh feeders...")
        grids = self.app.GetCalcRelevantObjects('*.ElmXnet')
        all_feeders = self.app.GetCalcRelevantObjects('*.ElmFeeder')

        radial_list = [
            feeder.loc_name for feeder in all_feeders
            if not (set(feeder.obj_id.GetAll(1, 0)) & set(grids) and
                    set(feeder.obj_id.GetAll(0, 0)) & set(grids) and
                    not feeder.IsOutOfService())
        ]

        if not radial_list:
            self.show_no_radial_feeders_message()

        return sorted(radial_list)

    def show_no_radial_feeders_message(self):
        root = tk.Tk()
        root.geometry("+200+200")
        root.title("Distribution fault study")
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
            horiz_offset = 400
            grids = [grid for grid in self.app.GetCalcRelevantObjects('*.ElmXnet') if grid.outserv == 0]
            grid_cols = len(grids)
            column_width = 360
            feeder_col = 285
            window_width = grid_cols * column_width + feeder_col
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
        root.title("Distribution Fault Study")

        canvas, frame = self.setup_scrollable_frame(root)

        list_length, var, load, grids = self.populate_feeders(root, frame, radial_list)

        root.mainloop()

        feeder_list = [radial_list[i] for i in range(list_length) if var[i].get() == 1]

        if not feeder_list:
            return self.window_error(radial_list, 3)

        new_grid_data = self.collect_grid_data(load, radial_list)

        if not self.validate_grid_data(new_grid_data):
            return self.window_error(radial_list, 1)

        self.update_grid_data(grids, new_grid_data)

        return feeder_list, new_grid_data

    def setup_scrollable_frame(self, root):
        canvas = tk.Canvas(root, borderwidth=0)
        frame = tk.Frame(canvas)
        vsb = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.create_window((4, 4), window=frame, anchor="nw")
        frame.bind("<Configure>", lambda event, canvas=canvas: self.onFrameConfigure(canvas))
        return canvas, frame

    def populate_feeders(self, root, frame, radial_list):
        ttk.Label(frame, text="Select all feeders to study:", font='Helvetica 14 bold').grid(columnspan=3, padx=5,
                                                                                             pady=5)
        ttk.Label(frame, text="(mesh feeders excluded)", font='Helvetica 12 bold').grid(columnspan=3, padx=5, pady=5)

        grids = [grid for grid in self.app.GetCalcRelevantObjects('*.ElmXnet') if grid.outserv == 0]

        grid_data = self.get_grid_data(grids)

        var = self.create_feeder_checkboxes(frame, radial_list)

        frame.columnconfigure(4, minsize=100)

        load = self.create_external_grid_interface(frame, grids, grid_data)

        row_index = max(len(radial_list) + 2, 12)

        ttk.Button(frame, text='Okay', command=root.destroy).grid(row=row_index, column=0, sticky="w", padx=5, pady=5)
        ttk.Button(frame, text='Exit', command=lambda: self.exit_script(root)).grid(row=row_index, column=1, sticky="w",
                                                                                    padx=5, pady=5)

        return len(radial_list), var, load, grids

    def get_grid_data(self, grids):
        grid_data = {}
        attributes = ['ikss', 'rntxn', 'z2tz1', 'x0tx1', 'r0tx0', 'ikssmin', 'rntxnmin', 'z2tz1min', 'x0tx1min',
                      'r0tx0min']
        for grid in grids:
            grid_data[grid] = [round(grid.GetAttribute(attr), 8 if 'min' in attr else 3) for attr in attributes]
        return grid_data

    def create_feeder_checkboxes(self, frame, radial_list):
        var = []
        for i, feeder in enumerate(radial_list):
            var.append(tk.IntVar())
            ttk.Checkbutton(frame, text=feeder, variable=var[-1]).grid(column=0, sticky="w", padx=30, pady=5)
        return var

    def create_external_grid_interface(self, frame, grids, grid_data):
        ttk.Label(frame, text="Enter external grid data:", font='Helvetica 14 bold').grid(column=5, columnspan=3,
                                                                                          sticky="w", row=0, padx=5,
                                                                                          pady=5)
        load = {}
        for i, grid in enumerate(grids):
            load[grid] = self.create_grid_entries(frame, grid, grid_data[grid], 5 + i * 3)
        return load

    def create_grid_entries(self, frame, grid, data, column):
        grid_entries = []
        labels = ["P-P-P fault max.", "P-P-P fault min.", "R/X max.", "Z2/Z1 max.", "X0/X1 max.", "R0/X0 max.",
                  "R/X min.", "Z2/Z1 min.", "X0/X1 min.", "R0/X0 min."]

        tk.Label(frame, text=f"{grid.loc_name}:", font='Helvetica 12 bold').grid(row=1, column=column, padx=5, pady=5)

        for i, (label, value) in enumerate(zip(labels, data)):
            tk.Label(frame, text=label).grid(row=i + 2, column=column, pady=5)
            var = tk.DoubleVar(value=value)
            tk.Entry(frame, textvariable=var).grid(row=i + 2, column=column + 1, pady=5)
            if i in [0, 1]:
                tk.Label(frame, text='kA').grid(row=i + 2, column=column + 2, ipadx=5, pady=5)
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

    def get_feeders_devices(self, radial_list: List[str]) -> Dict[str, List[Any]]:
        """Get active relays and fuses. """
        # Relays belong in the network model project folder.
        net_mod = self.app.GetProjectFolder("netmod")
        # Filter for for relays under network model recursively.
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

        device_object_dict = {feeder: [] for feeder in radial_list}
        for device in devices:
            term = device.cbranch
            feeder = [feeder for feeder in radial_list if term in self.app.GetCalcRelevantObjects(feeder + ".ElmFeeder")[0].GetAll()]
            if feeder:
                device_object_dict[feeder[0]].append(device)
        return device_object_dict

    def populate(self, frame, feeder_list, feeders_switches, region):
        if region == 'Regional Models':
            fdr_sw_locname = {feeder: [switch.loc_name for switch in switches]
                              for feeder, switches in feeders_switches.items()}
        else:  # region == 'SEQ Models'
            fdr_sw_locname = {feeder: [] for feeder in feeders_switches}
            for feeder, switches in feeders_switches.items():
                for switch in switches:
                    switch_term = switch.fold_id.cterm.loc_name
                    fdr_sw_locname[feeder].append(switch_term[:-5] if switch_term.endswith("_Term") else switch_term)

        ttk.Label(frame, text="Identify ALL ACRs and line fuses belonging to the given feeder:",
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

        ttk.Button(frame, text='Okay', command=lambda: frame.master.master.destroy()).grid(
            row=max_list_length + 5, column=0, sticky='W', padx=5, pady=5
        )
        ttk.Button(frame, text='Exit', command=lambda: self.exit_script(frame.master.master)).grid(
            row=max_list_length + 5, column=1, sticky='W', pady=5
        )

        return var, fdr_sw_locname

    def run_window(self, feeder_list: List[str], feeders_devices: Dict[str, List[Any]], region: str) -> Dict[
        str, List[Any]]:

        def _window_dim(feeder_list, feeders_switches, region):
            horiz_offset = 500
            num_columns = len(feeder_list)
            if region == 'Regional Models':
                column_width = 230
            else:
                column_width = 120
            col_padding = 0
            window_width = max((num_columns * column_width + col_padding), 550)
            if window_width > 1300:
                horiz_offset = 200
                window_width = 1500
            num_rows = max(len(list) for list in feeders_switches.values())
            row_height = 32
            row_padding = 100
            window_height = num_rows * row_height + row_padding
            if window_height > 900:
                window_height = 900
            return window_width, window_height, horiz_offset

        root = tk.Tk()
        root.title("Distribution Fault Study")
        window_width, window_height, horiz_offset = _window_dim(feeder_list, feeders_devices, region)
        root.geometry(f"{window_width}x{window_height}+{horiz_offset}+100")

        canvas, frame = self.setup_scrollable_frame(root)

        var, fdr_dev_locname = self.populate(frame, feeder_list, feeders_devices, region)

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
    def onFrameConfigure(canvas):
        canvas.configure(scrollregion=canvas.bbox("all"))

def get_input(app, region):
    study = FaultLevelStudy(app)
    return study.main(region)