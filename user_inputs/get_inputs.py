from tkinter import *  # noqa [F403]
from pf_config import pft
from domain.enums import ElementType
from devices import fuses
from relays import elements
import tkinter as tk
from tkinter import ttk
import sys
from typing import Dict, List, Tuple, Any


class FaultLevelStudy:
    def __init__(self, app):
        self.app = app

    def main(self, region: str, study_selections: List) -> Tuple[Dict, Dict, Dict, Dict]:

        radial_list, lines_oos = self.mesh_feeder_check()
        feeder_list, external_grid = self.feeders_external_grid(radial_list, lines_oos)
        if 'Fault Level Study (all relays configured in model)' in study_selections:
            feeders_devices, bu_devices = self.get_feeders_devices(feeder_list)
            self.chk_empty_fdrs(feeders_devices)
            user_selection = self.run_window(feeder_list, feeders_devices, region, relays_configured=True)
        else:
            feeders_devices, bu_devices = self.get_feeder_switches(feeder_list, region)
            self.chk_empty_fdrs(feeders_devices)
            user_selection = self.run_window(feeder_list, feeders_devices, region, relays_configured=False)
            feeders_devices = user_selection
        return feeders_devices, bu_devices, user_selection, external_grid

    def center_window(self, root, width, height):
        """Center the window on the user's screen"""
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        root.geometry(f"{width}x{height}+{x}+{y}")

    def mesh_feeder_check(self) -> Tuple[List[str], bool]:
        self.app.PrintPlain("Checking for radial feeders...")
        grids = self.app.GetCalcRelevantObjects('*.ElmXnet')
        all_feeders = self.app.GetCalcRelevantObjects('*.ElmFeeder')

        lines_oos = False
        for feeder in all_feeders:
            if self.get_lines_oos(feeder):
                lines_oos = True
                self.app.PrintWarn("WARNING: The following feeders have lines out of service. \n"
                                   "These feeders will not appear in the selections list if the lines serve as open points. \n"
                                   "To remedy this, return the lines to service and place an open switch at one end of each line.")
                break
        for feeder in all_feeders:
            if self.get_lines_oos(feeder):
                str_lines = [line.loc_name for line in self.get_lines_oos(feeder)]
                self.app.PrintWarn(f"Feeder: {feeder.loc_name}")
                self.app.PrintWarn(f"Lines out of service: {str_lines}")

        radial_list = [
            feeder.loc_name for feeder in all_feeders
            if not (set(feeder.obj_id.GetAll(1, 0)) & set(grids) and
                    set(feeder.obj_id.GetAll(0, 0)) & set(grids) and
                    not feeder.IsOutOfService())
        ]

        if radial_list:
            self.app.PrintPlain("Radial feeders detected.")
        else:
            self.show_no_radial_feeders_message()

        return sorted(radial_list), lines_oos

    def get_lines_oos(self, feeder):

        oos_lines = []
        line_prefixes = ["HV", "TR", "LN"]
        for grid in self.app.GetSummaryGrid().GetContents():
            oos_lines += [
                line
                for line in grid.obj_id.GetContents("*.ElmLne")
                if any(string in line.GetAttribute("loc_name") for string in line_prefixes)
                if line.IsOutOfService()
            ]

        fdr_lines_oos = []
        for line in oos_lines:
            line_terms = line.GetConnectedElements()
            if any(term in feeder.GetObjs('ElmTerm') for term in line_terms):
                fdr_lines_oos.append(line)
        return fdr_lines_oos

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

    def feeders_external_grid(self, radial_list: List[str], lines_oos) -> Tuple[List[str], Dict[str, List[float]]]:
        root = tk.Tk()

        def _window_dim():
            grids = [grid for grid in self.app.GetCalcRelevantObjects('*.ElmXnet') if grid.outserv == 0]
            grid_cols = len(grids)
            column_width = 360
            feeder_col = 285
            window_width = max((grid_cols * column_width + feeder_col), 700)
            if window_width > 1300:
                window_width = 1500
            # Height will be recomputed after widgets are built.
            window_height = 600  # provisional
            return window_width, window_height

        window_width, window_height = _window_dim()
        self.center_window(root, window_width, window_height)
        root.title("Distribution Fault Study")

        # Create main container frame
        main_frame = tk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True)

        canvas, inner_frame = self.setup_scrollable_frame(main_frame)

        # Create button frame at the bottom, outside the scroll area
        button_frame = tk.Frame(root)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=5)

        list_length, var, grid_entries, grids = self.populate_feeders(root, inner_frame, radial_list, button_frame, lines_oos)

        # After widgets exist, compute required height and resize window to avoid vertical scrollbar
        root.update_idletasks()
        content_h = inner_frame.winfo_reqheight() + 20  # requested content height
        buttons_h = button_frame.winfo_reqheight() + 12  # bottom buttons
        chrome_h = 24  # title/labels margin
        desired_h = content_h + buttons_h + chrome_h

        screen_h = root.winfo_screenheight()
        final_h = min(desired_h, int(screen_h * 0.90))  # cap at 90% of screen height
        self.center_window(root, window_width, final_h)

        root.mainloop()

        feeder_list = [radial_list[i] for i in range(list_length) if var[i].get() == 1]
        if not feeder_list:
            return self.window_error(radial_list, 3, lines_oos)

        new_grid_data = self.collect_grid_data(grid_entries, radial_list, lines_oos)
        if not self.validate_grid_data(new_grid_data):
            return self.window_error(radial_list, 1, lines_oos)
        self.update_grid_data(grids, new_grid_data)

        return feeder_list, new_grid_data

    def setup_scrollable_frame(self, parent):
        canvas = tk.Canvas(parent, borderwidth=0)
        inner_frame = tk.Frame(canvas)
        vsb = tk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        hsb = tk.Scrollbar(parent, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Initially hide scrollbars
        canvas.pack(fill="both", expand=True)
        canvas.create_window((4, 4), window=inner_frame, anchor="nw")

        # Bind events to show/hide scrollbars as needed
        inner_frame.bind("<Configure>", lambda event: self.onFrameConfigure(canvas, vsb, hsb))
        canvas.bind("<Configure>", lambda event: self.onCanvasConfigure(canvas, vsb, hsb))

        return canvas, inner_frame

    def populate_feeders(self, root, frame, radial_list, button_frame, lines_oos=False):
        # Row 0: Main title
        ttk.Label(frame, text="Select all feeders to study:", font='Helvetica 14 bold').grid(
            row=0, column=0, columnspan=3, padx=5, pady=5, sticky="w"
        )

        # Row 1: Mesh feeders excluded
        ttk.Label(frame, text="(mesh feeders excluded)", font='Helvetica 10 bold').grid(
            row=1, column=0, columnspan=3, padx=5, pady=5, sticky="w"
        )

        # Rows 2-3: Warning messages (if applicable) - placed BEFORE feeder checkboxes
        current_row = 2
        if lines_oos:
            ttk.Label(frame, text="Note: feeders with lines out of service were detected.",
                      font='Helvetica 10 bold', wraplength=250).grid(
                row=current_row, column=0, columnspan=3, padx=5, pady=2, sticky="w"
            )
            current_row += 1
            ttk.Label(frame, text="See output window for more information",
                      font='Helvetica 10 bold', wraplength=250).grid(
                row=current_row, column=0, columnspan=3, padx=5, pady=2, sticky="w"
            )
            current_row += 1

        grids = [grid for grid in self.app.GetCalcRelevantObjects('*.ElmXnet') if grid.outserv == 0]

        grid_data = self.get_grid_data(grids)
        self.app.PrintPlain("Please enter the requested inputs.")

        # Create a separate frame for feeder checkboxes starting after the warnings
        feeder_frame = tk.Frame(frame)
        feeder_frame.grid(row=current_row, column=0, columnspan=4, sticky="nw", padx=5, pady=5)

        var = self.create_feeder_checkboxes(feeder_frame, radial_list)

        frame.columnconfigure(4, minsize=100)

        # External grid interface starts at the same row as feeder checkboxes (current_row)
        grid_entries = self.create_external_grid_interface(frame, grids, grid_data, current_row)

        # Place buttons in the button_frame instead of the scrollable frame
        ttk.Button(button_frame, text='Okay', command=root.destroy).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text='Exit', command=lambda: self.exit_script(root)).pack(side=tk.LEFT, padx=5)

        return len(radial_list), var, grid_entries, grids
    # ------------------------------------------------

    def get_master_grid(self, grid_loc_name):

        project = self.app.GetActiveProject()
        derived_proj = project.GetAttribute('der_baseproject')
        if derived_proj is None:
            return False
        net_dat = derived_proj.GetContents("Network Model\\Network Data")
        for network in net_dat:
            elm_nets = network.GetContents("*.ElmNet")
            for elm_net in elm_nets:
                elm_substats = elm_net.GetContents("*.ElmSubstat")
                for elm_substat in elm_substats:
                    if elm_substat.GetContents(f'{grid_loc_name}.ElmXnet', 1):
                        master_proj_grids = elm_substat.GetContents(f'{grid_loc_name}.ElmXnet', 1)[0]
                        self.app.PrintPlain(f'Gathered master grid data for {grid_loc_name}.')
                        return master_proj_grids
        self.app.PrintPlain(f'Could not find master grid data for {grid_loc_name}.')
        return False


    def get_grid_data(self, grids):
        import math

        grid_data = {}
        attributes = ['ikss', 'rntxn', 'z2tz1', 'x0tx1', 'r0tx0', 'ikssmin', 'rntxnmin', 'z2tz1min', 'x0tx1min',
                      'r0tx0min']
        for grid in grids:
            grid_data[grid] = [grid.GetAttribute(attr) for attr in attributes]
            self.app.PrintPlain(f'Finding System normal source impedance for {grid}...')
            grid_loc_name = grid.GetAttribute('loc_name')
            master_grid = self.get_master_grid(grid_loc_name)
            if master_grid:
                grid_prw = master_grid.GetAttribute('snssmin')
                ikssmin = grid_prw / (11 * math.sqrt(3))
                master_grid_attr = ['rntxnmin', 'z2tz1min', 'x0tx1min', 'r0tx0min']
                master_grid_imp = [master_grid.GetAttribute(attr) for attr in master_grid_attr]
                grid_data[grid].append(ikssmin)
                grid_data[grid].extend(master_grid_imp)
            if len(grid_data[grid]) == 10:
                self.app.PrintPlain(f'Could not fine system normal source impedance for {grid}...')
                grid_data[grid].extend([0, 0, 0, 0, 0])
        return grid_data


    def create_feeder_checkboxes(self, feeder_frame, radial_list):
        var = []
        for i, feeder in enumerate(radial_list):
            var.append(tk.IntVar())
            ttk.Checkbutton(feeder_frame, text=feeder, variable=var[-1]).grid(row=i, column=0, sticky="w", padx=25,
                                                                              pady=5)
        return var

    def create_external_grid_interface(self, frame, grids, grid_data, start_row=0):
        ttk.Label(frame, text="Enter external grid data:", font='Helvetica 14 bold').grid(
            column=5, columnspan=3, sticky="w", row=0, padx=5, pady=5
        )

        # Create a dedicated frame for grid entries to avoid height conflicts
        grid_entries_frame = tk.Frame(frame)
        grid_entries_frame.grid(row=start_row, column=5, columnspan=10, sticky="nw", padx=5, pady=5)

        grid_entries = {}
        for i, grid in enumerate(grids):
            grid_entries[grid] = self.create_grid_entries(grid_entries_frame, grid, grid_data[grid], i * 3)
        return grid_entries

    def create_grid_entries(self, grid_frame, grid, data, column):
        grid_entries = []
        labels = [
            "P-P-P fault max.", "R/X max.", "Z2/Z1 max.", "X0/X1 max.", "R0/X0 max.",
            "P-P-P fault min.", "R/X min.", "Z2/Z1 min.", "X0/X1 min.", "R0/X0 min."
        ]

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
            value = round(value, 6)
            tk.Label(max_frame, text=label).grid(row=i, column=0, pady=2, sticky="w")
            var = tk.DoubleVar(value=value)
            tk.Entry(max_frame, textvariable=var).grid(row=i, column=1, pady=2, padx=(5, 0))
            if i in [0]:  # Only first entry (P-P-P fault max.) gets kA unit
                tk.Label(max_frame, text='kA').grid(row=i, column=2, ipadx=5, pady=2)
            grid_entries.append(var)

        # Create entries for minimum values (next 5 parameters)
        for i in range(5, 10):
            label, value = labels[i], data[i]
            value = round(value, 6)
            tk.Label(min_frame, text=label).grid(row=i - 5, column=0, pady=2, sticky="w")
            var = tk.DoubleVar(value=value)
            tk.Entry(min_frame, textvariable=var).grid(row=i - 5, column=1, pady=2, padx=(5, 0))
            if i == 5:  # Only P-P-P fault min. gets kA unit
                tk.Label(min_frame, text='kA').grid(row=i - 5, column=2, ipadx=5, pady=2)
            grid_entries.append(var)

        # Sys Norm Minimum
        sys_norm_min_frame = tk.LabelFrame(
            grid_frame, text=f"{grid.loc_name} Sys Norm Minimum Values", relief="solid", bd=1, padx=5, pady=5
        )
        sys_norm_min_frame.grid(row=2, column=column, columnspan=3, padx=5, pady=5, sticky="ew")

        for i in range(10, 15):
            label, value = labels[i - 5], data[i]
            value = round(value, 6)
            tk.Label(sys_norm_min_frame, text=label).grid(row=i - 10, column=0, pady=2, sticky="w")
            var = tk.DoubleVar(value=value)
            tk.Entry(sys_norm_min_frame, textvariable=var).grid(row=i - 10, column=1, pady=2, padx=(5, 0))
            if i == 10:
                tk.Label(sys_norm_min_frame, text='kA').grid(row=i - 10, column=2, ipadx=5, pady=2)
            grid_entries.append(var)

        if data[10] != 0:
            tk.Label(
                sys_norm_min_frame, text="Default Values Copied From Master Project"
            ).grid(row=6, columnspan=3, pady=2)

        return grid_entries

    def collect_grid_data(self, grid_entries, radial_list, lines_oos):
        new_grid_data = {}
        for grid in grid_entries:
            try:
                new_grid_data[grid] = [item.get() for item in grid_entries[grid]]
            except Exception:
                return self.window_error(radial_list, 2, lines_oos)
        return new_grid_data

    def validate_grid_data(self, new_grid_data):
        return all(new_grid_data[grid][0] <= 100
                   and new_grid_data[grid][5] <= 100
                   and new_grid_data[grid][10] <= 100 for grid in new_grid_data)

    def update_grid_data(self, grids, new_grid_data):
        for grid in grids:
            for attr, value in zip(
                ['ikss', 'rntxn', 'z2tz1', 'x0tx1', 'r0tx0', 'ikssmin', 'rntxnmin', 'z2tz1min', 'x0tx1min', 'r0tx0min'],
                new_grid_data[grid]
            ):
                setattr(grid, attr, value)
            grid.cmax = 1.1
            grid.cmin = 1
            # name = grid.loc_name
            # new_grid_data[name] = new_grid_data[grid]
            # del new_grid_data[grid]

    def window_error(self, radial_list, error_code, lines_oos):
        error_messages = {
            1: "Please enter fault level values in kA, not A",
            2: "Please enter numerical values only",
            3: "Please select at least one feeder to continue"
        }
        self.app.PrintPlain(error_messages[error_code])
        return self.feeders_external_grid(radial_list, lines_oos)

    def get_feeders_devices(self, radial_list: List[str]) -> Tuple[Dict[str,list],Dict[Any,list]]:
        """Get active relays and fuses.
        Map them to corresponding external grid or feeder using a dictionary"""

        # Filter for relays under network model recursively.
        all_relays = elements.get_all_relays(self.app)
        # Create a list of active fuses
        all_fuses = fuses.get_all_fuses(self.app)
        devices = all_relays + all_fuses

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


    def get_feeder_switches(self, feeder_list: List[str], region: str) -> Tuple[Dict[Any, Any], List[Any]]:
        """Get lists of all switches for all feeders"""
        bu_devices = []
        elm_coups = [switch for switch in self.app.GetCalcRelevantObjects("*.ElmCoup") if switch.on_off == 1]
        sta_switches = [switch for switch in self.app.GetCalcRelevantObjects("*.StaSwitch") if switch.on_off == 1]

        acr_prefix = ['(ACR/SECT/LBS)', '(CIRCUIT RECL)', '(CIRCUIT RECLOSER)', '(OIL CIRCUIT RECL)']
        edo_prefix = ['(EDO FUSE)']

        if region == 'Regional Models':
            relay_switches = [switch for switch in elm_coups + sta_switches
                              if any(substring in switch.loc_name for substring in acr_prefix)]
            line_fuse_switches = [switch for switch in sta_switches
                                  if any(substring in switch.loc_name for substring in edo_prefix)]
            all_switches = relay_switches + line_fuse_switches
        else:  # SEQ Models
            all_switches = sta_switches

        feeders_switches = {}
        for feeder in feeder_list:
            feeder_elm = self.app.GetCalcRelevantObjects(feeder + ".ElmFeeder")[0]
            switch_list = [
                switch for switch in all_switches
                if (switch.GetClassName() == 'StaSwitch' and switch.fold_id.cterm in feeder_elm.GetObjs('ElmTerm')) or
                   (switch.GetClassName() == 'ElmCoup' and switch in feeder_elm.GetObjs('ElmCoup'))
            ]

            if region == 'SEQ Models':
                # def name(switch_obj):
                #     return switch_obj.loc_name
                # switch_list.sort(key=name)
                switch_list.sort(key=lambda switch: switch.loc_name)

            feeders_switches[feeder] = [feeder_elm] + switch_list

        return feeders_switches, bu_devices


    def chk_empty_fdrs(self, fdrs_devices):
        """
        Check that the selected feeders have protection devices created.
        :param self:
        :param fdrs_devices:
        :return:
        """

        empty_feeders = [feeder for feeder, devices in fdrs_devices.items() if devices == []]

        if len(empty_feeders) == len(fdrs_devices):
            self.app.PrintError("No protection devices were detected in the model for the selected feeders. \n"
                           "Please add and configure the required protection devices and re-run the script.")
            sys.exit(0)
        for empty_feeder in empty_feeders:
            self.app.PrintWarn(f"No protection devices were detected in the model for feeder {empty_feeder}. \n"
                          "This feeder will be excluded from the study.")
            del fdrs_devices[empty_feeder]


    def populate(self, frame, feeder_list, feeders_switches, region, button_frame, relays_configured):
        if region == 'Regional Models':
            fdr_sw_locname = {feeder: [switch.loc_name for switch in switches]
                              for feeder, switches in feeders_switches.items()}
        else:  # region == 'SEQ Models'
            fdr_sw_locname = {feeder: [] for feeder in feeders_switches}
            for feeder, switches in feeders_switches.items():
                for switch in switches:
                    if switch.GetClassName() == ElementType.FEEDER.value:
                        cubicle = switch.obj_id
                    else:
                        cubicle = switch.fold_id
                    switch_term = cubicle.cterm.loc_name
                    fdr_sw_locname[feeder].append(switch_term[:-5] if switch_term.endswith("_Term") else switch_term)

        if relays_configured:
            ttk.Label(frame, text="Select all protection devices to study:",
                      font='Helvetica 12 bold').grid(columnspan=8, padx=5, pady=5)
        else:
            ttk.Label(frame, text="Select ALL protection devices:",
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

    def run_window(self, feeder_list: List[str], feeders_devices: Dict[str, List[Any]], region: str, relays_configured: bool) -> Dict[
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

        var, fdr_dev_locname = self.populate(frame, feeder_list, feeders_devices, region, button_frame, relays_configured)

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


def get_input(app: pft.Application, region: str, study_selections: List):
    study = FaultLevelStudy(app)
    return study.main(region, study_selections)
