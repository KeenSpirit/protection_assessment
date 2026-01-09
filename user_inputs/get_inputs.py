"""
Feeder and device selection dialog for protection assessment.

This module provides GUI dialogs for users to select feeders and
protection devices for fault level studies. It handles three main
user interaction phases:

1. Feeder Selection: Displays available radial feeders (excluding
   mesh configurations) with checkboxes for multi-selection.

2. External Grid Data: Collects and validates fault level parameters
   for external grid elements including maximum/minimum fault levels
   and impedance ratios.

3. Device Selection: Presents protection devices (relays, fuses,
   switches) organized by feeder for user selection.

The module automatically detects and excludes mesh feeders, warns
about lines out of service, and validates all user inputs before
proceeding with fault studies.

Classes:
    FaultLevelStudy: Main class orchestrating user input collection.

Functions:
    get_input: Module entry point for user input collection workflow.

Example:
    >>> from user_inputs import get_inputs
    >>> feeders, bu_devices, selection, grid = get_inputs.get_input(
    ...     app, 'SEQ', ['Fault Level Study (all relays configured)']
    ... )
"""

from tkinter import *  # noqa: F403
import math
import sys
import tkinter as tk
from tkinter import ttk
from typing import Any, Dict, List, Tuple

from pf_config import pft
from domain.enums import ElementType
from devices import fuses
from relays import elements


class FaultLevelStudy:
    """
    Orchestrates user input collection for fault level studies.

    This class manages the complete workflow of collecting user inputs
    including feeder selection, external grid parameter entry, and
    protection device selection through tkinter GUI dialogs.

    Attributes:
        app: PowerFactory application instance for model queries
            and output messaging.

    Example:
        >>> study = FaultLevelStudy(app)
        >>> feeders, bu_devices, selection, grid = study.main(
        ...     'SEQ', ['Fault Level Study']
        ... )
    """

    def __init__(self, app: pft.Application) -> None:
        """
        Initialize the fault level study instance.

        Args:
            app: PowerFactory application instance.
        """
        self.app = app

    def main(
        self, region: str, study_selections: List[str]
    ) -> Tuple[Dict, Dict, Dict, Dict]:
        """
        Execute the complete user input collection workflow.

        Orchestrates the sequence of user dialogs: feeder filtering,
        feeder/grid selection, and device selection based on the
        selected study type.

        Args:
            region: Network region identifier ('SEQ' or
                'Regional Models').
            study_selections: List of selected study type strings.

        Returns:
            Tuple containing:
                - feeders_devices: Dict mapping feeder names to lists
                  of protection device objects.
                - bu_devices: Dict mapping grid objects to lists of
                  backup device objects.
                - user_selection: Dict of user-selected devices per
                  feeder.
                - external_grid: Dict of grid objects to fault level
                  parameter lists.
        """
        radial_list, lines_oos = self.mesh_feeder_check()
        feeder_list, external_grid = self.feeders_external_grid(
            radial_list, lines_oos
        )

        if 'Fault Level Study (all relays configured in model)' in (
            study_selections
        ):
            feeders_devices, bu_devices = self.get_feeders_devices(feeder_list)
            self.chk_empty_fdrs(feeders_devices)
            user_selection = self.run_window(
                feeder_list, feeders_devices, region, relays_configured=True
            )
        else:
            feeders_devices, bu_devices = self.get_feeder_switches(
                feeder_list, region
            )
            self.chk_empty_fdrs(feeders_devices)
            user_selection = self.run_window(
                feeder_list, feeders_devices, region, relays_configured=False
            )
            feeders_devices = user_selection

        return feeders_devices, bu_devices, user_selection, external_grid

    def center_window(
        self, root: tk.Tk, width: int, height: int
    ) -> None:
        """
        Center a tkinter window on the user's screen.

        Args:
            root: The tkinter root window to center.
            width: Desired window width in pixels.
            height: Desired window height in pixels.
        """
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        root.geometry(f"{width}x{height}+{x}+{y}")

    def mesh_feeder_check(self) -> Tuple[List[str], bool]:
        """
        Filter feeders to exclude mesh configurations.

        Identifies radial feeders by checking external grid
        connectivity. Mesh feeders (connected to grids at both ends)
        are excluded. Also detects lines out of service which may
        affect feeder topology.

        Returns:
            Tuple containing:
                - radial_list: Sorted list of radial feeder names.
                - lines_oos: True if any lines are out of service.
        """
        self.app.PrintPlain("Checking for radial feeders...")
        grids = self.app.GetCalcRelevantObjects('*.ElmXnet')
        all_feeders = self.app.GetCalcRelevantObjects('*.ElmFeeder')

        lines_oos = False
        for feeder in all_feeders:
            if self.get_lines_oos(feeder):
                lines_oos = True
                self.app.PrintWarn(
                    "WARNING: The following feeders have lines out of "
                    "service.\n"
                    "These feeders will not appear in the selections list "
                    "if the lines serve as open points. \n"
                    "To remedy this, return the lines to service and place "
                    "an open switch at one end of each line."
                )
                break

        for feeder in all_feeders:
            if self.get_lines_oos(feeder):
                str_lines = [line.loc_name for line in self.get_lines_oos(
                    feeder
                )]
                self.app.PrintWarn(f"Feeder: {feeder.loc_name}")
                self.app.PrintWarn(f"Lines out of service: {str_lines}")

        radial_list = [
            feeder.loc_name for feeder in all_feeders
            if not (
                set(feeder.obj_id.GetAll(1, 0)) & set(grids)
                and set(feeder.obj_id.GetAll(0, 0)) & set(grids)
                and not feeder.IsOutOfService()
            )
        ]

        if radial_list:
            self.app.PrintPlain("Radial feeders detected.")
        else:
            self.show_no_radial_feeders_message()

        return sorted(radial_list), lines_oos

    def get_lines_oos(self, feeder: pft.ElmFeeder) -> List:
        """
        Get lines out of service for a given feeder.

        Identifies HV, TR, and LN prefixed lines that are out of
        service and connected to the specified feeder.

        Args:
            feeder: PowerFactory feeder element to check.

        Returns:
            List of out-of-service line elements connected to the
            feeder. Empty list if none found.
        """
        oos_lines = []
        line_prefixes = ["HV", "TR", "LN"]

        for grid in self.app.GetSummaryGrid().GetContents():
            oos_lines += [
                line
                for line in grid.obj_id.GetContents("*.ElmLne")
                if any(
                    string in line.GetAttribute("loc_name")
                    for string in line_prefixes
                )
                if line.IsOutOfService()
            ]

        fdr_lines_oos = []
        for line in oos_lines:
            line_terms = line.GetConnectedElements()
            if any(term in feeder.GetObjs('ElmTerm') for term in line_terms):
                fdr_lines_oos.append(line)

        return fdr_lines_oos

    def show_no_radial_feeders_message(self) -> None:
        """
        Display error dialog when no radial feeders are found.

        Creates a modal dialog informing the user that no radial
        feeders were detected and the script cannot proceed.
        """
        root = tk.Tk()
        root.title("Distribution fault study")

        window_width = 400
        window_height = 150
        self.center_window(root, window_width, window_height)

        ttk.Label(
            root, text="No radial feeders were found at the substation"
        ).grid(padx=5, pady=5)
        ttk.Label(
            root,
            text="To run this script, please radialise one or more feeders"
        ).grid(padx=5, pady=5)
        ttk.Button(
            root, text='Exit', command=lambda: self.exit_script(root)
        ).grid(sticky="s", padx=5, pady=5)

        root.mainloop()

    def exit_script(self, root: tk.Tk) -> None:
        """
        Clean exit handler for GUI dialogs.

        Prints termination message, destroys the window, and exits
        the script.

        Args:
            root: The tkinter root window to destroy.
        """
        self.app.PrintPlain("User terminated script.")
        root.destroy()
        sys.exit(0)

    def feeders_external_grid(
        self, radial_list: List[str], lines_oos: bool
    ) -> Tuple[List[str], Dict[str, List[float]]]:
        """
        Display combined feeder selection and external grid data dialog.

        Creates a scrollable dialog with feeder checkboxes and external
        grid parameter entry fields. Validates all inputs before
        returning.

        Args:
            radial_list: List of available radial feeder names.
            lines_oos: Flag indicating if lines are out of service.

        Returns:
            Tuple containing:
                - feeder_list: List of user-selected feeder names.
                - new_grid_data: Dict mapping grid objects to lists of
                  validated fault level parameters.
        """
        root = tk.Tk()

        def _window_dim():
            grids = [
                grid for grid in self.app.GetCalcRelevantObjects('*.ElmXnet')
                if grid.outserv == 0
            ]
            grid_cols = len(grids)
            column_width = 360
            feeder_col = 285
            window_width = max((grid_cols * column_width + feeder_col), 700)
            if window_width > 1300:
                window_width = 1500
            window_height = 600
            return window_width, window_height

        window_width, window_height = _window_dim()
        self.center_window(root, window_width, window_height)
        root.title("Distribution Fault Study")

        main_frame = tk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True)

        canvas, inner_frame = self.setup_scrollable_frame(main_frame)

        button_frame = tk.Frame(root)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=5)

        list_length, var, grid_entries, grids = self.populate_feeders(
            root, inner_frame, radial_list, button_frame, lines_oos
        )

        # Resize window to fit content
        root.update_idletasks()
        content_h = inner_frame.winfo_reqheight() + 20
        buttons_h = button_frame.winfo_reqheight() + 12
        chrome_h = 24
        desired_h = content_h + buttons_h + chrome_h

        screen_h = root.winfo_screenheight()
        final_h = min(desired_h, int(screen_h * 0.90))
        self.center_window(root, window_width, final_h)

        root.mainloop()

        feeder_list = [
            radial_list[i] for i in range(list_length) if var[i].get() == 1
        ]
        if not feeder_list:
            return self.window_error(radial_list, 3, lines_oos)

        new_grid_data = self.collect_grid_data(
            grid_entries, radial_list, lines_oos
        )
        if not self.validate_grid_data(new_grid_data):
            return self.window_error(radial_list, 1, lines_oos)

        self.update_grid_data(grids, new_grid_data)

        return feeder_list, new_grid_data

    def setup_scrollable_frame(
        self, parent: tk.Frame
    ) -> Tuple[tk.Canvas, tk.Frame]:
        """
        Create a scrollable frame within the parent container.

        Sets up a canvas with vertical and horizontal scrollbars that
        dynamically show/hide based on content size.

        Args:
            parent: Parent tkinter frame to contain the scrollable
                area.

        Returns:
            Tuple containing:
                - canvas: The canvas widget for scrolling.
                - inner_frame: The frame inside the canvas for content.
        """
        canvas = tk.Canvas(parent, borderwidth=0)
        inner_frame = tk.Frame(canvas)
        vsb = tk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        hsb = tk.Scrollbar(parent, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        canvas.pack(fill="both", expand=True)
        canvas.create_window((4, 4), window=inner_frame, anchor="nw")

        inner_frame.bind(
            "<Configure>",
            lambda event: self.onFrameConfigure(canvas, vsb, hsb)
        )
        canvas.bind(
            "<Configure>",
            lambda event: self.onCanvasConfigure(canvas, vsb, hsb)
        )

        return canvas, inner_frame

    @staticmethod
    def onFrameConfigure(
        canvas: tk.Canvas, vsb: tk.Scrollbar, hsb: tk.Scrollbar
    ) -> None:
        """
        Update scroll region and show/hide scrollbars as needed.

        Called when the inner frame is resized to update the canvas
        scroll region.

        Args:
            canvas: The canvas widget.
            vsb: Vertical scrollbar widget.
            hsb: Horizontal scrollbar widget.
        """
        canvas.configure(scrollregion=canvas.bbox("all"))
        FaultLevelStudy.update_scrollbars(canvas, vsb, hsb)

    @staticmethod
    def onCanvasConfigure(
        canvas: tk.Canvas, vsb: tk.Scrollbar, hsb: tk.Scrollbar
    ) -> None:
        """
        Handle canvas resize events.

        Updates scrollbar visibility when the canvas is resized.

        Args:
            canvas: The canvas widget.
            vsb: Vertical scrollbar widget.
            hsb: Horizontal scrollbar widget.
        """
        FaultLevelStudy.update_scrollbars(canvas, vsb, hsb)

    @staticmethod
    def update_scrollbars(
        canvas: tk.Canvas, vsb: tk.Scrollbar, hsb: tk.Scrollbar
    ) -> None:
        """
        Show or hide scrollbars based on content size vs canvas size.

        Dynamically displays scrollbars only when content exceeds the
        visible canvas area.

        Args:
            canvas: The canvas widget.
            vsb: Vertical scrollbar widget.
            hsb: Horizontal scrollbar widget.
        """
        scrollregion = canvas.cget("scrollregion")
        if not scrollregion:
            return

        x1, y1, x2, y2 = map(float, scrollregion.split())
        content_width = x2 - x1
        content_height = y2 - y1

        canvas_width = canvas.winfo_width()
        canvas_height = canvas.winfo_height()

        if content_height > canvas_height:
            vsb.pack(side="right", fill="y", before=canvas)
        else:
            vsb.pack_forget()

        if content_width > canvas_width:
            hsb.pack(side="bottom", fill="x", before=canvas)
        else:
            hsb.pack_forget()

        canvas.pack(fill="both", expand=True)

    def populate_feeders(
        self,
        root: tk.Tk,
        frame: tk.Frame,
        radial_list: List[str],
        button_frame: tk.Frame,
        lines_oos: bool = False
    ) -> Tuple[int, List[tk.IntVar], Dict, List]:
        """
        Create feeder selection and external grid entry widgets.

        Populates the dialog frame with feeder checkboxes and external
        grid parameter entry fields.

        Args:
            root: Root tkinter window.
            frame: Frame to contain the widgets.
            radial_list: List of radial feeder names.
            button_frame: Frame for action buttons.
            lines_oos: Flag indicating if lines are out of service.

        Returns:
            Tuple containing:
                - list_length: Number of feeders in the list.
                - var: List of IntVar checkbox variables.
                - grid_entries: Dict of grid entry field variables.
                - grids: List of active external grid objects.
        """
        ttk.Label(
            frame,
            text="Select all feeders to study:",
            font='Helvetica 14 bold'
        ).grid(row=0, column=0, columnspan=3, padx=5, pady=5, sticky="w")

        ttk.Label(
            frame,
            text="(mesh feeders excluded)",
            font='Helvetica 10 bold'
        ).grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky="w")

        current_row = 2
        if lines_oos:
            ttk.Label(
                frame,
                text="Note: feeders with lines out of service were detected.",
                font='Helvetica 10 bold',
                wraplength=250
            ).grid(
                row=current_row, column=0, columnspan=3,
                padx=5, pady=2, sticky="w"
            )
            current_row += 1
            ttk.Label(
                frame,
                text="See output window for more information",
                font='Helvetica 10 bold',
                wraplength=250
            ).grid(
                row=current_row, column=0, columnspan=3,
                padx=5, pady=2, sticky="w"
            )
            current_row += 1

        grids = [
            grid for grid in self.app.GetCalcRelevantObjects('*.ElmXnet')
            if grid.outserv == 0
        ]
        grid_data = self.get_grid_data(grids)
        self.app.PrintPlain("Please enter the requested inputs.")

        feeder_frame = tk.Frame(frame)
        feeder_frame.grid(
            row=current_row, column=0, columnspan=4, sticky="nw", padx=5, pady=5
        )

        var = self.create_feeder_checkboxes(feeder_frame, radial_list)

        frame.columnconfigure(4, minsize=100)

        grid_entries = self.create_external_grid_interface(
            frame, grids, grid_data, current_row
        )

        ttk.Button(
            button_frame, text='Okay', command=root.destroy
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            button_frame, text='Exit', command=lambda: self.exit_script(root)
        ).pack(side=tk.LEFT, padx=5)

        return len(radial_list), var, grid_entries, grids

    def get_master_grid(self, grid_loc_name: str):
        """
        Retrieve master project grid data for system normal minimum.

        Searches the derived base project for matching external grid
        elements to obtain system normal minimum fault level data.

        Args:
            grid_loc_name: Location name of the grid to find.

        Returns:
            The master project grid object if found, False otherwise.
        """
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
                    grids = elm_substat.GetContents(
                        f'{grid_loc_name}.ElmXnet', 1
                    )
                    if grids:
                        master_proj_grids = grids[0]
                        self.app.PrintPlain(
                            f'Gathered master grid data for {grid_loc_name}.'
                        )
                        return master_proj_grids

        self.app.PrintPlain(
            f'Could not find master grid data for {grid_loc_name}.'
        )
        return False

    def get_grid_data(self, grids: List) -> Dict:
        """
        Collect fault level parameters from external grid elements.

        Retrieves maximum and minimum fault level attributes from each
        grid, including system normal minimum values from the master
        project if available.

        Args:
            grids: List of external grid (ElmXnet) objects.

        Returns:
            Dict mapping grid objects to lists of 15 fault level
            parameters: [ikss, rntxn, z2tz1, x0tx1, r0tx0] for max,
            min, and system normal minimum conditions.
        """
        grid_data = {}
        attributes = [
            'ikss', 'rntxn', 'z2tz1', 'x0tx1', 'r0tx0',
            'ikssmin', 'rntxnmin', 'z2tz1min', 'x0tx1min', 'r0tx0min'
        ]

        for grid in grids:
            grid_data[grid] = [
                grid.GetAttribute(attr) for attr in attributes
            ]
            self.app.PrintPlain(
                f'Finding System normal source impedance for {grid}...'
            )
            grid_loc_name = grid.GetAttribute('loc_name')
            master_grid = self.get_master_grid(grid_loc_name)

            if master_grid:
                grid_prw = master_grid.GetAttribute('snssmin')
                ikssmin = grid_prw / (11 * math.sqrt(3))
                master_grid_attr = [
                    'rntxnmin', 'z2tz1min', 'x0tx1min', 'r0tx0min'
                ]
                master_grid_imp = [
                    master_grid.GetAttribute(attr) for attr in master_grid_attr
                ]
                grid_data[grid].append(ikssmin)
                grid_data[grid].extend(master_grid_imp)

            if len(grid_data[grid]) == 10:
                self.app.PrintPlain(
                    f'Could not find system normal source impedance '
                    f'for {grid}...'
                )
                grid_data[grid].extend([0, 0, 0, 0, 0])

        return grid_data

    def create_feeder_checkboxes(
        self, feeder_frame: tk.Frame, radial_list: List[str]
    ) -> List[tk.IntVar]:
        """
        Create checkbox widgets for feeder selection.

        Args:
            feeder_frame: Frame to contain the checkboxes.
            radial_list: List of feeder names to display.

        Returns:
            List of IntVar variables linked to each checkbox.
        """
        var = []
        for i, feeder in enumerate(radial_list):
            var.append(tk.IntVar())
            ttk.Checkbutton(
                feeder_frame, text=feeder, variable=var[-1]
            ).grid(row=i, column=0, sticky="w", padx=25, pady=5)
        return var

    def create_external_grid_interface(
        self,
        frame: tk.Frame,
        grids: List,
        grid_data: Dict,
        start_row: int = 0
    ) -> Dict:
        """
        Create external grid parameter entry interface.

        Generates labeled entry fields for fault level parameters
        organized by grid element.

        Args:
            frame: Parent frame for the interface.
            grids: List of external grid objects.
            grid_data: Dict of current grid parameter values.
            start_row: Starting row for grid placement.

        Returns:
            Dict mapping grid objects to lists of entry field
            variables.
        """
        ttk.Label(
            frame,
            text="Enter external grid data:",
            font='Helvetica 14 bold'
        ).grid(column=5, columnspan=3, sticky="w", row=0, padx=5, pady=5)

        grid_entries_frame = tk.Frame(frame)
        grid_entries_frame.grid(
            row=start_row, column=5, columnspan=10,
            sticky="nw", padx=5, pady=5
        )

        grid_entries = {}
        for i, grid in enumerate(grids):
            grid_entries[grid] = self.create_grid_entries(
                grid_entries_frame, grid, grid_data[grid], i * 3
            )
        return grid_entries

    def create_grid_entries(
        self,
        grid_frame: tk.Frame,
        grid: Any,
        data: List[float],
        column: int
    ) -> List[tk.DoubleVar]:
        """
        Create entry fields for a single external grid's parameters.

        Generates three labeled frames (Maximum, Minimum, System Normal
        Minimum) each containing five parameter entry fields.

        Args:
            grid_frame: Parent frame for the entry fields.
            grid: External grid object.
            data: List of 15 current parameter values.
            column: Starting column position.

        Returns:
            List of 15 DoubleVar variables for the entry fields.
        """
        grid_entries = []
        labels = [
            "P-P-P fault max.", "R/X max.", "Z2/Z1 max.",
            "X0/X1 max.", "R0/X0 max.",
            "P-P-P fault min.", "R/X min.", "Z2/Z1 min.",
            "X0/X1 min.", "R0/X0 min."
        ]

        # Maximum values frame
        max_frame = tk.LabelFrame(
            grid_frame,
            text=f"{grid.loc_name} Maximum Values",
            relief="solid", bd=1, padx=5, pady=5
        )
        max_frame.grid(
            row=0, column=column, columnspan=3, padx=5, pady=5, sticky="ew"
        )

        for i in range(5):
            label, value = labels[i], data[i]
            value = round(value, 6)
            tk.Label(max_frame, text=label).grid(
                row=i, column=0, pady=2, sticky="w"
            )
            var = tk.DoubleVar(value=value)
            tk.Entry(max_frame, textvariable=var).grid(
                row=i, column=1, pady=2, padx=(5, 0)
            )
            if i == 0:
                tk.Label(max_frame, text='kA').grid(
                    row=i, column=2, ipadx=5, pady=2
                )
            grid_entries.append(var)

        # Minimum values frame
        min_frame = tk.LabelFrame(
            grid_frame,
            text=f"{grid.loc_name} Minimum Values",
            relief="solid", bd=1, padx=5, pady=5
        )
        min_frame.grid(
            row=1, column=column, columnspan=3, padx=5, pady=5, sticky="ew"
        )

        for i in range(5, 10):
            label, value = labels[i], data[i]
            value = round(value, 6)
            tk.Label(min_frame, text=label).grid(
                row=i - 5, column=0, pady=2, sticky="w"
            )
            var = tk.DoubleVar(value=value)
            tk.Entry(min_frame, textvariable=var).grid(
                row=i - 5, column=1, pady=2, padx=(5, 0)
            )
            if i == 5:
                tk.Label(min_frame, text='kA').grid(
                    row=i - 5, column=2, ipadx=5, pady=2
                )
            grid_entries.append(var)

        # System normal minimum frame
        sys_norm_min_frame = tk.LabelFrame(
            grid_frame,
            text=f"{grid.loc_name} Sys Norm Minimum Values",
            relief="solid", bd=1, padx=5, pady=5
        )
        sys_norm_min_frame.grid(
            row=2, column=column, columnspan=3, padx=5, pady=5, sticky="ew"
        )

        for i in range(10, 15):
            label, value = labels[i - 5], data[i]
            value = round(value, 6)
            tk.Label(sys_norm_min_frame, text=label).grid(
                row=i - 10, column=0, pady=2, sticky="w"
            )
            var = tk.DoubleVar(value=value)
            tk.Entry(sys_norm_min_frame, textvariable=var).grid(
                row=i - 10, column=1, pady=2, padx=(5, 0)
            )
            if i == 10:
                tk.Label(sys_norm_min_frame, text='kA').grid(
                    row=i - 10, column=2, ipadx=5, pady=2
                )
            grid_entries.append(var)

        if data[10] != 0:
            tk.Label(
                sys_norm_min_frame,
                text="Default Values Copied From Master Project"
            ).grid(row=6, columnspan=3, pady=2)

        return grid_entries

    def collect_grid_data(
        self,
        grid_entries: Dict,
        radial_list: List[str],
        lines_oos: bool
    ) -> Dict:
        """
        Collect values from grid entry fields.

        Reads all entry field values and handles conversion errors
        by re-prompting the user.

        Args:
            grid_entries: Dict of grid objects to entry field lists.
            radial_list: List of radial feeder names (for error
                recovery).
            lines_oos: Lines out of service flag (for error recovery).

        Returns:
            Dict mapping grid objects to lists of entered values.
        """
        new_grid_data = {}
        for grid in grid_entries:
            try:
                new_grid_data[grid] = [
                    item.get() for item in grid_entries[grid]
                ]
            except Exception:
                return self.window_error(radial_list, 2, lines_oos)
        return new_grid_data

    def validate_grid_data(self, new_grid_data: Dict) -> bool:
        """
        Validate that fault level values are in correct units.

        Checks that fault current values are in kA (not A) by
        verifying they are less than 100.

        Args:
            new_grid_data: Dict of grid parameter values to validate.

        Returns:
            True if all values are valid, False otherwise.
        """
        return all(
            new_grid_data[grid][0] <= 100
            and new_grid_data[grid][5] <= 100
            and new_grid_data[grid][10] <= 100
            for grid in new_grid_data
        )

    def update_grid_data(self, grids: List, new_grid_data: Dict) -> None:
        """
        Apply validated grid data back to PowerFactory objects.

        Updates external grid element attributes with user-entered
        values and sets c-factor values.

        Args:
            grids: List of external grid objects to update.
            new_grid_data: Dict of validated parameter values.
        """
        attributes = [
            'ikss', 'rntxn', 'z2tz1', 'x0tx1', 'r0tx0',
            'ikssmin', 'rntxnmin', 'z2tz1min', 'x0tx1min', 'r0tx0min'
        ]

        for grid in grids:
            for attr, value in zip(attributes, new_grid_data[grid]):
                setattr(grid, attr, value)
            grid.cmax = 1.1
            grid.cmin = 1

    def window_error(
        self, radial_list: List[str], error_code: int, lines_oos: bool
    ) -> Tuple[List[str], Dict]:
        """
        Display error message and re-prompt user for input.

        Shows an appropriate error message based on the error code
        and recursively calls the input dialog.

        Args:
            radial_list: List of radial feeder names.
            error_code: Error type:
                1 = Fault level in wrong units (A instead of kA)
                2 = Non-numerical value entered
                3 = No feeder selected
            lines_oos: Lines out of service flag.

        Returns:
            Result of re-prompting feeders_external_grid dialog.
        """
        error_messages = {
            1: "Please enter fault level values in kA, not A",
            2: "Please enter numerical values only",
            3: "Please select at least one feeder to continue"
        }
        self.app.PrintPlain(error_messages[error_code])
        return self.feeders_external_grid(radial_list, lines_oos)

    def get_feeders_devices(
        self, radial_list: List[str]
    ) -> Tuple[Dict[str, list], Dict[Any, list]]:
        """
        Get active relays and fuses mapped to feeders and grids.

        Retrieves all configured protection devices and maps them to
        their associated feeders or external grid backup positions.

        Args:
            radial_list: List of radial feeder names.

        Returns:
            Tuple containing:
                - feeder_device_dict: Dict mapping feeder names to
                  lists of protection device objects.
                - grid_device_dict: Dict mapping grid objects to lists
                  of backup device objects.
        """
        all_relays = elements.get_all_relays(self.app)
        all_fuses = fuses.get_all_fuses(self.app)
        devices = all_relays + all_fuses

        feeder_device_dict = {feeder: [] for feeder in radial_list}
        grid_device_dict = {
            grid: []
            for grid in self.app.GetCalcRelevantObjects('*.ElmXnet')
            if grid.bus1 is not None
        }

        for device in devices:
            term = device.cbranch
            feeder = [
                feeder for feeder in radial_list
                if term in self.app.GetCalcRelevantObjects(
                    feeder + ".ElmFeeder"
                )[0].GetAll()
            ]
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

    def get_feeder_switches(
        self, feeder_list: List[str], region: str
    ) -> Tuple[Dict[Any, Any], List[Any]]:
        """
        Get lists of all switches for all feeders.

        Retrieves switch elements based on regional naming conventions
        and maps them to their associated feeders.

        Args:
            feeder_list: List of feeder names.
            region: Network region ('SEQ' or 'Regional Models').

        Returns:
            Tuple containing:
                - feeders_switches: Dict mapping feeder names to lists
                  of switch objects (with feeder element first).
                - bu_devices: Empty list (backup devices not applicable
                  for switch mode).
        """
        bu_devices = []
        elm_coups = [
            switch for switch in self.app.GetCalcRelevantObjects("*.ElmCoup")
            if switch.on_off == 1
        ]
        sta_switches = [
            switch for switch
            in self.app.GetCalcRelevantObjects("*.StaSwitch")
            if switch.on_off == 1
        ]

        acr_prefix = [
            '(ACR/SECT/LBS)', '(CIRCUIT RECL)',
            '(CIRCUIT RECLOSER)', '(OIL CIRCUIT RECL)'
        ]
        edo_prefix = ['(EDO FUSE)']

        if region == 'Regional Models':
            relay_switches = [
                switch for switch in elm_coups + sta_switches
                if any(
                    substring in switch.loc_name for substring in acr_prefix
                )
            ]
            line_fuse_switches = [
                switch for switch in sta_switches
                if any(
                    substring in switch.loc_name for substring in edo_prefix
                )
            ]
            all_switches = relay_switches + line_fuse_switches
        else:
            all_switches = sta_switches

        feeders_switches = {}
        for feeder in feeder_list:
            feeder_elm = self.app.GetCalcRelevantObjects(
                feeder + ".ElmFeeder"
            )[0]
            switch_list = [
                switch for switch in all_switches
                if (
                    switch.GetClassName() == 'StaSwitch'
                    and switch.fold_id.cterm in feeder_elm.GetObjs('ElmTerm')
                ) or (
                    switch.GetClassName() == 'ElmCoup'
                    and switch in feeder_elm.GetObjs('ElmCoup')
                )
            ]

            if region == 'SEQ Models':
                switch_list.sort(key=lambda switch: switch.loc_name)

            feeders_switches[feeder] = [feeder_elm] + switch_list

        return feeders_switches, bu_devices

    def chk_empty_fdrs(self, fdrs_devices: Dict) -> None:
        """
        Check that selected feeders have protection devices.

        Validates that at least one feeder has devices and removes
        empty feeders from the selection with warnings.

        Args:
            fdrs_devices: Dict of feeder names to device lists.

        Raises:
            SystemExit: If no feeders have any protection devices.
        """
        empty_feeders = [
            feeder for feeder, devices in fdrs_devices.items()
            if devices == []
        ]

        if len(empty_feeders) == len(fdrs_devices):
            self.app.PrintError(
                "No protection devices were detected in the model for the "
                "selected feeders. \n"
                "Please add and configure the required protection devices "
                "and re-run the script."
            )
            sys.exit(0)

        for empty_feeder in empty_feeders:
            self.app.PrintWarn(
                f"No protection devices were detected in the model for "
                f"feeder {empty_feeder}. \n"
                "This feeder will be excluded from the study."
            )
            del fdrs_devices[empty_feeder]

    def populate(
        self,
        frame: tk.Frame,
        feeder_list: List[str],
        feeders_switches: Dict,
        region: str,
        button_frame: tk.Frame,
        relays_configured: bool
    ) -> Tuple[List[tk.IntVar], Dict[str, List[str]]]:
        """
        Create device selection checkboxes for each feeder.

        Generates a grid of checkboxes organized by feeder columns
        for user device selection.

        Args:
            frame: Parent frame for the checkboxes.
            feeder_list: List of feeder names.
            feeders_switches: Dict of feeder names to device lists.
            region: Network region for display name formatting.
            button_frame: Frame for action buttons.
            relays_configured: True if relays are pre-configured.

        Returns:
            Tuple containing:
                - var: List of IntVar checkbox variables.
                - fdr_sw_locname: Dict mapping feeders to device
                  display names.
        """
        if region == 'Regional Models':
            fdr_sw_locname = {
                feeder: [switch.loc_name for switch in switches]
                for feeder, switches in feeders_switches.items()
            }
        else:
            fdr_sw_locname = {feeder: [] for feeder in feeders_switches}
            for feeder, switches in feeders_switches.items():
                for switch in switches:
                    if switch.GetClassName() == ElementType.FEEDER.value:
                        cubicle = switch.obj_id
                    else:
                        cubicle = switch.fold_id
                    switch_term = cubicle.cterm.loc_name
                    display_name = (
                        switch_term[:-5]
                        if switch_term.endswith("_Term")
                        else switch_term
                    )
                    fdr_sw_locname[feeder].append(display_name)

        if relays_configured:
            ttk.Label(
                frame,
                text="Select all protection devices to study:",
                font='Helvetica 12 bold'
            ).grid(columnspan=8, padx=5, pady=5)
        else:
            ttk.Label(
                frame,
                text="Select ALL protection devices:",
                font='Helvetica 12 bold'
            ).grid(columnspan=8, padx=5, pady=5)

        for idx, fid in enumerate(feeder_list):
            ttk.Label(frame, text=fid).grid(
                row=1, column=idx, sticky='W', padx=10, pady=5
            )

        var = []

        for feeder, switch_list in fdr_sw_locname.items():
            col = list(fdr_sw_locname).index(feeder)
            for i, switch in enumerate(switch_list):
                var.append(tk.IntVar())
                ttk.Checkbutton(
                    frame, text=switch, variable=var[-1]
                ).grid(row=i + 4, column=col, sticky='W', padx=10, pady=5)

        def select_all():
            for checkbox_var in var:
                checkbox_var.set(1)

        def unselect_all():
            for checkbox_var in var:
                checkbox_var.set(0)

        ttk.Button(
            button_frame, text='Select All', command=select_all
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            button_frame, text='Unselect All', command=unselect_all
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            button_frame,
            text='Okay',
            command=lambda: button_frame.master.destroy()
        ).pack(side=tk.LEFT, padx=5)
        ttk.Button(
            button_frame,
            text='Exit',
            command=lambda: self.exit_script(button_frame.master)
        ).pack(side=tk.LEFT, padx=5)

        return var, fdr_sw_locname

    def run_window(
        self,
        feeder_list: List[str],
        feeders_devices: Dict[str, List[Any]],
        region: str,
        relays_configured: bool
    ) -> Dict[str, List[Any]]:
        """
        Display device selection window and collect user choices.

        Creates a scrollable window with device checkboxes organized
        by feeder and returns the user's selection.

        Args:
            feeder_list: List of feeder names.
            feeders_devices: Dict mapping feeders to device lists.
            region: Network region for display formatting.
            relays_configured: True if relays are pre-configured.

        Returns:
            Dict mapping feeder names to lists of selected device
            objects.
        """
        def _window_dim(feeder_list, feeders_switches, region):
            num_columns = len(feeder_list)
            if region == 'Regional Models':
                column_width = 230
            else:
                column_width = 120
            col_padding = 0
            window_width = max(
                (num_columns * column_width + col_padding), 600
            )

            if window_width > 1300:
                window_width = 1500

            num_rows = max(
                len(lst) for lst in feeders_switches.values()
            )
            row_height = 32
            row_padding = 100
            window_height = max(
                (num_rows * row_height + row_padding), 350
            )

            if window_height > 900:
                window_height = 900

            return window_width, window_height

        root = tk.Tk()
        root.title("Distribution Fault Study")
        window_width, window_height = _window_dim(
            feeder_list, feeders_devices, region
        )
        self.center_window(root, window_width, window_height)

        main_frame = tk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True)

        canvas, frame = self.setup_scrollable_frame(main_frame)

        button_frame = tk.Frame(root)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=5)

        var, fdr_dev_locname = self.populate(
            frame, feeder_list, feeders_devices,
            region, button_frame, relays_configured
        )

        canvas.pack(expand=True, fill=tk.BOTH)
        root.mainloop()

        # Build list of selected devices
        acr_fuse_list = [
            dev for feeder in fdr_dev_locname
            for i, dev in enumerate(feeders_devices[feeder])
            if var[
                sum(
                    len(fdr_dev_locname[f])
                    for f in list(fdr_dev_locname)[
                        :list(fdr_dev_locname).index(feeder)
                    ]
                ) + i
            ].get() == 1
        ]

        feeders_relays = {
            feeder: [
                switch for switch in switches
                if switch in acr_fuse_list
            ]
            for feeder, switches in feeders_devices.items()
        }

        return feeders_relays


def get_input(
    app: pft.Application, region: str, study_selections: List[str]
) -> Tuple[Dict, Dict, Dict, Dict]:
    """
    Main entry point for user input collection.

    Creates a FaultLevelStudy instance and executes the complete
    input collection workflow.

    Args:
        app: PowerFactory application instance.
        region: Network region ('SEQ' or 'Regional Models').
        study_selections: List of selected study type strings.

    Returns:
        Tuple containing:
            - feeders_devices: Dict of feeder names to device lists.
            - bu_devices: Dict of grid objects to backup device lists.
            - user_selection: Dict of user-selected devices per feeder.
            - external_grid: Dict of grid fault level parameters.

    Example:
        >>> feeders, bu, selection, grid = get_input(
        ...     app, 'SEQ', ['Fault Level Study']
        ... )
    """
    study = FaultLevelStudy(app)
    return study.main(region, study_selections)