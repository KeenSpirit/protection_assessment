import tkinter as tk
from tkinter import ttk, messagebox
import sys


class TransformerSpecificationGUI:
    def __init__(self, items_list):
        self.items_list = items_list
        self.user_inputs = {}
        self.root = tk.Tk()
        self.setup_window()
        self.create_widgets()

    def setup_window(self):
        """Configure the main window"""
        self.root.title("Transformer Fuse Specification")

        # Calculate window size based on number of items (with reasonable limits)
        base_height = 200
        item_height = 80  # Height per item
        max_display_items = 8  # Maximum items before scrolling

        display_items = min(len(self.items_list), max_display_items)
        window_height = base_height + (display_items * item_height)
        window_width = 600

        # Center the window on screen
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.resizable(True, True)

    def create_widgets(self):
        """Create all GUI widgets"""
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)

        # Header text
        header_label = ttk.Label(
            main_frame,
            text="Please enter the requested data for the following downstream distribution transformers:",
            wraplength=550,
            justify=tk.CENTER
        )
        header_label.grid(row=0, column=0, pady=(0, 20), sticky=(tk.W, tk.E))

        # Scrollable frame for user inputs
        self.create_scrollable_frame(main_frame)

        # Bottom frame (always visible)
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.grid(row=2, column=0, pady=(20, 0), sticky=(tk.W, tk.E))
        bottom_frame.columnconfigure(1, weight=1)

        # Reference text
        ref_label = ttk.Label(
            bottom_frame,
            text="Refer to Technical Instruction TS0013i: RMU Fuse Selection Guide for further information.",
            font=("TkDefaultFont", 9)
        )
        ref_label.grid(row=0, column=0, columnspan=3, pady=(0, 10))

        # Buttons
        ok_button = ttk.Button(
            bottom_frame,
            text="Okay",
            command=self.validate_and_close
        )
        ok_button.grid(row=1, column=0, padx=(0, 10), sticky=tk.W)

        exit_button = ttk.Button(
            bottom_frame,
            text="Exit",
            command=self.exit_application
        )
        exit_button.grid(row=1, column=2, sticky=tk.E)

    def create_scrollable_frame(self, parent):
        """Create scrollable frame for user inputs"""
        # Create canvas and scrollbar
        self.canvas = tk.Canvas(parent, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(parent, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        # Store references for scrollbar management
        self.canvas_parent = parent

        # Configure scrolling
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.update_scroll_region()
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Grid the canvas (scrollbar will be added conditionally)
        self.canvas.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure canvas to expand
        parent.rowconfigure(1, weight=1)
        parent.columnconfigure(0, weight=1)

        # Mouse wheel binding for scrolling
        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        self.canvas.bind("<MouseWheel>", _on_mousewheel)  # Windows
        self.canvas.bind("<Button-4>", lambda e: self.canvas.yview_scroll(-1, "units"))  # Linux
        self.canvas.bind("<Button-5>", lambda e: self.canvas.yview_scroll(1, "units"))  # Linux

        # Create input widgets for each item
        self.create_input_widgets(self.scrollable_frame)

        # Schedule scrollbar check after widgets are created
        self.root.after(100, self.check_scrollbar_needed)

    def update_scroll_region(self):
        """Update the scroll region and check if scrollbar is needed"""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.root.after_idle(self.check_scrollbar_needed)

    def check_scrollbar_needed(self):
        """Check if scrollbar is needed and show/hide accordingly"""
        self.canvas.update_idletasks()

        # Get the actual content height and canvas height
        bbox = self.canvas.bbox("all")
        if bbox:
            content_height = bbox[3] - bbox[1]
            canvas_height = self.canvas.winfo_height()

            if content_height > canvas_height:
                # Show scrollbar
                self.scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
                self.canvas_parent.columnconfigure(1, weight=0)
            else:
                # Hide scrollbar
                self.scrollbar.grid_remove()
                self.canvas_parent.columnconfigure(1, weight=0)

    def create_input_widgets(self, parent):
        """Create input widgets for each item"""
        self.insulation_vars = {}
        self.impedance_vars = {}
        self.impedance_frames = {}

        for i, item in enumerate(self.items_list):
            # Item frame
            item_frame = ttk.LabelFrame(parent, text=f"Item: {item[:-2]}", padding="10")
            item_frame.grid(row=i, column=0, pady=5, padx=5, sticky=(tk.W, tk.E))
            parent.columnconfigure(0, weight=1)

            # Insulation type selection
            insulation_frame = ttk.Frame(item_frame)
            insulation_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

            ttk.Label(insulation_frame, text="Insulation Type:").grid(row=0, column=0, sticky=tk.W)

            self.insulation_vars[item] = tk.StringVar()
            air_radio = ttk.Radiobutton(
                insulation_frame,
                text="Air Insulated",
                variable=self.insulation_vars[item],
                value="air",
                command=lambda i=item: self.on_insulation_change(i)
            )
            air_radio.grid(row=1, column=0, sticky=tk.W, padx=(20, 0))

            oil_radio = ttk.Radiobutton(
                insulation_frame,
                text="Oil Insulated",
                variable=self.insulation_vars[item],
                value="oil",
                command=lambda i=item: self.on_insulation_change(i)
            )
            oil_radio.grid(row=1, column=1, sticky=tk.W, padx=(20, 0))

            # Impedance selection frame (initially hidden)
            impedance_frame = ttk.Frame(item_frame)
            impedance_frame.grid(row=1, column=0, sticky=(tk.W, tk.E))
            self.impedance_frames[item] = impedance_frame

            ttk.Label(impedance_frame, text="Impedance Type:").grid(row=0, column=0, sticky=tk.W)

            self.impedance_vars[item] = tk.StringVar()
            high_radio = ttk.Radiobutton(
                impedance_frame,
                text="High Impedance",
                variable=self.impedance_vars[item],
                value="high"
            )
            high_radio.grid(row=1, column=0, sticky=tk.W, padx=(20, 0))

            low_radio = ttk.Radiobutton(
                impedance_frame,
                text="Low Impedance",
                variable=self.impedance_vars[item],
                value="low"
            )
            low_radio.grid(row=1, column=1, sticky=tk.W, padx=(20, 0))

            # Initially hide impedance frame
            self.toggle_impedance_frame(item, False)

    def on_insulation_change(self, item):
        """Handle insulation type change"""
        insulation_type = self.insulation_vars[item].get()
        if insulation_type == "air":
            try:
                last_char = item[-1]
                last_digit = int(last_char)
                show_impedance = last_digit > 8
            except (ValueError, IndexError):
                show_impedance = False

            self.toggle_impedance_frame(item, show_impedance)

            if not show_impedance:
                self.impedance_vars[item].set("")
        else:
            self.toggle_impedance_frame(item, False)
            self.impedance_vars[item].set("")  # Clear impedance selection

        # Check scrollbar after layout change
        self.root.after_idle(self.check_scrollbar_needed)

    def toggle_impedance_frame(self, item, show):
        """Show or hide impedance selection frame"""
        if show:
            self.impedance_frames[item].grid()
        else:
            self.impedance_frames[item].grid_remove()

    def validate_inputs(self):
        """Validate that all required inputs are provided"""
        missing_items = []

        for item in self.items_list:
            insulation_type = self.insulation_vars[item].get()

            if not insulation_type:
                missing_items.append(f"{item}: No insulation type selected")
            elif insulation_type == "air":
                # Check if impedance selection is required (last character = 1)
                try:
                    last_char = item[-1]
                    last_digit = int(last_char)
                    impedance_required = last_digit == 1
                except (ValueError, IndexError):
                    # If last character is not a valid integer, impedance not required
                    impedance_required = False

                if impedance_required:
                    impedance_type = self.impedance_vars[item].get()
                    if not impedance_type:
                        missing_items.append(f"{item}: No impedance type selected")

        return missing_items

    def collect_inputs(self):
        """Collect all user inputs into a dictionary"""
        inputs = {}

        for item in self.items_list:
            insulation_type = self.insulation_vars[item].get()
            inputs[item] = {"insulation": insulation_type}

            if insulation_type == "air":
                impedance_type = self.impedance_vars[item].get()
                inputs[item]["impedance"] = impedance_type
            else:
                inputs[item]["impedance"] = None

        return inputs

    def validate_and_close(self):
        """Validate inputs and close if valid"""
        missing_items = self.validate_inputs()

        if missing_items:
            error_message = "Please provide the following missing information:\n\n"
            error_message += "\n".join(missing_items)
            messagebox.showerror("Missing Information", error_message)
        else:
            self.user_inputs = self.collect_inputs()
            self.root.quit()

    def exit_application(self):
        """Exit the application"""
        self.root.destroy()
        sys.exit(0)
        #self.root.quit()
        #sys.exit()

    def run(self):
        """Run the GUI and return user inputs"""
        self.root.mainloop()
        self.root.destroy()
        return self.user_inputs


def get_transformer_specifications(items_list):
    """
    Main function to get transformer specifications from user

    Args:
        items_list: List of item names (strings)

    Returns:
        Dictionary with user inputs for each item
    """
    if not items_list:
        print("No items provided")
        return {}

    gui = TransformerSpecificationGUI(items_list)
    return gui.run()


