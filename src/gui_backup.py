"""Tkinter GUI Application for SPU Processing Tool."""

import os
import sys
import subprocess
import platform
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading

from .excel_handler import ExcelHandler
from .processor import SPUProcessor
from .utils import get_input_folder, get_template_folder


class SPUToolGUI:
    """Main GUI Application class."""

    VERSION = "1.0.0"

    def __init__(self, root):
        self.root = root
        self.root.title(f"Process_CDD_ZTE_v{self.VERSION}")
        self.root.geometry("1400x800")
        self.root.minsize(1200, 600)

        # Initialize handlers
        self.excel_handler = ExcelHandler()
        self.processor = SPUProcessor()

        # File paths
        self.input_file_path = None
        self.template_file_path = None

        # Build UI
        self._create_widgets()

    def _create_widgets(self):
        """Create all UI widgets."""
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Top section: Input File
        self._create_input_section(main_frame)

        # Middle section: ZTE SPU Output
        self._create_spu_section(main_frame)

        # Status bar
        self._create_status_bar(main_frame)

        # Bottom section: Data Display (Tabbed)
        self._create_data_display(main_frame)

    def _create_input_section(self, parent):
        """Create the Input File section."""
        input_frame = ttk.LabelFrame(parent, text="Input File", padding="10")
        input_frame.pack(fill=tk.X, pady=(0, 10))

        # Button to select input file
        self.btn_select_input = ttk.Button(
            input_frame,
            text="Select CDD Input File",
            command=self._select_input_file,
            width=80
        )
        self.btn_select_input.pack(pady=5)

        # Label to show selected file
        self.lbl_input_file = ttk.Label(input_frame, text="No file selected", foreground="gray")
        self.lbl_input_file.pack(pady=5)

    def _create_spu_section(self, parent):
        """Create the Template and Output sections."""
        # Template section
        template_frame = ttk.LabelFrame(parent, text="Template", padding="10")
        template_frame.pack(fill=tk.X, pady=(0, 10))

        # Button to select template
        self.btn_select_template = ttk.Button(
            template_frame,
            text="Select SPU Template",
            command=self._select_template_file,
            width=80
        )
        self.btn_select_template.pack(pady=5)

        # Label to show selected template file path
        self.lbl_template_file = ttk.Label(template_frame, text="No template selected", foreground="gray")
        self.lbl_template_file.pack(pady=5)

        # Output section
        output_frame = ttk.LabelFrame(parent, text="Output", padding="10")
        output_frame.pack(fill=tk.X, pady=(0, 10))

        # Button to process
        self.btn_process = ttk.Button(
            output_frame,
            text="Process SPU Output",
            command=self._process_spu_output,
            width=80
        )
        self.btn_process.pack(pady=5)

        # Label to show output file path
        self.lbl_output_file = ttk.Label(output_frame, text="No output generated", foreground="gray")
        self.lbl_output_file.pack(pady=5)

    def _create_status_bar(self, parent):
        """Create the status bar."""
        status_frame = ttk.Frame(parent)
        status_frame.pack(fill=tk.X, pady=(0, 10))

        self.lbl_status = ttk.Label(status_frame, text="Ready", foreground="blue")
        self.lbl_status.pack(side=tk.LEFT)

        # Progress bar
        self.progress = ttk.Progressbar(status_frame, mode='determinate', length=200)
        self.progress.pack(side=tk.RIGHT, padx=10)

    def _create_data_display(self, parent):
        """Create the tabbed data display area."""
        # Label for section
        ttk.Label(parent, text="Integrate Output", font=('Arial', 10, 'bold')).pack(anchor=tk.W)

        # Notebook (tabbed view)
        self.notebook = ttk.Notebook(parent)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=5)

        # Create tabs for each sheet
        self.tabs = {}
        self.treeviews = {}

        sheet_names = [
            "IP", "Radio 2G", "Radio 3G", "Radio 4G", "Radio 5G",
            "2G-2G", "2G-3G", "2G-4G", "3G-2G", "3G-3G", "3G-4G",
            "RET", "Mapping"
        ]

        for sheet_name in sheet_names:
            tab = ttk.Frame(self.notebook)
            self.notebook.add(tab, text=sheet_name)
            self.tabs[sheet_name] = tab

            # Create Treeview with scrollbars
            tree_frame = ttk.Frame(tab)
            tree_frame.pack(fill=tk.BOTH, expand=True)

            # Scrollbars
            vsb = ttk.Scrollbar(tree_frame, orient="vertical")
            hsb = ttk.Scrollbar(tree_frame, orient="horizontal")

            # Treeview
            tree = ttk.Treeview(
                tree_frame,
                yscrollcommand=vsb.set,
                xscrollcommand=hsb.set,
                show="headings"
            )

            vsb.config(command=tree.yview)
            hsb.config(command=tree.xview)

            # Pack scrollbars and treeview
            vsb.pack(side=tk.RIGHT, fill=tk.Y)
            hsb.pack(side=tk.BOTTOM, fill=tk.X)
            tree.pack(fill=tk.BOTH, expand=True)

            self.treeviews[sheet_name] = tree

    def _select_input_file(self):
        """Handle input file selection."""
        initial_dir = get_input_folder()
        if not os.path.exists(initial_dir):
            initial_dir = os.path.expanduser("~")

        file_path = filedialog.askopenfilename(
            title="Select CDD Input File",
            initialdir=initial_dir,
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )

        if file_path:
            self.input_file_path = file_path
            self.lbl_input_file.config(text=file_path, foreground="black")
            self._update_status(f"Loading: {os.path.basename(file_path)}")

            # Load the file in a separate thread
            thread = threading.Thread(target=self._load_input_file)
            thread.daemon = True
            thread.start()

    def _load_input_file(self):
        """Load input file in background thread."""
        try:
            # Read input file
            data = self.excel_handler.read_input_file(self.input_file_path)

            # Update processor
            self.processor.set_input_data(data)

            # Update UI in main thread
            self.root.after(0, self._populate_data_display, data)
            self.root.after(0, self._update_status, "File loaded successfully")

        except Exception as e:
            self.root.after(0, self._show_error, f"Failed to load file: {e}")

    def _populate_data_display(self, data):
        """Populate the data display tabs with loaded data."""
        for sheet_name, df in data.items():
            if sheet_name in self.treeviews:
                tree = self.treeviews[sheet_name]

                # Clear existing data
                tree.delete(*tree.get_children())

                if not df.empty:
                    # Set columns
                    columns = list(df.columns)
                    tree["columns"] = columns

                    # Configure column headers
                    for col in columns:
                        tree.heading(col, text=col)
                        tree.column(col, width=100, minwidth=50)

                    # Add rows (limit to first 1000 for performance)
                    for idx, row in df.head(1000).iterrows():
                        values = [str(v) if v is not None else "" for v in row.tolist()]
                        tree.insert("", tk.END, values=values)

    def _select_template_file(self):
        """Handle template file selection."""
        initial_dir = get_template_folder()
        if not os.path.exists(initial_dir):
            initial_dir = os.path.expanduser("~")

        file_path = filedialog.askopenfilename(
            title="Select SPU Template File",
            initialdir=initial_dir,
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )

        if file_path:
            self.template_file_path = file_path
            self.lbl_template_file.config(text=file_path, foreground="black")
            self._update_status(f"Template selected: {os.path.basename(file_path)}")

            try:
                self.processor.set_template(file_path)
            except Exception as e:
                self._show_error(f"Failed to load template: {e}")

    def _process_spu_output(self):
        """Handle SPU output processing."""
        # Validate inputs
        if not self.input_file_path:
            self._show_error("Please select a CDD input file first")
            return

        if not self.template_file_path:
            self._show_error("Please select an SPU template file first")
            return

        # Disable button during processing
        self.btn_process.config(state=tk.DISABLED)
        self._update_status("Processing...")
        self.progress["value"] = 0

        # Process in background thread
        thread = threading.Thread(target=self._run_processing)
        thread.daemon = True
        thread.start()

    def _run_processing(self):
        """Run processing in background thread."""
        try:
            def progress_callback(message, percentage):
                self.root.after(0, self._update_progress, message, percentage)

            output_files = self.processor.process(progress_callback)

            # Update Output label with output file path
            if output_files:
                self.root.after(0, self._update_output_label, output_files[0])

            # Show success message
            files_list = "\n".join(output_files)
            self.root.after(0, self._show_success, f"Output files created:\n{files_list}")

            # Automatically open output files
            for output_file in output_files:
                self.root.after(0, self._open_file, output_file)

        except Exception as e:
            self.root.after(0, self._show_error, f"Processing failed: {e}")

        finally:
            self.root.after(0, self._reset_processing_state)

    def _update_output_label(self, file_path):
        """Update the output file label."""
        self.lbl_output_file.config(text=file_path, foreground="black")

    def _update_progress(self, message, percentage):
        """Update progress bar and status."""
        self.lbl_status.config(text=message)
        self.progress["value"] = percentage

    def _reset_processing_state(self):
        """Reset UI after processing."""
        self.btn_process.config(state=tk.NORMAL)

    def _update_status(self, message):
        """Update status label."""
        self.lbl_status.config(text=message)

    def _show_error(self, message):
        """Show error message."""
        self._update_status("Error")
        messagebox.showerror("Error", message)

    def _show_success(self, message):
        """Show success message."""
        self._update_status("Complete")
        self.progress["value"] = 100
        messagebox.showinfo("Success", message)

    def _open_file(self, file_path):
        """Open file with default application based on OS."""
        try:
            if platform.system() == "Windows":
                os.startfile(file_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", file_path], check=True)
            else:  # Linux
                subprocess.run(["xdg-open", file_path], check=True)
        except Exception as e:
            self._update_status(f"Could not open file: {e}")


def run_app():
    """Run the GUI application."""
    root = tk.Tk()
    app = SPUToolGUI(root)
    root.mainloop()
