"""Tkinter GUI Application for SPU Processing Tool - Modern Design."""

import os
import sys
import subprocess
import platform
import tkinter as tk
from pathlib import Path
from tkinter import ttk, filedialog, messagebox
import threading

from .excel_handler import ExcelHandler
from .processor import SPUProcessor
from .utils import get_input_folder, get_template_folder
from CDD.excel_to_json import convert as calloff_convert, default_output_path as calloff_default_output
from CDD.fill_mimo import fill_workbook as cdd_fill_workbook


class ModernStyle:
    """Modern color scheme and styling."""

    # Colors
    BG_PRIMARY = "#f5f6fa"
    BG_SECONDARY = "#ffffff"
    BG_ACCENT = "#3498db"
    BG_SUCCESS = "#27ae60"
    BG_WARNING = "#f39c12"

    TEXT_PRIMARY = "#2c3e50"
    TEXT_SECONDARY = "#7f8c8d"
    TEXT_LIGHT = "#ffffff"

    BORDER_COLOR = "#dcdde1"

    # Button colors
    BTN_PRIMARY = "#3498db"
    BTN_PRIMARY_HOVER = "#2980b9"
    BTN_SUCCESS = "#27ae60"
    BTN_SUCCESS_HOVER = "#219a52"


class SPUToolGUI:
    """Main GUI Application class with modern design."""

    VERSION = "2.0.0"

    def __init__(self, root):
        self.root = root
        self.root.title(f"SPU Processing Tool v{self.VERSION}")
        self.root.geometry("1400x850")
        self.root.minsize(1200, 700)

        # Set custom window icon
        self._set_window_icon()

        # Set window background
        self.root.configure(bg=ModernStyle.BG_PRIMARY)

    def _set_window_icon(self):
        """Set custom window icon."""
        try:
            # Get base path (handles PyInstaller frozen exe)
            if getattr(sys, 'frozen', False):
                base_path = os.path.dirname(sys.executable)
            else:
                base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

            # Try to load icon from various locations
            icon_paths = [
                os.path.join(base_path, "icon.png"),
                os.path.join(os.path.dirname(os.path.dirname(__file__)), "icon.png"),
                os.path.join(os.path.dirname(__file__), "icon.png"),
                "icon.png"
            ]

            for icon_path in icon_paths:
                if os.path.exists(icon_path):
                    icon = tk.PhotoImage(file=icon_path)
                    self.root.iconphoto(True, icon)
                    self._icon = icon  # Keep reference to prevent garbage collection
                    break
        except Exception:
            pass  # Use default icon if custom icon fails

        # Configure modern style
        self._configure_style()

        # Initialize handlers
        self.excel_handler = ExcelHandler()
        self.processor = SPUProcessor()

        # File paths
        self.input_file_path = None
        self.template_file_path = None
        self.calloff_file_path = None
        self.calloff_json_path = None

        # Build UI
        self._create_widgets()

    def _configure_style(self):
        """Configure ttk styles for modern look."""
        self.style = ttk.Style()

        # Try to use a modern theme
        available_themes = self.style.theme_names()
        if 'clam' in available_themes:
            self.style.theme_use('clam')

        # Configure frame styles
        self.style.configure(
            "Card.TFrame",
            background=ModernStyle.BG_SECONDARY,
            relief="flat"
        )

        self.style.configure(
            "Main.TFrame",
            background=ModernStyle.BG_PRIMARY
        )

        # Configure label styles
        self.style.configure(
            "Title.TLabel",
            background=ModernStyle.BG_SECONDARY,
            foreground=ModernStyle.TEXT_PRIMARY,
            font=("Segoe UI", 11, "bold"),
            padding=(10, 5)
        )

        self.style.configure(
            "Path.TLabel",
            background=ModernStyle.BG_SECONDARY,
            foreground=ModernStyle.TEXT_SECONDARY,
            font=("Segoe UI", 9),
            padding=(10, 5)
        )

        self.style.configure(
            "Status.TLabel",
            background=ModernStyle.BG_PRIMARY,
            foreground=ModernStyle.BG_ACCENT,
            font=("Segoe UI", 10, "bold")
        )

        self.style.configure(
            "Section.TLabel",
            background=ModernStyle.BG_PRIMARY,
            foreground=ModernStyle.TEXT_PRIMARY,
            font=("Segoe UI", 12, "bold")
        )

        # Configure button styles
        self.style.configure(
            "Primary.TButton",
            font=("Segoe UI", 10),
            padding=(20, 10)
        )

        self.style.configure(
            "Action.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(20, 12)
        )

        # Configure labelframe
        self.style.configure(
            "Card.TLabelframe",
            background=ModernStyle.BG_SECONDARY,
            relief="solid",
            borderwidth=1
        )

        self.style.configure(
            "Card.TLabelframe.Label",
            background=ModernStyle.BG_SECONDARY,
            foreground=ModernStyle.TEXT_PRIMARY,
            font=("Segoe UI", 11, "bold"),
            padding=(5, 2)
        )

        # Configure notebook (tabs)
        self.style.configure(
            "TNotebook",
            background=ModernStyle.BG_PRIMARY,
            tabmargins=[5, 5, 5, 0]
        )

        self.style.configure(
            "TNotebook.Tab",
            background=ModernStyle.BG_SECONDARY,
            foreground=ModernStyle.TEXT_PRIMARY,
            padding=[15, 8],
            font=("Segoe UI", 9)
        )

        self.style.map(
            "TNotebook.Tab",
            background=[("selected", ModernStyle.BG_ACCENT)],
            foreground=[("selected", ModernStyle.TEXT_LIGHT)]
        )

        # Configure progress bar
        self.style.configure(
            "Custom.Horizontal.TProgressbar",
            troughcolor=ModernStyle.BORDER_COLOR,
            background=ModernStyle.BG_SUCCESS,
            thickness=8
        )

        # Configure treeview
        self.style.configure(
            "Custom.Treeview",
            background=ModernStyle.BG_SECONDARY,
            foreground=ModernStyle.TEXT_PRIMARY,
            fieldbackground=ModernStyle.BG_SECONDARY,
            font=("Segoe UI", 9),
            rowheight=28
        )

        self.style.configure(
            "Custom.Treeview.Heading",
            background=ModernStyle.BG_ACCENT,
            foreground=ModernStyle.TEXT_LIGHT,
            font=("Segoe UI", 9, "bold"),
            padding=(5, 8)
        )

        self.style.map(
            "Custom.Treeview",
            background=[("selected", ModernStyle.BG_ACCENT)],
            foreground=[("selected", ModernStyle.TEXT_LIGHT)]
        )

    def _create_widgets(self):
        """Create all UI widgets."""
        # Main container with padding
        main_frame = ttk.Frame(self.root, style="Main.TFrame", padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header
        self._create_header(main_frame)

        # Top section container (Input, Template, Output side by side)
        top_container = ttk.Frame(main_frame, style="Main.TFrame")
        top_container.pack(fill=tk.X, pady=(0, 15))

        # Configure grid columns for equal width
        top_container.columnconfigure(0, weight=1)
        top_container.columnconfigure(1, weight=1)
        top_container.columnconfigure(2, weight=1)
        top_container.columnconfigure(3, weight=1)
        top_container.columnconfigure(4, weight=1)

        # Call-off File Card (Step 1: Excel -> JSON)
        self._create_calloff_card(top_container, 0)

        # CDD Verification Card
        self._create_verify_card(top_container, 1)

        # Input File Card
        self._create_input_card(top_container, 2)

        # Template Card
        self._create_template_card(top_container, 3)

        # Output Card
        self._create_output_card(top_container, 4)

        # Status bar
        self._create_status_bar(main_frame)

        # Bottom section: Data Display (Tabbed)
        self._create_data_display(main_frame)

    def _create_header(self, parent):
        """Create application header."""
        header_frame = ttk.Frame(parent, style="Main.TFrame")
        header_frame.pack(fill=tk.X, pady=(0, 15))

        # Title
        title_label = tk.Label(
            header_frame,
            text="SPU Processing Tool",
            font=("Segoe UI", 18, "bold"),
            fg=ModernStyle.TEXT_PRIMARY,
            bg=ModernStyle.BG_PRIMARY
        )
        title_label.pack(side=tk.LEFT)

        # Version
        version_label = tk.Label(
            header_frame,
            text=f"v{self.VERSION}",
            font=("Segoe UI", 10),
            fg=ModernStyle.TEXT_SECONDARY,
            bg=ModernStyle.BG_PRIMARY
        )
        version_label.pack(side=tk.LEFT, padx=(10, 0), pady=(8, 0))

    def _create_card_frame(self, parent, title, column):
        """Create a card-style frame."""
        # Outer frame with border effect
        outer_frame = tk.Frame(
            parent,
            bg=ModernStyle.BORDER_COLOR,
            padx=1,
            pady=1
        )
        outer_frame.grid(row=0, column=column, sticky="nsew", padx=5)

        # Inner frame (white card)
        card = tk.Frame(
            outer_frame,
            bg=ModernStyle.BG_SECONDARY,
            padx=15,
            pady=15
        )
        card.pack(fill=tk.BOTH, expand=True)

        # Title
        title_label = tk.Label(
            card,
            text=title,
            font=("Segoe UI", 11, "bold"),
            fg=ModernStyle.TEXT_PRIMARY,
            bg=ModernStyle.BG_SECONDARY,
            anchor="w"
        )
        title_label.pack(fill=tk.X, pady=(0, 10))

        # Separator line
        separator = tk.Frame(card, height=2, bg=ModernStyle.BG_ACCENT)
        separator.pack(fill=tk.X, pady=(0, 15))

        return card

    def _create_calloff_card(self, parent, column):
        """Create the Call-off File card (Excel -> JSON)."""
        card = self._create_card_frame(parent, "Call-off File", column)

        self.btn_select_calloff = tk.Button(
            card,
            text="Select Call-off File",
            command=self._select_calloff_file,
            font=("Segoe UI", 10),
            bg=ModernStyle.BTN_PRIMARY,
            fg=ModernStyle.TEXT_LIGHT,
            activebackground=ModernStyle.BTN_PRIMARY_HOVER,
            activeforeground=ModernStyle.TEXT_LIGHT,
            relief="flat",
            cursor="hand2",
            padx=20,
            pady=10
        )
        self.btn_select_calloff.pack(fill=tk.X, pady=(0, 8))

        self.btn_autofill_cdd = tk.Button(
            card,
            text="CDD to autofill",
            command=self._select_cdd_to_autofill,
            font=("Segoe UI", 10, "bold"),
            bg=ModernStyle.BTN_SUCCESS,
            fg=ModernStyle.TEXT_LIGHT,
            activebackground=ModernStyle.BTN_SUCCESS_HOVER,
            activeforeground=ModernStyle.TEXT_LIGHT,
            relief="flat",
            cursor="hand2",
            padx=20,
            pady=10
        )
        self.btn_autofill_cdd.pack(fill=tk.X, pady=(0, 10))

        self.lbl_calloff_file = tk.Label(
            card,
            text="No Call-off selected",
            font=("Segoe UI", 9),
            fg=ModernStyle.TEXT_SECONDARY,
            bg=ModernStyle.BG_SECONDARY,
            wraplength=350,
            justify="left",
            anchor="w"
        )
        self.lbl_calloff_file.pack(fill=tk.X)

    def _create_input_card(self, parent, column):
        """Create the Input File card."""
        card = self._create_card_frame(parent, "Input File", column)

        # Button
        self.btn_select_input = tk.Button(
            card,
            text="Select CDD Input File",
            command=self._select_input_file,
            font=("Segoe UI", 10),
            bg=ModernStyle.BTN_PRIMARY,
            fg=ModernStyle.TEXT_LIGHT,
            activebackground=ModernStyle.BTN_PRIMARY_HOVER,
            activeforeground=ModernStyle.TEXT_LIGHT,
            relief="flat",
            cursor="hand2",
            padx=20,
            pady=10
        )
        self.btn_select_input.pack(fill=tk.X, pady=(0, 10))

        # File path label
        self.lbl_input_file = tk.Label(
            card,
            text="No file selected",
            font=("Segoe UI", 9),
            fg=ModernStyle.TEXT_SECONDARY,
            bg=ModernStyle.BG_SECONDARY,
            wraplength=350,
            justify="left",
            anchor="w"
        )
        self.lbl_input_file.pack(fill=tk.X)

    def _create_template_card(self, parent, column):
        """Create the Template card."""
        card = self._create_card_frame(parent, "Template", column)

        # Button
        self.btn_select_template = tk.Button(
            card,
            text="Select SPU Template",
            command=self._select_template_file,
            font=("Segoe UI", 10),
            bg=ModernStyle.BTN_PRIMARY,
            fg=ModernStyle.TEXT_LIGHT,
            activebackground=ModernStyle.BTN_PRIMARY_HOVER,
            activeforeground=ModernStyle.TEXT_LIGHT,
            relief="flat",
            cursor="hand2",
            padx=20,
            pady=10
        )
        self.btn_select_template.pack(fill=tk.X, pady=(0, 10))

        # File path label
        self.lbl_template_file = tk.Label(
            card,
            text="No template selected",
            font=("Segoe UI", 9),
            fg=ModernStyle.TEXT_SECONDARY,
            bg=ModernStyle.BG_SECONDARY,
            wraplength=350,
            justify="left",
            anchor="w"
        )
        self.lbl_template_file.pack(fill=tk.X)

    def _create_output_card(self, parent, column):
        """Create the Output card."""
        card = self._create_card_frame(parent, "Output", column)

        # Button (Green for action)
        self.btn_process = tk.Button(
            card,
            text="Process SPU Output",
            command=self._process_spu_output,
            font=("Segoe UI", 10, "bold"),
            bg=ModernStyle.BTN_SUCCESS,
            fg=ModernStyle.TEXT_LIGHT,
            activebackground=ModernStyle.BTN_SUCCESS_HOVER,
            activeforeground=ModernStyle.TEXT_LIGHT,
            relief="flat",
            cursor="hand2",
            padx=20,
            pady=10
        )
        self.btn_process.pack(fill=tk.X, pady=(0, 10))

        # File path label
        self.lbl_output_file = tk.Label(
            card,
            text="No output generated",
            font=("Segoe UI", 9),
            fg=ModernStyle.TEXT_SECONDARY,
            bg=ModernStyle.BG_SECONDARY,
            wraplength=350,
            justify="left",
            anchor="w"
        )
        self.lbl_output_file.pack(fill=tk.X)

    def _create_verify_card(self, parent, column):
        """Create the CDD Verification card."""
        card = self._create_card_frame(parent, "CDD Verification", column)

        # Verify button
        self.btn_verify_cdd = tk.Button(
            card,
            text="Verify CDD",
            command=self._verify_cdd,
            font=("Segoe UI", 10),
            bg=ModernStyle.BTN_PRIMARY,
            fg=ModernStyle.TEXT_LIGHT,
            activebackground=ModernStyle.BTN_PRIMARY_HOVER,
            activeforeground=ModernStyle.TEXT_LIGHT,
            relief="flat",
            cursor="hand2",
            padx=20,
            pady=10
        )
        self.btn_verify_cdd.pack(fill=tk.X, pady=(0, 8))

        # Auto-correct button
        self.btn_autocorrect_cdd = tk.Button(
            card,
            text="Auto-correct",
            command=self._autocorrect_cdd,
            font=("Segoe UI", 10, "bold"),
            bg=ModernStyle.BG_WARNING,
            fg=ModernStyle.TEXT_LIGHT,
            activebackground="#d68910",
            activeforeground=ModernStyle.TEXT_LIGHT,
            relief="flat",
            cursor="hand2",
            padx=20,
            pady=10
        )
        self.btn_autocorrect_cdd.pack(fill=tk.X, pady=(0, 10))

        # Result label
        self.lbl_verify_result = tk.Label(
            card,
            text="No verification run yet",
            font=("Segoe UI", 9),
            fg=ModernStyle.TEXT_SECONDARY,
            bg=ModernStyle.BG_SECONDARY,
            wraplength=350,
            justify="left",
            anchor="w"
        )
        self.lbl_verify_result.pack(fill=tk.X)

    def _create_status_bar(self, parent):
        """Create the status bar."""
        status_frame = tk.Frame(parent, bg=ModernStyle.BG_PRIMARY)
        status_frame.pack(fill=tk.X, pady=(0, 10))

        self.lbl_status = tk.Label(
            status_frame,
            text="Ready",
            font=("Segoe UI", 10, "bold"),
            fg=ModernStyle.BG_ACCENT,
            bg=ModernStyle.BG_PRIMARY
        )
        self.lbl_status.pack(side=tk.LEFT)

        # Progress bar
        self.progress = ttk.Progressbar(
            status_frame,
            style="Custom.Horizontal.TProgressbar",
            mode='determinate',
            length=300
        )
        self.progress.pack(side=tk.RIGHT, padx=10)

    def _create_data_display(self, parent):
        """Create the tabbed data display area."""
        # Section label
        section_label = tk.Label(
            parent,
            text="Data Preview",
            font=("Segoe UI", 12, "bold"),
            fg=ModernStyle.TEXT_PRIMARY,
            bg=ModernStyle.BG_PRIMARY
        )
        section_label.pack(anchor=tk.W, pady=(0, 10))

        # Notebook (tabbed view)
        self.notebook = ttk.Notebook(parent)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Create tabs for each sheet
        self.tabs = {}
        self.treeviews = {}

        sheet_names = [
            "IP", "Radio 2G", "Radio 3G", "Radio 4G", "Radio 5G",
            "2G-2G", "2G-3G", "2G-4G", "3G-2G", "3G-3G", "3G-4G",
            "RET", "Mapping"
        ]

        for sheet_name in sheet_names:
            tab = ttk.Frame(self.notebook, style="Card.TFrame")
            self.notebook.add(tab, text=f"  {sheet_name}  ")
            self.tabs[sheet_name] = tab

            # Create Treeview with scrollbars
            tree_frame = ttk.Frame(tab)
            tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

            # Scrollbars
            vsb = ttk.Scrollbar(tree_frame, orient="vertical")
            hsb = ttk.Scrollbar(tree_frame, orient="horizontal")

            # Treeview with custom style
            tree = ttk.Treeview(
                tree_frame,
                style="Custom.Treeview",
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

    def _select_calloff_file(self):
        """Pick a Call-off xlsx and convert it to JSON."""
        initial_dir = get_input_folder()
        if not os.path.exists(initial_dir):
            initial_dir = os.path.expanduser("~")

        file_path = filedialog.askopenfilename(
            title="Select Call-off File",
            initialdir=initial_dir,
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not file_path:
            return

        self.calloff_file_path = file_path
        self.lbl_calloff_file.config(
            text=f"{os.path.basename(file_path)}\nConverting to JSON...",
            fg=ModernStyle.TEXT_PRIMARY,
        )
        self._update_status(f"Converting Call-off: {os.path.basename(file_path)}")

        thread = threading.Thread(target=self._run_calloff_conversion, args=(file_path,))
        thread.daemon = True
        thread.start()

    def _run_calloff_conversion(self, file_path):
        """Background worker: convert Call-off xlsx to JSON."""
        try:
            xlsx_path = Path(file_path)
            data = calloff_convert(xlsx_path)
            out_path = calloff_default_output(xlsx_path)
            out_path.parent.mkdir(parents=True, exist_ok=True)
            import json
            out_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
            self.calloff_json_path = str(out_path)
            self.root.after(
                0,
                self._on_calloff_done,
                f"{os.path.basename(file_path)}\nJSON: {out_path}",
                f"Call-off converted: {len(data)} sites",
            )
        except Exception as exc:
            self.root.after(0, self._show_error, f"Call-off conversion failed: {exc}")

    def _on_calloff_done(self, label_text, status_text):
        self.lbl_calloff_file.config(text=label_text, fg=ModernStyle.TEXT_PRIMARY)
        self._update_status(status_text)

    def _select_cdd_to_autofill(self):
        """Pick the empty CDD xlsx and run auto-fill against the Call-off JSON."""
        if not self.calloff_json_path:
            self._show_error("Select a Call-off file first to produce the JSON.")
            return

        initial_dir = get_input_folder()
        if not os.path.exists(initial_dir):
            initial_dir = os.path.expanduser("~")

        file_path = filedialog.askopenfilename(
            title="Select CDD to Auto-Fill",
            initialdir=initial_dir,
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not file_path:
            return

        self._update_status(f"Auto-filling: {os.path.basename(file_path)}")
        thread = threading.Thread(target=self._run_cdd_autofill, args=(file_path,))
        thread.daemon = True
        thread.start()

    def _run_cdd_autofill(self, file_path):
        try:
            counts, counts_5g, skipped, out_path = cdd_fill_workbook(
                Path(file_path), Path(self.calloff_json_path)
            )
            total_4g = sum(counts.values())
            total_5g = sum(counts_5g.values())
            self.input_file_path = str(out_path)
            self.root.after(0, self._on_autofill_done, str(out_path), total_4g, total_5g, len(skipped))
        except Exception as exc:
            self.root.after(0, self._show_error, f"Auto-fill failed: {exc}")

    def _on_autofill_done(self, out_path, total_4g, total_5g, skipped_count):
        self.lbl_input_file.config(text=out_path, fg=ModernStyle.TEXT_PRIMARY)
        self._update_status(
            f"Auto-fill done. 4G:{total_4g} 5G:{total_5g} skipped:{skipped_count}. Loaded as Input File."
        )
        thread = threading.Thread(target=self._load_input_file)
        thread.daemon = True
        thread.start()

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
            self.lbl_input_file.config(text=file_path, fg=ModernStyle.TEXT_PRIMARY)
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
            self.lbl_template_file.config(text=file_path, fg=ModernStyle.TEXT_PRIMARY)
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
        self.btn_process.config(state=tk.DISABLED, bg=ModernStyle.TEXT_SECONDARY)
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
        self.lbl_output_file.config(text=file_path, fg=ModernStyle.TEXT_PRIMARY)

    def _update_progress(self, message, percentage):
        """Update progress bar and status."""
        self.lbl_status.config(text=message)
        self.progress["value"] = percentage

    def _reset_processing_state(self):
        """Reset UI after processing."""
        self.btn_process.config(state=tk.NORMAL, bg=ModernStyle.BTN_SUCCESS)

    def _update_status(self, message):
        """Update status label."""
        self.lbl_status.config(text=message)

    def _show_error(self, message):
        """Show error message."""
        self._update_status("Error")
        self.lbl_status.config(fg="#e74c3c")
        messagebox.showerror("Error", message)
        self.lbl_status.config(fg=ModernStyle.BG_ACCENT)

    def _show_success(self, message):
        """Show success message."""
        self._update_status("Complete")
        self.lbl_status.config(fg=ModernStyle.BG_SUCCESS)
        self.progress["value"] = 100
        messagebox.showinfo("Success", message)

    def _verify_cdd(self):
        """Run CDD verification on the selected input file."""
        if not self.input_file_path:
            self._show_error("Please select a CDD input file first")
            return

        self.btn_verify_cdd.config(state=tk.DISABLED, bg=ModernStyle.TEXT_SECONDARY)
        self.btn_autocorrect_cdd.config(state=tk.DISABLED, bg=ModernStyle.TEXT_SECONDARY)
        self._update_status("Verifying CDD...")
        self.lbl_verify_result.config(text="Running verification...", fg=ModernStyle.TEXT_SECONDARY)

        thread = threading.Thread(target=self._run_verify_cdd)
        thread.daemon = True
        thread.start()

    def _run_verify_cdd(self):
        """Run verify_cdd in background."""
        try:
            from pathlib import Path
            import verify_cdd as vc

            xlsx_path = Path(self.input_file_path)
            wb = vc._load(xlsx_path, data_only=True)
            if wb is None:
                raise RuntimeError("Workbook missing required 'IP' sheet")

            ip, rows4g, rows5g = vc._gather(wb)
            findings = vc.collect_findings(ip, rows4g, rows5g)

            json_path = xlsx_path.with_suffix(xlsx_path.suffix + ".findings.json")
            csv_path = xlsx_path.with_suffix(xlsx_path.suffix + ".findings.csv")
            vc.write_json(findings, xlsx_path, json_path)
            vc.write_csv(findings, csv_path)

            count = len(findings)
            if count == 0:
                summary = "OK: no findings"
                detail = f"OK: no findings.\n\nReports written:\n{json_path}\n{csv_path}"
            else:
                summary = f"{count} finding(s). See:\n{csv_path.name}"
                detail = (f"Found {count} finding(s).\n\nReports written:\n"
                          f"{json_path}\n{csv_path}")

            self.root.after(0, self._verify_done, summary, detail, count == 0)
        except Exception as e:
            self.root.after(0, self._verify_failed, f"Verification failed: {e}")

    def _autocorrect_cdd(self):
        """Run CDD auto-correction on the selected input file."""
        if not self.input_file_path:
            self._show_error("Please select a CDD input file first")
            return

        self.btn_verify_cdd.config(state=tk.DISABLED, bg=ModernStyle.TEXT_SECONDARY)
        self.btn_autocorrect_cdd.config(state=tk.DISABLED, bg=ModernStyle.TEXT_SECONDARY)
        self._update_status("Auto-correcting CDD...")
        self.lbl_verify_result.config(text="Applying fixes...", fg=ModernStyle.TEXT_SECONDARY)

        thread = threading.Thread(target=self._run_autocorrect_cdd)
        thread.daemon = True
        thread.start()

    def _run_autocorrect_cdd(self):
        """Run verify_cdd --fix in background."""
        try:
            from pathlib import Path
            import verify_cdd as vc

            xlsx_path = Path(self.input_file_path)
            wb = vc._load(xlsx_path, data_only=False)
            if wb is None:
                raise RuntimeError("Workbook missing required 'IP' sheet")

            ip, rows4g, rows5g = vc._gather(wb)
            findings = vc.collect_findings(ip, rows4g, rows5g)
            applied, skipped = vc.apply_fixes(wb, findings)

            out_path = xlsx_path.with_name(xlsx_path.stem + "_fixed" + xlsx_path.suffix)
            wb.save(out_path)

            ip2, rows4g2, rows5g2 = vc._gather(wb)
            remaining = vc.collect_findings(ip2, rows4g2, rows5g2)

            json_path = out_path.with_suffix(out_path.suffix + ".findings.json")
            csv_path = out_path.with_suffix(out_path.suffix + ".findings.csv")
            vc.write_json(remaining, out_path, json_path)
            vc.write_csv(remaining, csv_path)

            summary = (f"Applied {len(applied)} fix(es); "
                       f"{len(remaining)} remaining.\nSaved: {out_path.name}")
            detail = (f"Applied {len(applied)} fix(es).\n"
                      f"{len(remaining)} finding(s) remain (non-deterministic).\n\n"
                      f"Fixed workbook:\n{out_path}\n\n"
                      f"Reports:\n{json_path}\n{csv_path}")

            self.input_file_path = str(out_path)
            self.root.after(0, self._autocorrect_done, str(out_path), summary, detail, len(remaining) == 0)
        except Exception as e:
            self.root.after(0, self._verify_failed, f"Auto-correct failed: {e}")

    def _autocorrect_done(self, out_path, summary, detail, clean):
        """Reload fixed workbook as Input File, then show verification result."""
        self.lbl_input_file.config(text=out_path, fg=ModernStyle.TEXT_PRIMARY)
        thread = threading.Thread(target=self._load_input_file)
        thread.daemon = True
        thread.start()
        self._verify_done(summary, detail, clean)

    def _verify_done(self, summary, detail, clean):
        """Re-enable verify buttons and show result."""
        color = ModernStyle.BG_SUCCESS if clean else ModernStyle.BG_WARNING
        self.lbl_verify_result.config(text=summary, fg=color)
        self._update_status("Verification complete")
        self.btn_verify_cdd.config(state=tk.NORMAL, bg=ModernStyle.BTN_PRIMARY)
        self.btn_autocorrect_cdd.config(state=tk.NORMAL, bg=ModernStyle.BG_WARNING)
        if clean:
            messagebox.showinfo("CDD Verification", detail)
        else:
            messagebox.showwarning("CDD Verification", detail)

    def _verify_failed(self, message):
        """Re-enable verify buttons and show error."""
        self.lbl_verify_result.config(text="Failed", fg="#e74c3c")
        self.btn_verify_cdd.config(state=tk.NORMAL, bg=ModernStyle.BTN_PRIMARY)
        self.btn_autocorrect_cdd.config(state=tk.NORMAL, bg=ModernStyle.BG_WARNING)
        self._show_error(message)

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
