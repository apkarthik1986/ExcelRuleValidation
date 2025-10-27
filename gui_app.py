"""
Excel Rule Validation GUI Application
Provides a Windows GUI for Excel file validation with natural language rules.
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
from typing import List, Optional
import os
import sys
import logging
import traceback
from datetime import datetime

from excel_reader import ExcelReader
from rule_parser import RuleParser, Rule
from rule_engine import RuleEngine, ValidationResult

# Setup logging to file with timestamped filename and also stream to stdout
LOG_DIR = os.path.join(os.path.dirname(__file__), "logs")
os.makedirs(LOG_DIR, exist_ok=True)
_TS = datetime.now().strftime("%Y%m%d_%H%M%S")
LOG_FILE = os.path.join(LOG_DIR, f"gui_{_TS}.log")
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s %(levelname)s: %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


def _handle_uncaught_exception(exc_type, exc_value, exc_tb):
    """Global exception handler that logs uncaught exceptions to the GUI log."""
    tb = ''.join(traceback.format_exception(exc_type, exc_value, exc_tb))
    logger.error("Uncaught exception:\n%s", tb)
    try:
        # Try to show a messagebox if tkinter is initialized
        messagebox.showerror("Fatal Error", f"An unexpected error occurred.\nSee {LOG_FILE} for details.")
    except Exception:
        pass


# Register the global exception handler
sys.excepthook = _handle_uncaught_exception


class ExcelRuleValidationApp:
    """
    Main GUI application for Excel rule validation.
    """
    
    def __init__(self, root):
        """Initialize the application."""
        self.root = root
        self.root.title("Excel Rule Validation System")
        self.root.geometry("1200x800")
        # center window on start
        try:
            self.center_window()
        except Exception:
            pass
        
    def center_window(self):
        """Center the main window on the primary monitor."""
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        if w <= 1 and h <= 1:
            # geometry not yet applied; use requested size
            geom = self.root.winfo_geometry().split('+')[0]
            try:
                w, h = map(int, geom.split('x'))
            except Exception:
                w, h = 1200, 800
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        x = max(0, (screen_w - w) // 2)
        y = max(0, (screen_h - h) // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")
        
        # Data storage
        self.excel_reader: Optional[ExcelReader] = None
        self.data: Optional[pd.DataFrame] = None
        self.rules: List[Rule] = []
        self.rule_parser = RuleParser()
        self.rule_engine = RuleEngine()
        
        # Setup UI
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the user interface."""
        # Configure style
        style = ttk.Style()
        style.theme_use('clam')
        
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights so the UI is responsive
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)

        # Title
        title_label = ttk.Label(
            main_frame, 
            text="Excel Rule Validation System", 
            font=('Arial', 16, 'bold')
        )
        title_label.grid(row=0, column=0, sticky=(tk.W), pady=6)

        # File Selection Section (kept at top so users can load data before switching tabs)
        self.create_file_section(main_frame)

        # Notebook (tabs) to keep the UI compact on small screens
        notebook = ttk.Notebook(main_frame)
        # place notebook below the file selection section (file section uses row=1)
        notebook.grid(row=2, column=0, sticky=(tk.N, tk.S, tk.E, tk.W), pady=8)
        main_frame.rowconfigure(2, weight=1)

        # Create tab frames
        rules_tab = ttk.Frame(notebook)
        data_tab = ttk.Frame(notebook)
        validation_tab = ttk.Frame(notebook)
        examples_tab = ttk.Frame(notebook)

        # Add Data tab first so it appears as the left-most tab on small screens
        notebook.add(data_tab, text='Data')
        notebook.add(rules_tab, text='Rules')
        notebook.add(validation_tab, text='Validation')
        notebook.add(examples_tab, text='Examples')

        # Make tabs fill available space
        for tab in (rules_tab, data_tab, validation_tab, examples_tab):
            tab.columnconfigure(0, weight=1)
            tab.rowconfigure(0, weight=1)

        # Populate tabs with the relevant sections
        # Data tab first for small screens
        self.create_data_section(data_tab)
        self.create_rule_section(rules_tab)
        self.create_results_section(validation_tab)
        # Controls (Validate / Export / Clear) belong in validation tab
        # Place controls under the results in the validation tab
        self.create_control_section(validation_tab)
        # Examples tab
        self.create_examples_section(examples_tab)

        # Select the Data tab programmatically so it's active on startup
        try:
            notebook.select(data_tab)
        except Exception:
            logger.exception('Failed to select Data tab on startup')

        # Status Bar (bottom)
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(
            main_frame, 
            textvariable=self.status_var, 
            relief=tk.SUNKEN, 
            anchor=tk.W
        )
        status_bar.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=5)
        
    def create_file_section(self, parent):
        """Create file selection section."""
        file_frame = ttk.LabelFrame(parent, text="Select Excel File", padding="6")
        # place at top of the parent tab/frame
        file_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        file_frame.columnconfigure(1, weight=1)
        
        self.file_path_var = tk.StringVar()
        
        ttk.Label(file_frame, text="File:").grid(row=0, column=0, sticky=tk.W, padx=5)
        ttk.Entry(
            file_frame, 
            textvariable=self.file_path_var, 
            state='readonly'
        ).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)

        ttk.Button(
            file_frame, 
            text="Browse...", 
            command=self.browse_file
        ).grid(row=0, column=2, padx=5)

        ttk.Button(
            file_frame, 
            text="Load Data", 
            command=self.load_data
        ).grid(row=0, column=3, padx=5)

        # Header row selector (1-based for user friendliness) + auto-detect
        ttk.Label(file_frame, text="Header row (1-based):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=(6,0))
        # default to 1 (first row) for typical spreadsheets
        self.header_row_var = tk.IntVar(value=1)
        hdr_spin = ttk.Spinbox(file_frame, from_=1, to=100, textvariable=self.header_row_var, width=5)
        hdr_spin.grid(row=1, column=1, sticky=tk.W, padx=5, pady=(6,0))
        self.auto_header_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(file_frame, text='Auto-detect header', variable=self.auto_header_var).grid(row=1, column=2, sticky=tk.W, padx=5, pady=(6,0))
        
    def create_rule_section(self, parent):
        """Create rule input section."""
        rule_frame = ttk.LabelFrame(parent, text="2. Define Validation Rules", padding="6")
        rule_frame.grid(sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        rule_frame.columnconfigure(0, weight=1)
        rule_frame.rowconfigure(0, weight=1)

        # Use a horizontal PanedWindow so the rule list and preview are side-by-side
        paned = tk.PanedWindow(rule_frame, orient=tk.HORIZONTAL)
        paned.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))

        left = ttk.Frame(paned)
        right = ttk.Frame(paned)
        paned.add(left)
        paned.add(right)

        # Left pane: rule entry, toolbar, and rules tree
        # Compact instructions (single-line or hidden) to save vertical space
        instr = "Enter rules using expressions. Example: (A>B) AND (X=G) or name contains \"cc_r\""
        ttk.Label(left, text=instr, font=('Arial', 9), foreground='blue').grid(row=0, column=0, sticky=tk.W, pady=(2,6))

        entry_frame = ttk.Frame(left)
        entry_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=4)
        entry_frame.columnconfigure(0, weight=1)

        ttk.Label(entry_frame, text="Rule:").grid(row=0, column=0, sticky=tk.W)
        self.rule_entry = ttk.Entry(entry_frame)
        self.rule_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=6)
        ttk.Button(entry_frame, text="Add", width=8, command=self.add_rule).grid(row=0, column=2, padx=4)

        # Toolbar for quick actions
        toolbar = ttk.Frame(left)
        toolbar.grid(row=2, column=0, sticky=tk.W, pady=6)
        ttk.Button(toolbar, text="Edit", command=self.edit_rule).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Remove", command=self.remove_rule).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Save", command=self.save_rules).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Load", command=self.load_rules).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Enable All", command=self.enable_all_rules).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Disable All", command=self.disable_all_rules).pack(side=tk.LEFT, padx=2)

        # Rules Tree (Enabled, Name, Expression)
        tree_frame = ttk.Frame(left)
        tree_frame.grid(row=3, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        left.rowconfigure(3, weight=1)

        self.rules_tree = ttk.Treeview(tree_frame, columns=('enabled','name','expr'), show='headings')
        self.rules_tree.heading('enabled', text='Enabled')
        self.rules_tree.heading('name', text='Rule')
        self.rules_tree.heading('expr', text='Expression')
        self.rules_tree.column('enabled', width=70, anchor=tk.CENTER)
        self.rules_tree.column('name', width=140, anchor=tk.W)
        self.rules_tree.column('expr', width=360, anchor=tk.W)
        self.rules_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.rules_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.rules_tree.configure(yscrollcommand=scrollbar.set)

        # Bind selection and double-click events
        self.rules_tree.bind('<<TreeviewSelect>>', self.on_rule_select)
        self.rules_tree.bind('<Double-1>', self.on_rule_double_click)

        # Right pane: preview and column inserter
        right.columnconfigure(0, weight=1)
        ttk.Label(right, text="Preview", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky=tk.W, pady=(2,6))
        self.rule_preview = scrolledtext.ScrolledText(right, height=12, wrap=tk.WORD, font=('Consolas', 9), state='disabled')
        self.rule_preview.grid(row=1, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        right.rowconfigure(1, weight=1)

        # Column inserter under preview
        inserter = ttk.Frame(right)
        inserter.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=6)
        ttk.Label(inserter, text="Columns:").grid(row=0, column=0, sticky=tk.W)
        self.column_combo = ttk.Combobox(inserter, values=[], state='readonly', width=30)
        self.column_combo.grid(row=0, column=1, sticky=tk.W, padx=6)
        ttk.Button(inserter, text="Insert", command=self.insert_column).grid(row=0, column=2, padx=6)
        # Rule inserter (dropdown of rule names)
        ttk.Label(inserter, text="Rules:").grid(row=1, column=0, sticky=tk.W, pady=(6,0))
        self.rule_combo = ttk.Combobox(inserter, values=[], state='readonly', width=30)
        self.rule_combo.grid(row=1, column=1, sticky=tk.W, padx=6, pady=(6,0))
        ttk.Button(inserter, text="Insert Rule", command=self.insert_rule_name).grid(row=1, column=2, padx=6, pady=(6,0))
        
    def create_data_section(self, parent):
        """Create data preview section."""
        data_frame = ttk.LabelFrame(parent, text="Data Preview", padding="6")
        # place at top of data tab
        data_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        data_frame.columnconfigure(0, weight=1)
        data_frame.rowconfigure(0, weight=1)
        
        # Create Treeview for data preview
        self.data_tree = ttk.Treeview(data_frame, height=8)
        self.data_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbars
        vsb = ttk.Scrollbar(data_frame, orient=tk.VERTICAL, command=self.data_tree.yview)
        vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        hsb = ttk.Scrollbar(data_frame, orient=tk.HORIZONTAL, command=self.data_tree.xview)
        hsb.grid(row=1, column=0, sticky=(tk.W, tk.E))
        self.data_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    def create_examples_section(self, parent):
        """Create examples tab with many sample expressions (double-click to insert)."""
        frame = ttk.LabelFrame(parent, text="Sample Expressions", padding="6")
        frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W), pady=5)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        lb_frame = ttk.Frame(frame)
        lb_frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        lb_frame.columnconfigure(0, weight=1)
        lb_frame.rowconfigure(0, weight=1)

        self.examples_list = tk.Listbox(lb_frame, height=20)
        self.examples_list.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        ex_scroll = ttk.Scrollbar(lb_frame, orient=tk.VERTICAL, command=self.examples_list.yview)
        ex_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.examples_list.configure(yscrollcommand=ex_scroll.set)

        examples = [
            "A > 100",
            "B = 'OK'",
            "name contains 'alpha'",
            "status in ('NEW','PENDING')",
            "amount between 10 and 100",
            "date >= '2025-01-01'",
            "columnX startswith 'ABC'",
            "NOT (flag = 'Y')",
            "(A > 10) AND (B < 5)",
            "(A > 10) AND (B < 5 OR C = 'x')",
            "((A>10 AND B<5) OR (C='x' AND D contains 'zz'))",
            "(A > 100 AND (B < 50 OR (C = 'X' AND D contains 'Z')))",
            "(A > 0 AND (B > 0 AND (C > 0 OR D > 0)))",
            "(Country = 'US' AND (State = 'CA' OR State = 'NY'))",
            "((A between 1 and 5) OR (B between 10 and 20)) AND NOT (C is null)",
            "regex_match(email, '^[^@]+@example\\.com$')",
            "trim(name) = 'John'",
            "LOWER(category) = 'widget'",
            "(A > 100 AND (B < 5 OR (C = 'x' AND (D > 0 OR E < 0))))",
            "((A > 10 AND B < 5) OR (C = 'x' AND (D > 0 OR (E < 0 AND F != 'Z'))))",
        ]

        for ex in examples:
            self.examples_list.insert(tk.END, ex)

        # allow double-click to insert into rule entry
        self.examples_list.bind('<Double-1>', self.insert_example)
        # right-click context menu to copy sample(s)
        try:
            # create popup menu attached to the listbox
            self.examples_menu = tk.Menu(self.examples_list, tearoff=0)
            self.examples_menu.add_command(label='Copy', command=self.copy_selected_example)
            self.examples_menu.add_command(label='Copy All', command=self.copy_all_examples)
            # bind right-click (Button-3 on Windows) to show the menu and select the clicked item
            self.examples_list.bind('<Button-3>', self.show_examples_context_menu)
            # also support Mac/other where middle button might be used
            self.examples_list.bind('<Button-2>', self.show_examples_context_menu)
            # keyboard shortcut for convenience
            self.examples_list.bind('<Control-c>', lambda e: self.copy_selected_example())
        except Exception:
            logger.exception('Failed to attach examples context menu')

    def detect_header_row(self, file_path, max_rows=10):
        """Attempt to auto-detect which row contains headers.

        Returns 0-based header row index or 0 if detection fails.
        """
        try:
            df = pd.read_excel(file_path, header=None, nrows=max_rows)
            # choose the row with maximum non-null values
            best_idx = 0
            best_count = -1
            for i in range(min(len(df), max_rows)):
                count = df.iloc[i].count()
                if count > best_count:
                    best_count = count
                    best_idx = i
            return int(best_idx)
        except Exception:
            logger.exception('Auto-detect header failed')
            return 0
        
    def create_results_section(self, parent):
        """Create validation results section."""
        results_frame = ttk.LabelFrame(parent, text="Validation Results", padding="6")
        # Place at top of validation tab
        results_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        
        self.results_text = scrolledtext.ScrolledText(
            results_frame, 
            height=10, 
            wrap=tk.WORD,
            font=('Consolas', 9)
        )
        self.results_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
    def create_control_section(self, parent):
        """Create control buttons section."""
        control_frame = ttk.Frame(parent)
        # place controls below results in validation tab
        control_frame.grid(row=1, column=0, columnspan=2, pady=10)
        
        ttk.Button(
            control_frame, 
            text="Validate Data", 
            command=self.validate_data,
            style='Accent.TButton'
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            control_frame, 
            text="Clear Results", 
            command=self.clear_results
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            control_frame, 
            text="Export Results", 
            command=self.export_results
        ).pack(side=tk.LEFT, padx=5)
        
    def browse_file(self):
        """Browse for Excel file."""
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.file_path_var.set(filename)
            self.status_var.set(f"Selected: {os.path.basename(filename)}")
            
    def load_data(self):
        """Load Excel data."""
        file_path = self.file_path_var.get()
        if not file_path:
            messagebox.showwarning("Warning", "Please select an Excel file first.")
            return
        
        try:
            self.excel_reader = ExcelReader(file_path)
            # header_row_var is 1-based for the user; convert to 0-based for pandas
            if getattr(self, 'auto_header_var', tk.BooleanVar(value=False)).get():
                hdr = self.detect_header_row(file_path)
            else:
                hdr = int(getattr(self, 'header_row_var', tk.IntVar(value=2)).get()) - 1
            self.data = self.excel_reader.load(header_row=hdr)
            
            self.display_data_preview()
            self.status_var.set(f"Loaded {len(self.data)} rows from {os.path.basename(file_path)}")
            messagebox.showinfo("Success", f"Loaded {len(self.data)} rows successfully!")
            
        except Exception as e:
            logger.exception("Failed to load Excel file: %s", file_path)
            messagebox.showerror("Error", f"Failed to load Excel file:\n{str(e)}")
            self.status_var.set("Error loading file")
            
    def display_data_preview(self):
        """Display data in the preview tree."""
        # Clear existing data
        self.data_tree.delete(*self.data_tree.get_children())
        
        if self.data is None or self.data.empty:
            return
        
        # Configure columns, ensuring duplicate names are suffixed
        orig_columns = list(self.data.columns)
        counts = {}
        new_columns = []
        for col in orig_columns:
            if col not in counts:
                counts[col] = 0
                new_columns.append(col)
            else:
                counts[col] += 1
                new_columns.append(f"{col}{counts[col]}")
        # apply unique column names back to dataframe
        try:
            self.data.columns = new_columns
        except Exception:
            logger.exception('Failed to set unique column names')
        columns = new_columns
        self.data_tree['columns'] = columns
        self.data_tree['show'] = 'headings'
        
        # Set column headings
        for col in columns:
            self.data_tree.heading(col, text=col)
            self.data_tree.column(col, width=100, minwidth=50)
        
        # Add data (first 100 rows for preview)
        for index, row in self.data.head(100).iterrows():
            values = [str(val) for val in row.tolist()]
            self.data_tree.insert('', tk.END, values=values)
        
        # Populate column combobox for rule insertion (if present)
        try:
            if hasattr(self, 'column_combo'):
                self.column_combo['values'] = columns
                if columns:
                    self.column_combo.current(0)
        except Exception:
            logger.exception("Failed to populate column combobox")
            
    def add_rule(self):
        """Add a new rule."""
        rule_text = self.rule_entry.get().strip()
        if not rule_text:
            messagebox.showwarning("Warning", "Please enter a rule.")
            return
        
        if self.data is None:
            messagebox.showwarning("Warning", "Please load data first.")
            return
        
        try:
            columns = self.data.columns.tolist()
            # Auto-generate rule name from expression
            rule_name = f"Rule{len(self.rules) + 1}"
            rule = self.rule_parser.parse_rule(rule_text, columns, rule_name=rule_name)
            self.rules.append(rule)
            
            # Add to rules tree (show enabled, name, expression)
            display_text = f"{rule.name}: {rule_text}"
            enabled_text = 'Yes' if getattr(rule, 'enabled', True) else 'No'
            self.rules_tree.insert('', tk.END, values=(enabled_text, rule.name, rule_text))
            # Update rules dropdown (for easy insertion into expressions)
            try:
                if hasattr(self, 'rule_combo'):
                    vals = [r.name for r in self.rules]
                    self.rule_combo['values'] = vals
                    if vals:
                        self.rule_combo.current(len(vals)-1)
            except Exception:
                logger.exception('Failed to update rule combo')
            
            # Clear entry
            self.rule_entry.delete(0, tk.END)
            
            self.status_var.set(f"Added rule: {rule.name}")
            
        except Exception as e:
            logger.exception("Failed to parse rule: %s", rule_text)
            messagebox.showerror("Error", f"Failed to parse rule:\n{str(e)}")
            
    def remove_rule(self):
        """Remove selected rule(s)."""
        selection = self.rules_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a rule to remove.")
            return

        # Remove the first selected rule (and corresponding object)
        item = selection[0]
        vals = self.rules_tree.item(item, 'values')
        rule_name = vals[1] if len(vals) > 1 else None
        if rule_name:
            # Find rule index by name
            idx = next((i for i, r in enumerate(self.rules) if r.name == rule_name), None)
            if idx is not None:
                del self.rules[idx]
        self.rules_tree.delete(item)
        self.status_var.set("Rule removed")
        # refresh rule combo
        try:
            if hasattr(self, 'rule_combo'):
                vals = [r.name for r in self.rules]
                self.rule_combo['values'] = vals
                if vals:
                    self.rule_combo.current(0)
        except Exception:
            logger.exception('Failed to refresh rule combo after removal')

    def edit_rule(self):
        """Edit the selected rule in a modal dialog (Treeview-based)."""
        selection = self.rules_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a rule to edit.")
            return

        item = selection[0]
        vals = self.rules_tree.item(item, 'values')
        rule_name = vals[1] if len(vals) > 1 else None
        orig_text = vals[2] if len(vals) > 2 else ''
        idx = next((i for i, r in enumerate(self.rules) if r.name == rule_name), None)
        orig_rule = self.rules[idx] if idx is not None else None

        # Create modal dialog
        dlg = tk.Toplevel(self.root)
        dlg.title(f"Edit Rule - {rule_name}")
        dlg.transient(self.root)
        dlg.grab_set()

        ttk.Label(dlg, text=f"Editing {rule_name}").grid(row=0, column=0, columnspan=2, pady=5, padx=5)
        entry = ttk.Entry(dlg, width=80)
        entry.grid(row=1, column=0, columnspan=2, pady=5, padx=5)
        entry.insert(0, orig_text)

        def on_save():
            new_text = entry.get().strip()
            if not new_text:
                messagebox.showwarning("Warning", "Rule expression cannot be empty.")
                return
            try:
                # Re-parse the rule using the same rule name to validate/update
                columns = self.data.columns.tolist() if self.data is not None else []
                new_rule = self.rule_parser.parse_rule(new_text, columns, rule_name=rule_name)
                # replace in our rules list if index known
                if idx is not None:
                    self.rules[idx] = new_rule
                # Update tree item
                enabled_text = 'Yes' if getattr(new_rule, 'enabled', True) else 'No'
                self.rules_tree.item(item, values=(enabled_text, new_rule.name, new_text))
                # refresh rule combo
                try:
                    if hasattr(self, 'rule_combo'):
                        vals = [r.name for r in self.rules]
                        self.rule_combo['values'] = vals
                except Exception:
                    logger.exception('Failed to refresh rule combo after edit')
                self.status_var.set(f"Updated rule: {new_rule.name}")
                dlg.destroy()
            except Exception as e:
                logger.exception("Failed to update rule %s", rule_name)
                messagebox.showerror("Error", f"Failed to parse/validate rule:\n{str(e)}")

        def on_cancel():
            dlg.destroy()

        ttk.Button(dlg, text="Save", command=on_save).grid(row=2, column=0, pady=10, padx=5)
        ttk.Button(dlg, text="Cancel", command=on_cancel).grid(row=2, column=1, pady=10, padx=5)

        entry.focus_set()
        dlg.wait_window()

    def insert_column(self):
        """Insert the selected column name into the rule entry at the cursor position."""
        try:
            col = self.column_combo.get()
            if not col:
                messagebox.showwarning("Warning", "No column selected to insert.")
                return
            # Insert at current cursor position
            pos = self.rule_entry.index(tk.INSERT)
            current = self.rule_entry.get()
            new = current[:pos] + col + current[pos:]
            self.rule_entry.delete(0, tk.END)
            self.rule_entry.insert(0, new)
            # Restore cursor after inserted column
            self.rule_entry.icursor(pos + len(col))
            self.rule_entry.focus_set()
        except Exception as e:
            logger.exception("Failed to insert column into rule entry")
            messagebox.showerror("Error", f"Failed to insert column:\n{str(e)}")

    def insert_rule_name(self):
        """Insert the selected rule name from the rule combobox into the rule entry."""
        try:
            name = self.rule_combo.get()
            if not name:
                messagebox.showwarning("Warning", "No rule selected to insert.")
                return
            pos = self.rule_entry.index(tk.INSERT)
            current = self.rule_entry.get()
            new = current[:pos] + name + current[pos:]
            self.rule_entry.delete(0, tk.END)
            self.rule_entry.insert(0, new)
            self.rule_entry.icursor(pos + len(name))
            self.rule_entry.focus_set()
        except Exception:
            logger.exception('Failed to insert rule name')
            messagebox.showerror('Error', 'Failed to insert rule name')

    def insert_example(self, event):
        """Insert the double-clicked example expression into the rule entry."""
        try:
            sel = self.examples_list.curselection()
            if not sel:
                return
            ex = self.examples_list.get(sel[0])
            # Replace the rule entry with the example (user can modify)
            self.rule_entry.delete(0, tk.END)
            self.rule_entry.insert(0, ex)
            self.rule_entry.focus_set()
        except Exception:
            logger.exception('Failed to insert example')

    def show_examples_context_menu(self, event):
        """Show right-click context menu for the examples list and select clicked item."""
        try:
            # find the nearest item to the click and select it
            idx = self.examples_list.nearest(event.y)
            if idx is not None:
                # clear existing selection and select the clicked line
                self.examples_list.selection_clear(0, tk.END)
                self.examples_list.selection_set(idx)
                # ensure the clicked item is visible
                self.examples_list.see(idx)
            # show popup menu
            self.examples_menu.tk_popup(event.x_root, event.y_root)
        except Exception:
            logger.exception('Failed to show examples context menu')

    def copy_selected_example(self):
        """Copy the currently selected example text to the clipboard."""
        try:
            sel = self.examples_list.curselection()
            if not sel:
                self.status_var.set('No sample selected to copy')
                return
            text = self.examples_list.get(sel[0])
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.status_var.set('Copied sample to clipboard')
        except Exception:
            logger.exception('Failed to copy selected example')
            messagebox.showerror('Error', 'Failed to copy sample to clipboard')

    def copy_all_examples(self):
        """Copy all sample expressions (one per line) to the clipboard."""
        try:
            all_text = '\n'.join(self.examples_list.get(0, tk.END))
            if not all_text:
                self.status_var.set('No samples available to copy')
                return
            self.root.clipboard_clear()
            self.root.clipboard_append(all_text)
            self.status_var.set('Copied all samples to clipboard')
        except Exception:
            logger.exception('Failed to copy all examples')
            messagebox.showerror('Error', 'Failed to copy all samples to clipboard')
        
    def save_rules(self):
        """Save current rules to a JSON file."""
        if not self.rules:
            messagebox.showwarning("Warning", "No rules to save.")
            return

        filename = filedialog.asksaveasfilename(
            title="Save Rules",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not filename:
            return

        try:
            import json
            data = [r.to_dict() for r in self.rules]
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2)
            messagebox.showinfo("Saved", f"Saved {len(self.rules)} rules to {filename}")
            self.status_var.set(f"Saved {len(self.rules)} rules")
        except Exception:
            logger.exception("Failed to save rules to %s", filename)
            messagebox.showerror("Error", f"Failed to save rules:\n{str(Exception)}")

    def load_rules(self):
        """Load rules from a JSON file; ask whether to replace or append."""
        filename = filedialog.askopenfilename(
            title="Load Rules",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not filename:
            return

        try:
            import json
            with open(filename, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if not isinstance(data, list):
                messagebox.showerror("Error", "Invalid rules file format (expected a list of rules).")
                return

            # Ask whether to replace existing rules
            if self.rules:
                resp = messagebox.askyesno("Replace Rules?", "Replace existing rules with loaded rules? (Yes = replace, No = append)")
            else:
                resp = True

            if resp:
                # clear existing
                self.rules.clear()
                for iid in list(self.rules_tree.get_children()):
                    self.rules_tree.delete(iid)
                self.rule_parser.rules.clear()
                self.rule_parser.rules_by_name.clear()

            count = 0
            for rd in data:
                rule = Rule.from_dict(rd)
                # add to parser and local list
                self.rule_parser.rules.append(rule)
                self.rule_parser.rules_by_name[rule.name] = rule
                self.rules.append(rule)
                enabled_text = 'Yes' if getattr(rule, 'enabled', True) else 'No'
                self.rules_tree.insert('', tk.END, values=(enabled_text, rule.name, rule.description or ''))
                count += 1

            # refresh rule combobox
            try:
                if hasattr(self, 'rule_combo'):
                    vals = [r.name for r in self.rules]
                    self.rule_combo['values'] = vals
            except Exception:
                logger.exception('Failed to refresh rule combo after load')

            messagebox.showinfo("Loaded", f"Loaded {count} rules from {filename}")
            self.status_var.set(f"Loaded {count} rules")
        except Exception:
            logger.exception("Failed to load rules from %s", filename)
            messagebox.showerror("Error", f"Failed to load rules:\n{str(Exception)}")

    def enable_all_rules(self):
        """Enable all loaded rules and update the Treeview."""
        try:
            for i, rule in enumerate(self.rules):
                rule.enabled = True
            # update treeview rows
            for item in self.rules_tree.get_children():
                vals = list(self.rules_tree.item(item, 'values'))
                vals[0] = 'Yes'
                self.rules_tree.item(item, values=vals)
            self.status_var.set('All rules enabled')
            # refresh preview
            self.on_rule_select(None)
        except Exception:
            logger.exception('Failed to enable all rules')
            messagebox.showerror('Error', 'Failed to enable all rules')

    def disable_all_rules(self):
        """Disable all loaded rules and update the Treeview."""
        try:
            for i, rule in enumerate(self.rules):
                rule.enabled = False
            # update treeview rows
            for item in self.rules_tree.get_children():
                vals = list(self.rules_tree.item(item, 'values'))
                vals[0] = 'No'
                self.rules_tree.item(item, values=vals)
            self.status_var.set('All rules disabled')
            # refresh preview
            self.on_rule_select(None)
        except Exception:
            logger.exception('Failed to disable all rules')
            messagebox.showerror('Error', 'Failed to disable all rules')
        
    def on_rule_select(self, event):
        """Display a small preview of the parsed conditions for the selected rule."""
        try:
            selection = self.rules_tree.selection()
            if not selection:
                # clear preview
                self.rule_preview.config(state='normal')
                self.rule_preview.delete(1.0, tk.END)
                self.rule_preview.config(state='disabled')
                return

            item = selection[0]
            vals = self.rules_tree.item(item, 'values')
            rule_name = vals[1] if len(vals) > 1 else None
            rule = next((r for r in self.rules if r.name == rule_name), None)
            self.rule_preview.config(state='normal')
            self.rule_preview.delete(1.0, tk.END)
            if rule:
                # show description and each parsed condition
                self.rule_preview.insert(tk.END, f"Name: {rule.name}\n")
                self.rule_preview.insert(tk.END, f"Description: {rule.description}\n")
                self.rule_preview.insert(tk.END, "Conditions:\n")
                for i, cond in enumerate(rule.conditions):
                    line = f"  - {cond.column} {cond.operator.value} {cond.value}\n"
                    self.rule_preview.insert(tk.END, line)
                if rule.logical_ops:
                    ops = ' '.join([op.value for op in rule.logical_ops])
                    self.rule_preview.insert(tk.END, f"Logical Ops: {ops}\n")
                self.rule_preview.insert(tk.END, f"Enabled: {getattr(rule, 'enabled', True)}\n")
            self.rule_preview.config(state='disabled')
        except Exception:
            logger.exception("Failed to display rule preview")

    def on_rule_double_click(self, event):
        """Toggle enabled state when double-clicking a rule row."""
        try:
            rowid = self.rules_tree.identify_row(event.y)
            if not rowid:
                return
            vals = self.rules_tree.item(rowid, 'values')
            rule_name = vals[1] if len(vals) > 1 else None
            if not rule_name:
                return
            # find rule
            idx = next((i for i, r in enumerate(self.rules) if r.name == rule_name), None)
            if idx is None:
                return
            rule = self.rules[idx]
            # toggle
            rule.enabled = not getattr(rule, 'enabled', True)
            enabled_text = 'Yes' if rule.enabled else 'No'
            # update tree cell
            self.rules_tree.item(rowid, values=(enabled_text, rule.name, vals[2] if len(vals)>2 else ''))
            self.status_var.set(f"Rule {rule.name} {'enabled' if rule.enabled else 'disabled'}")
            # refresh preview
            self.on_rule_select(None)
        except Exception:
            logger.exception("Failed to toggle rule enabled state")
        
    def validate_data(self):
        """Validate data using defined rules."""
        if self.data is None:
            messagebox.showwarning("Warning", "Please load data first.")
            return
        
        if not self.rules:
            messagebox.showwarning("Warning", "Please add at least one rule.")
            return
        
        try:
            # Run validation on enabled rules only
            enabled_rules = [r for r in self.rules if getattr(r, 'enabled', True)]
            results = self.rule_engine.validate(self.data, enabled_rules)
            
            # Display results
            report = self.rule_engine.generate_report()
            self.results_text.delete(1.0, tk.END)
            self.results_text.insert(1.0, report)
            
            # Update status
            passed = len(self.rule_engine.get_passed_validations())
            failed = len(self.rule_engine.get_failed_validations())
            self.status_var.set(f"Validation complete: {passed} passed, {failed} failed")
            
            messagebox.showinfo(
                "Validation Complete", 
                f"Validated {len(results)} items\nPassed: {passed}\nFailed: {failed}"
            )
            
        except Exception as e:
            logger.exception("Validation failed")
            messagebox.showerror("Error", f"Validation failed:\n{str(e)}")
            
    def clear_results(self):
        """Clear validation results."""
        self.results_text.delete(1.0, tk.END)
        self.status_var.set("Results cleared")
        
    def export_results(self):
        """Export validation results to a file."""
        if not self.rule_engine.results:
            messagebox.showwarning("Warning", "No results to export.")
            return
        
        filename = filedialog.asksaveasfilename(
            title="Export Results",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                report = self.rule_engine.generate_report()
                with open(filename, 'w') as f:
                    f.write(report)
                messagebox.showinfo("Success", f"Results exported to {filename}")
                self.status_var.set(f"Results exported to {os.path.basename(filename)}")
            except Exception as e:
                logger.exception("Failed to export results to %s", filename)
                messagebox.showerror("Error", f"Failed to export results:\n{str(e)}")


def main():
    """Main entry point for the application."""
    logger.info("Starting Excel Rule Validation GUI")
    try:
        root = tk.Tk()
        app = ExcelRuleValidationApp(root)
        # Bring window to front briefly so it isn't hidden off-screen
        try:
            # ensure window is deiconified, focused and briefly topmost
            root.deiconify()
            root.update()
            root.lift()
            root.attributes('-topmost', True)
            root.focus_force()
            # remove topmost shortly after to avoid stealing focus permanently
            root.after(1500, lambda: root.attributes('-topmost', False))
        except Exception:
            logger.exception('Failed to force window to front')
        root.mainloop()
    except Exception:
        logger.exception("Unhandled exception in main loop")
        raise
    finally:
        logger.info("GUI exited")


if __name__ == "__main__":
    main()
