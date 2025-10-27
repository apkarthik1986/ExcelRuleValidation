"""
Excel Rule Validation GUI Application
Provides a Windows GUI for Excel file validation with natural language rules.
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
from typing import List, Optional
import os

from excel_reader import ExcelReader
from rule_parser import RuleParser, Rule
from rule_engine import RuleEngine, ValidationResult


class ExcelRuleValidationApp:
    """
    Main GUI application for Excel rule validation.
    """
    
    def __init__(self, root):
        """Initialize the application."""
        self.root = root
        self.root.title("Excel Rule Validation System")
        self.root.geometry("1200x800")
        
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
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
        # Title
        title_label = ttk.Label(
            main_frame, 
            text="Excel Rule Validation System", 
            font=('Arial', 16, 'bold')
        )
        title_label.grid(row=0, column=0, columnspan=2, pady=10)
        
        # File Selection Section
        self.create_file_section(main_frame)
        
        # Rule Input Section
        self.create_rule_section(main_frame)
        
        # Data Preview Section
        self.create_data_section(main_frame)
        
        # Validation Results Section
        self.create_results_section(main_frame)
        
        # Control Buttons
        self.create_control_section(main_frame)
        
        # Status Bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(
            main_frame, 
            textvariable=self.status_var, 
            relief=tk.SUNKEN, 
            anchor=tk.W
        )
        status_bar.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
    def create_file_section(self, parent):
        """Create file selection section."""
        file_frame = ttk.LabelFrame(parent, text="1. Select Excel File", padding="10")
        file_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
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
        
    def create_rule_section(self, parent):
        """Create rule input section."""
        rule_frame = ttk.LabelFrame(parent, text="2. Define Validation Rules", padding="10")
        rule_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        rule_frame.columnconfigure(0, weight=1)
        
        # Instructions
        instructions = (
            "Enter rules using expression-based syntax. Examples:\n"
            "  • (A>B) AND (X=G) - column comparison with AND\n"
            "  • voltage contains \"cc_r\" - string contains check\n"
            "  • Current>Threshold - column-to-column comparison"
        )
        ttk.Label(
            rule_frame, 
            text=instructions, 
            font=('Arial', 9), 
            foreground='blue'
        ).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # Rule input
        ttk.Label(rule_frame, text="Rule:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.rule_entry = ttk.Entry(rule_frame, width=80)
        self.rule_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5)
        
        ttk.Button(
            rule_frame, 
            text="Add Rule", 
            command=self.add_rule
        ).grid(row=1, column=2, padx=5)
        
        # Rules list
        ttk.Label(rule_frame, text="Active Rules:").grid(row=2, column=0, sticky=tk.W, pady=5)
        
        self.rules_listbox = tk.Listbox(rule_frame, height=5)
        self.rules_listbox.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), padx=5)
        
        scrollbar = ttk.Scrollbar(rule_frame, orient=tk.VERTICAL, command=self.rules_listbox.yview)
        scrollbar.grid(row=3, column=2, sticky=(tk.N, tk.S))
        self.rules_listbox.config(yscrollcommand=scrollbar.set)
        
        ttk.Button(
            rule_frame, 
            text="Remove Selected", 
            command=self.remove_rule
        ).grid(row=4, column=0, sticky=tk.W, pady=5)
        
    def create_data_section(self, parent):
        """Create data preview section."""
        data_frame = ttk.LabelFrame(parent, text="3. Data Preview", padding="10")
        data_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
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
        
    def create_results_section(self, parent):
        """Create validation results section."""
        results_frame = ttk.LabelFrame(parent, text="4. Validation Results", padding="10")
        results_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
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
        control_frame.grid(row=5, column=0, columnspan=2, pady=10)
        
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
            self.data = self.excel_reader.load()
            
            self.display_data_preview()
            self.status_var.set(f"Loaded {len(self.data)} rows from {os.path.basename(file_path)}")
            messagebox.showinfo("Success", f"Loaded {len(self.data)} rows successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel file:\n{str(e)}")
            self.status_var.set("Error loading file")
            
    def display_data_preview(self):
        """Display data in the preview tree."""
        # Clear existing data
        self.data_tree.delete(*self.data_tree.get_children())
        
        if self.data is None or self.data.empty:
            return
        
        # Configure columns
        columns = list(self.data.columns)
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
            
            # Add to listbox
            self.rules_listbox.insert(tk.END, rule_text)
            
            # Clear entry
            self.rule_entry.delete(0, tk.END)
            
            self.status_var.set(f"Added rule: {rule.name}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse rule:\n{str(e)}")
            
    def remove_rule(self):
        """Remove selected rule."""
        selection = self.rules_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a rule to remove.")
            return
        
        index = selection[0]
        self.rules_listbox.delete(index)
        del self.rules[index]
        self.status_var.set("Rule removed")
        
    def validate_data(self):
        """Validate data using defined rules."""
        if self.data is None:
            messagebox.showwarning("Warning", "Please load data first.")
            return
        
        if not self.rules:
            messagebox.showwarning("Warning", "Please add at least one rule.")
            return
        
        try:
            # Run validation
            results = self.rule_engine.validate(self.data, self.rules)
            
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
                messagebox.showerror("Error", f"Failed to export results:\n{str(e)}")


def main():
    """Main entry point for the application."""
    root = tk.Tk()
    app = ExcelRuleValidationApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
