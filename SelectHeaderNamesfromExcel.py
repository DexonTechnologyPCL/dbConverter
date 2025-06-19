
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import json
from tkinter.scrolledtext import ScrolledText
from datetime import datetime
import traceback
import sys

class HeaderSelector:
    def __init__(self, root):
        self.root = root
        self.root.title("Header Names Selector")
        self.root.geometry("900x600")
        self.root.minsize(700, 500)
        
        # Variables for storing data
        self.excel_file = ""
        self.pipe_tally_headers = []
        self.all_sheets_data = {}
        self.selected_headers = {}
        self.sheet_processing_mode = {}
        
        # Import conversion functions
        try:
            from ExcelToAccessDB import excel_to_access
            self.excel_to_access = excel_to_access
            self.conversion_available = True
            print("✅ Excel to Access conversion: Available")
        except ImportError:
            self.conversion_available = False
            print("⚠️ Excel to Access conversion: Not available")

        self.setup_ui()         # Create UI
        self.setup_styles()     # Set style

    def setup_styles(self):
        """Set styles for the UI"""
        style = ttk.Style()
        
        # Define colors for tags in Treeview
        style.configure("Match.Treeview", foreground="green")
        style.configure("NoMatch.Treeview", foreground="red")
        style.configure("Missing.Treeview", foreground="orange")
     
    def setup_ui(self):
        """Create a User Interface"""
        self.create_header_section()    # Header Section
        self.create_file_section()      # File Selection Section
        self.create_main_content()      # Main Content Section
        self.create_control_section()   # Control Buttons Section
        self.create_results_section()   # Results Section
    
    def create_header_section(self):
        """Create the Header section"""
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill="x", padx=10, pady=5)
        
        title_label = ttk.Label(header_frame, text="Header Selector Tool", font=("Arial", 16, "bold"))
        title_label.pack()

    def create_file_section(self):
        """Create the file selection section"""
        file_frame = ttk.LabelFrame(self.root, text="Select Excel File")
        file_frame.pack(fill="x", padx=10, pady=5)
        
        inner_frame = ttk.Frame(file_frame)
        inner_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(inner_frame, text="File:").pack(side="left")
        
        self.file_label = ttk.Label(inner_frame, text="No file selected", foreground="gray", font=("Arial", 9))
        self.file_label.pack(side="left", padx=(10, 0))
        
        ttk.Button(inner_frame, text="🔍 Browse", command=self.browse_file).pack(side="right")
        
        # Status bar
        self.status_frame = ttk.Frame(file_frame)
        self.status_frame.pack(fill="x", padx=10, pady=(0, 5))
        
        self.pipe_tally_status = ttk.Label(self.status_frame, text="📋 Status: Sheet 'List of Pipe Tally' not found yet", foreground="red")
        self.pipe_tally_status.pack(side="left")

    def create_main_content(self):
        """Create the main content section"""
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Divide into 2 sections: Left=Standard Headers, Right=Sheet Comparison
        left_frame = ttk.LabelFrame(main_frame, text="📊 Standard Headers from 'List of Pipe Tally'")
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))
        
        right_frame = ttk.LabelFrame(main_frame, text="🔍 Compare with Other Sheets")
        right_frame.pack(side="right", fill="both", expand=True, padx=(5, 0))
      
        self.setup_standard_headers_section(left_frame)         # Left Section: Standard Headers
        self.setup_comparison_section(right_frame)              # Right Section: Sheet Comparison

    def setup_standard_headers_section(self, parent):
        """Create the Standard Headers section"""
        # Scrollable frame for headers
        canvas = tk.Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        self.standard_frame = ttk.Frame(canvas)
        
        self.standard_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        canvas.create_window((0, 0), window=self.standard_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Add mouse wheel support
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def setup_comparison_section(self, parent):
        """Create the comparison section"""
        # Notebook for tabs
        self.comparison_notebook = ttk.Notebook(parent)
        self.comparison_notebook.pack(fill="both", expand=True, pady=5)

    def create_control_section(self):
        """Create the control buttons section"""
        control_frame = ttk.Frame(self.root)
        control_frame.pack(fill="x", padx=10, pady=5)
        
        # Left buttons - Selection
        left_buttons = ttk.Frame(control_frame)
        left_buttons.pack(side="left")
        
        ttk.Button(left_buttons, text="✅ Select All", command=self.select_all, width=12).pack(side="left", padx=(0, 5))
        ttk.Button(left_buttons, text="❌ Cancel All", command=self.deselect_all, width=12).pack(side="left", padx=(0, 5))
        
        # Right buttons - Actions
        right_buttons = ttk.Frame(control_frame)
        right_buttons.pack(side="right")
        
        ttk.Button(right_buttons, text="📋 Display Results", command=self.show_results, width=15).pack(side="right", padx=(5, 0))
        ttk.Button(right_buttons, text="🔄 Convert to AccessDB", command=self.convert_to_access, width=20).pack(side="right")

    def create_results_section(self):
        """Create the results display section"""
        result_frame = ttk.LabelFrame(self.root, text="📄 Results and Summary")
        result_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.result_text = ScrolledText(result_frame, height=8, font=("Consolas", 9))
        self.result_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Display initial message
        welcome_text = """
        🔧 Header Selector Tool - Ready to use

        📋 How to use:
        1. Click "🔍 Browse" to select Excel file with "List of Pipe Tally" sheet
        2. Check Standard Headers loaded from "List of Pipe Tally"
        3. View comparison with other sheets on the right side
        4. Select headers you want to use as standard
        5. Click "📋 Display Results" to view summary
        6. Click "🔄 Convert to Access DB" to convert file
        """
        self.result_text.insert(tk.END, welcome_text)

    def browse_file(self):
        """Open dialog to select Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files (*.xlsx)", "*.xlsx"), ("Excel files (*.xls)", "*.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            self.excel_file = file_path
            filename = os.path.basename(file_path)
            self.file_label.config(text=f"📁 {filename}", foreground="black")
            self.load_excel_data()

    def load_excel_data(self):
        """Load data from Excel file"""
        try:
            # Show loading
            self.pipe_tally_status.config(text="⏳ Loading data...", foreground="blue")
            self.root.update()
            self.clear_all_data()                                               # Clear old data
           
            xls = pd.ExcelFile(self.excel_file)                                 # Read Excel file
            pipe_tally_sheet = self.find_pipe_tally_sheet(xls.sheet_names)      # Find "List of Pipe Tally" sheet
            
            if pipe_tally_sheet:
                self.load_standard_headers(xls, pipe_tally_sheet)               # Load Standard Headers
                self.load_other_sheets(xls, pipe_tally_sheet)                   # Load other sheets
                self.update_ui_after_load()                                     # Update UI
                
                messagebox.showinfo("✅ Success", 
                    f"Data loaded successfully!\n\n"
                    f"📊 Standard Headers: {len(self.pipe_tally_headers)} items\n"
                    f"📋 Other Sheets: {len(self.all_sheets_data)} sheets\n"
                    f"📁 From file: {os.path.basename(self.excel_file)}")
            else:
                messagebox.showwarning("⚠️ Warning", 
                    "Sheet 'List of Pipe Tally' not found in this Excel file\n\n"
                    "📋 Sheets found:\n" + "\n".join([f"• {name}" for name in xls.sheet_names[:10]]) +
                    ("\n... and more" if len(xls.sheet_names) > 10 else "") +
                    "\n\n💡 Please check the sheet name")
            
        except Exception as e:
            error_msg = f"Cannot read Excel file\n\n❌ Error: {str(e)}"
            messagebox.showerror("❌ Error", error_msg)
            self.pipe_tally_status.config(text="❌ Error loading file", foreground="red")

    def find_pipe_tally_sheet(self, sheet_names):
        """Find sheet with name similar to 'List of Pipe Tally'"""
        # Try to find exact match
        for sheet_name in sheet_names:
            if "List of Pipe Tally" in sheet_name:
                return sheet_name
        
        # Try to find similar names
        keywords = ["pipe tally", "pipe", "tally", "list"]
        for sheet_name in sheet_names:
            sheet_lower = sheet_name.lower()
            if any(keyword in sheet_lower for keyword in keywords):
                return sheet_name
        
        return None

    def load_standard_headers(self, xls, pipe_tally_sheet):
        """Load Standard Headers from List of Pipe Tally"""
        try:
            df = pd.read_excel(xls, sheet_name=pipe_tally_sheet, header=None, nrows=3)
            self.pipe_tally_headers = self.extract_headers(df)
            
            # Update status
            self.pipe_tally_status.config(
                text=f"✅ Found {len(self.pipe_tally_headers)} headers from '{pipe_tally_sheet}'",
                foreground="green"
            )
            
        except Exception as e:
            self.pipe_tally_status.config(
                text=f"❌ Error reading '{pipe_tally_sheet}': {str(e)}",
                foreground="red"
            )
            raise e

    def load_other_sheets(self, xls, pipe_tally_sheet):
        """Load data from other sheets"""
        self.all_sheets_data = {}
        
        for sheet_name in xls.sheet_names:
            if sheet_name == pipe_tally_sheet:
                continue
            
            try:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=3)
                headers = self.extract_headers(df)
                self.all_sheets_data[sheet_name] = headers
                
            except Exception as e:
                print(f"Warning: Cannot read sheet '{sheet_name}': {str(e)}")
                continue

    def update_ui_after_load(self):
        """Update UI after loading data"""
        self.create_standard_checkboxes()       # Create checkboxes for standard headers
        self.create_comparison_tabs()           # Create comparison tabs

    def create_standard_checkboxes(self):
        """Create checkboxes for standard headers"""
        # Clear old widgets
        for widget in self.standard_frame.winfo_children():
            widget.destroy()
        
        self.selected_headers = {}
        
        for i, header in enumerate(self.pipe_tally_headers):
            var = tk.BooleanVar()
            self.selected_headers[header] = var
            
            frame = ttk.Frame(self.standard_frame)
            frame.pack(fill="x", padx=5, pady=1)
            
            # Checkbox with command to update tabs when changed
            cb = ttk.Checkbutton(frame, text=f"{i+1:2d}. {header}", variable=var, command=self.on_header_selection_changed)
            cb.pack(side="left")

    def create_comparison_tabs(self):
        """Create tabs for comparison"""
        # Clear old tabs
        for tab in self.comparison_notebook.tabs():
            self.comparison_notebook.forget(tab)
        
        for sheet_name, headers in self.all_sheets_data.items():
            self.create_comparison_tab(sheet_name, headers)

    def create_comparison_tab(self, sheet_name, headers):
        """Create tab for comparing headers"""
        tab_frame = ttk.Frame(self.comparison_notebook)
        
        # Shorten tab name
        tab_name = sheet_name[:12] + "..." if len(sheet_name) > 12 else sheet_name
        self.comparison_notebook.add(tab_frame, text=tab_name)
        
        # Processing mode selection for this sheet
        mode_frame = ttk.LabelFrame(tab_frame, text=f"⚙️ Processing '{sheet_name}'")
        mode_frame.pack(fill="x", padx=5, pady=5)
        
        # Create variable for this sheet's mode
        sheet_mode = tk.StringVar(value="standard")
        self.sheet_processing_mode[sheet_name] = sheet_mode
        
        mode_inner_frame = ttk.Frame(mode_frame)
        mode_inner_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Radiobutton(mode_inner_frame, text="📊 Use Standard Headers", variable=sheet_mode, value="standard",
                       command=lambda: self.update_sheet_preview(sheet_name)).pack(side="left")
        ttk.Radiobutton(mode_inner_frame, text="🎯 Use Own Headers", variable=sheet_mode, value="individual",
                       command=lambda: self.update_sheet_preview(sheet_name)).pack(side="left", padx=(8, 0))
        
        # Create Treeview
        tree_frame = ttk.Frame(tab_frame)
        tree_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        tree = ttk.Treeview(tree_frame, columns=("Status", "Header"), show="headings", height=10)
        
        tree.heading("Status", text="Status")
        tree.heading("Header", text="Header Name")
        
        tree.column("Status", width=80, anchor="center")
        tree.column("Header", width=300)
        
        # Store tree reference for updates
        setattr(tab_frame, 'tree', tree)
        setattr(tab_frame, 'sheet_headers', headers)
        setattr(tab_frame, 'sheet_name', sheet_name)
        
        # Show initial data
        self.populate_tree_data(tree, sheet_name, headers, "standard")
        
        # Statistics at the top
        stats_frame = ttk.Frame(tab_frame)
        stats_frame.pack(fill="x", padx=5, pady=(0, 5))
        
        stats_label = ttk.Label(stats_frame, text="", font=("Arial", 9))
        stats_label.pack(side="left")
        setattr(tab_frame, 'stats_label', stats_label)
        
        # Update statistics
        self.update_sheet_statistics(tab_frame, sheet_name, headers, "standard")
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        h_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        tree.pack(side="left", fill="both", expand=True)
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")

    def on_header_selection_changed(self):
        """Called when header selection changes - update all tabs"""
        # Update all comparison tabs to reflect current selection
        for tab_id in self.comparison_notebook.tabs():
            tab_frame = self.comparison_notebook.nametowidget(tab_id)
            if hasattr(tab_frame, 'sheet_name'):
                sheet_name  = tab_frame.sheet_name
                headers     = tab_frame.sheet_headers
                mode        = self.sheet_processing_mode[sheet_name].get()
                
                # Clear and repopulate tree
                tree = tab_frame.tree
                for item in tree.get_children():
                    tree.delete(item)
                
                self.populate_tree_data(tree, sheet_name, headers, mode)
                self.update_sheet_statistics(tab_frame, sheet_name, headers, mode)

    def update_sheet_preview(self, sheet_name):
        """Update display when mode changes"""
        # Find matching tab
        for tab_id in self.comparison_notebook.tabs():
            tab_frame = self.comparison_notebook.nametowidget(tab_id)
            if hasattr(tab_frame, 'sheet_name') and tab_frame.sheet_name == sheet_name:
                mode    = self.sheet_processing_mode[sheet_name].get()
                headers = tab_frame.sheet_headers
                
                # Clear old data in tree
                tree = tab_frame.tree
                for item in tree.get_children():
                    tree.delete(item)
                
                # Show new data
                self.populate_tree_data(tree, sheet_name, headers, mode)
                self.update_sheet_statistics(tab_frame, sheet_name, headers, mode)
                break

    def populate_tree_data(self, tree, sheet_name, headers, mode):
        """Fill data in Treeview according to selected mode"""
        if mode == "standard":
            # Standard Headers mode - show only selected headers
            selected_standard_headers = self.get_selected_headers()
            
            if not selected_standard_headers:
                # If no headers selected, show message
                tree.insert("", "end", values=("⚠️", "No Standard Headers Selected"), tags=("warning",))
                tree.tag_configure("warning", foreground="orange")
                return

            matching_count = 0
            
            # Show headers that exist in current sheet
            for header in headers:
                if header in selected_standard_headers:
                    tree.insert("", "end", values=("✅ Use", header), tags=("use",))
                    matching_count += 1
                else:
                    tree.insert("", "end", values=("❌ Skip", header), tags=("skip",))
            
            # Show selected standard headers that are missing (will be added as empty columns)
            for header in selected_standard_headers:
                if header not in headers:
                    tree.insert("", "end", values=("➕ Add", header), tags=("add",))
        
        else:  # mode == "individual" 
            # Use own headers mode
            for header in headers:
                tree.insert("", "end", values=("✅ Use", header), tags=("use",))
        
        # Set colors
        tree.tag_configure("use", foreground="green")
        tree.tag_configure("skip", foreground="red")
        tree.tag_configure("add", foreground="blue")

    def update_sheet_statistics(self, tab_frame, sheet_name, headers, mode):
        """Update sheet statistics"""
        if mode == "standard":
            selected_standard_headers = self.get_selected_headers()
            
            if not selected_standard_headers:
                stats_text = f"📊 Standard Mode: No headers selected | Sheets: {len(headers)} | Output: 0"
            else:
                matching_count      = sum(1 for h in headers if h in selected_standard_headers)
                total_selected      = len(selected_standard_headers)
                match_percentage    = (matching_count / total_selected) * 100 if total_selected > 0 else 0
                
                stats_text = f"📊 Standard: {matching_count}/{total_selected} ({match_percentage:.0f}%) | Output: {total_selected}"
        else:
            stats_text = f"🎯 Individual Mode: Use original headers {len(headers)} items | Output: {len(headers)} columns"
        
        tab_frame.stats_label.config(text=stats_text)

    def extract_headers(self, df):
        """Extract headers from DataFrame (same method as original code)"""
        if len(df) < 2: return []
        
        first_row   = df.iloc[0].fillna("").tolist()
        second_row  = df.iloc[1].fillna("").tolist()
        
        headers     = []
        max_cols    = max(len(first_row), len(second_row))
        
        for i in range(max_cols):
            second_val  = second_row[i] if i < len(second_row) else ""
            first_val   = first_row[i] if i < len(first_row) else ""
            
            if str(second_val).strip() != '':
                headers.append(str(second_val).strip())
            elif str(first_val).strip() != '':
                headers.append(str(first_val).strip())
            else:
                headers.append(f"Unnamed_{i}")
        
        return headers

    def count_header_in_sheets(self, header):
        """Count number of sheets that have this header"""
        count = 0
        for sheet_headers in self.all_sheets_data.values():
            if header in sheet_headers:
                count += 1
        return count

    def select_all(self):
        """Select all headers"""
        for var in self.selected_headers.values():
            var.set(True)
        
        # Update tabs to reflect changes
        self.on_header_selection_changed()
        
        messagebox.showinfo("✅ Complete", f"Selected all {len(self.selected_headers)} headers")

    def deselect_all(self):
        """Deselect all headers"""
        for var in self.selected_headers.values():
            var.set(False)
        
        # Update tabs to reflect changes
        self.on_header_selection_changed()
        
        messagebox.showinfo("❌ Complete", "Deselected all headers")

    def get_selected_headers(self):
        """Get list of selected headers"""
        selected = []
        for header, var in self.selected_headers.items():
            if var.get():
                selected.append(header)
        return selected

    def show_results(self):
        """Display selected results"""
        selected = self.get_selected_headers()
        
        self.result_text.delete(1.0, tk.END)
        
        if not selected:
            self.result_text.insert(tk.END, "⚠️ No headers selected\n\n"
                                            "💡 Please select headers you want to use as standard")
            return
        
        # Generate report
        report = self.generate_report(selected)
        self.result_text.insert(tk.END, report)

    def generate_report(self, selected_headers):
        """Generate results report"""
        total_headers   = len(self.pipe_tally_headers)
        total_sheets    = len(self.all_sheets_data)
        
        # Count sheets by mode
        standard_mode_count = 0
        individual_mode_count = 0
        
        for sheet_name, mode_var in self.sheet_processing_mode.items():
            if mode_var.get() == "standard":
                standard_mode_count += 1
            else:
                individual_mode_count += 1
        
        report = f"""
        {'='*80}
        📊 Header Selection and Processing
        {'='*80}

        📁 File: {os.path.basename(self.excel_file)}

        📈 Overview:
        ├─ 📋 Total Standard Headers: {total_headers} items
        ├─ ✅ Selected Standard Headers: {len(selected_headers)} items ({(len(selected_headers)/total_headers)*100:.1f}%)
        ├─ 📊 Total sheets: {total_sheets} sheets
        ├─ 🎯 Use Standard Mode: {standard_mode_count} sheets
        ├─ 🔧 Use Individual Mode: {individual_mode_count} sheets
        └─ 📈 Coverage level: {self.calculate_coverage(selected_headers):.1f}%

        {'─'*80}
        📋 Selected Standard Headers:
        {'─'*80}
        """
        
        for i, header in enumerate(selected_headers, 1):
            count = self.count_header_in_sheets(header)
            coverage = (count / total_sheets) * 100 if total_sheets > 0 else 0
            
            if count == total_sheets:
                status = "✅ All sheets"
            elif count > 0:
                status = f"⚠️ {count}/{total_sheets} sheets ({coverage:.0f}%)"
            else:
                status = "❌ Not found in other sheets"
            
            report += f"{i:2d}. {header:<50} {status}\n"
        
        report += f"""
        {'─'*80}
        🔧 Processing for each sheet:
        {'─'*80}
        """
        
        for sheet_name in self.all_sheets_data.keys():
            mode = self.sheet_processing_mode[sheet_name].get()
            headers_count = len(self.all_sheets_data[sheet_name])
            
            if mode == "standard":
                matching = len([h for h in self.all_sheets_data[sheet_name] if h in selected_headers])
                report += f"📊 {sheet_name}:\n"
                report += f"   • Mode: Standard Headers ({len(selected_headers)} columns)\n"
                report += f"   • Original headers: {headers_count} items\n"
                report += f"   • Match with Standard: {matching}/{len(selected_headers)} items\n"
            else:
                report += f"🎯 {sheet_name}:\n"
                report += f"   • Mode: Individual Headers ({headers_count} columns)\n"
                report += f"   • Original headers: {headers_count} items (use all)\n"
            
            report += "\n"
        
        report += f"""
        {'─'*80}
        💡 Application:
        {'─'*80}
        1. 📝 Sheets using Standard Mode will have headers as selected
        2. 🎯 Sheets using Individual Mode will use original headers
        3. ➕ Standard Mode: Missing columns will be added as empty columns
        4. ➖ Standard Mode: Unselected headers will be ignored
        5. 🔄 Individual Mode: Use all original headers
        {'='*80}
        """
        
        return report

    def convert_to_access(self):
        """Convert Excel to Access Database with selected headers"""
        if not self.excel_file:
            messagebox.showwarning("⚠️ Warning", "Please select Excel file first")
            return
    
        if not self.conversion_available:
            messagebox.showerror("❌ Error", 
                "Excel to Access conversion not available.\n"
                "Please ensure ExcelToAccessDB.py is in the same directory.")
            return
    
        # Check if any headers are selected
        selected_headers = self.get_selected_headers()
        if not selected_headers:
            messagebox.showwarning("⚠️ Warning", 
                "Please select at least one header before converting.\n"
                "Select headers from the Standard Headers section.")
            return
    
        # Check for Access drivers first
        if not self.check_access_drivers():
            return
    
        # Show conversion dialog with selected headers info
        result = messagebox.askyesno(
            "🔄 Convert to Access Database",
            f"Convert Excel file to Access Database?\n\n"
            f"📁 Source: {os.path.basename(self.excel_file)}\n"
            f"📊 Selected Headers: {len(selected_headers)} items\n"
            f"📋 Sheets: {len(self.all_sheets_data) + 1} sheets\n\n"
            f"✅ Will use SELECTED headers for processing"
        )
    
        if not result:
            return
    
        # Convert with selected headers
        self.convert_with_selected_headers()

    def create_temp_header_file(self, selected_headers):
        """Create a temporary Excel file with selected headers for conversion"""
        try:
            # Create temporary filename
            temp_filename = f"temp_headers_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
            # Create a DataFrame with selected headers
            header_data = {
                'Header Name': selected_headers,
                'Order': list(range(1, len(selected_headers) + 1)),
                'Selected': ['Yes'] * len(selected_headers)
            }
        
            df = pd.DataFrame(header_data)
        
            # Write to Excel file
            with pd.ExcelWriter(temp_filename, engine='openpyxl') as writer:
                # Create "List of Pipe Tally" sheet with selected headers
                header_df = pd.DataFrame([selected_headers, selected_headers])  # Two rows as per original format
                header_df.to_excel(writer, sheet_name='List of Pipe Tally', index=False, header=False)
            
                # Create summary sheet
                df.to_excel(writer, sheet_name='Selected Headers', index=False)
        
            return temp_filename
        
        except Exception as e:
            print(f"Error creating temp header file: {str(e)}")
            return None

    def convert_with_selected_headers(self):
        """Convert Excel to Access Database using selected headers"""
        # Clear results area and show conversion progress
        self.result_text.delete(1.0, tk.END)
        selected_headers = self.get_selected_headers()
    
        self.result_text.insert(tk.END, "🔄 Starting Excel to Access Database conversion...\n")
        self.result_text.insert(tk.END, f"📁 Source file: {os.path.basename(self.excel_file)}\n")
        self.result_text.insert(tk.END, f"📊 Using {len(selected_headers)} selected headers\n")
        self.result_text.insert(tk.END, "⏳ Please wait, this may take a few moments...\n\n")
        self.result_text.update()
    
        try:
            # Show selected headers being used
            self.result_text.insert(tk.END, "📋 Selected headers for conversion:\n")
            for i, header in enumerate(selected_headers, 1):
                self.result_text.insert(tk.END, f"  {i:2d}. {header}\n")
            self.result_text.insert(tk.END, "\n")
        
            # Show sheet processing modes
            self.result_text.insert(tk.END, "⚙️ Sheet processing modes:\n")
            for sheet_name, mode_var in self.sheet_processing_mode.items():
                mode = mode_var.get()
                mode_text = "Use Standard Headers" if mode == "standard" else "Use Own Headers"
                self.result_text.insert(tk.END, f"  • {sheet_name}: {mode_text}\n")
            self.result_text.insert(tk.END, "\n")
            self.result_text.update()
        
            # Prepare sheet modes dictionary
            sheet_modes = {}
            for sheet_name, mode_var in self.sheet_processing_mode.items():
                sheet_modes[sheet_name] = mode_var.get()
        
            # Log conversion start
            self.result_text.insert(tk.END, f"🚀 Starting conversion process...\n")
            self.result_text.insert(tk.END, f"📊 Processing {len(self.all_sheets_data) + 1} sheets\n")
            self.result_text.see(tk.END)
            self.result_text.update()
        
            # Convert with selected headers and sheet modes
            print(f"Converting file: {self.excel_file}")
            print(f"Selected headers: {selected_headers}")
            print(f"Sheet modes: {sheet_modes}")
        
            result = self.excel_to_access(self.excel_file, 
                                        header_file=None, 
                                        selected_headers=selected_headers,
                                        sheet_modes=sheet_modes)
        
            if result:
                access_file = os.path.splitext(self.excel_file)[0] + ".accdb"
            
                # Show success message in results panel
                self.result_text.insert(tk.END, "\n" + "="*60 + "\n")
                self.result_text.insert(tk.END, "✅ CONVERSION COMPLETED SUCCESSFULLY!\n")
                self.result_text.insert(tk.END, "="*60 + "\n\n")
            
                self.result_text.insert(tk.END, f"📁 Source file: {os.path.basename(self.excel_file)}\n")
                self.result_text.insert(tk.END, f"💾 Output file: {os.path.basename(access_file)}\n")
                self.result_text.insert(tk.END, f"📂 Location: {os.path.dirname(access_file)}\n")
                self.result_text.insert(tk.END, f"📊 Sheets processed: {len(self.all_sheets_data) + 1}\n")
                self.result_text.insert(tk.END, f"📋 Headers selected: {len(selected_headers)} items\n")
                self.result_text.insert(tk.END, f"📅 Completed at: {datetime.now().strftime('%H:%M:%S')}\n")
            
                # Show which sheets used which mode
                self.result_text.insert(tk.END, "\n📊 Processing summary:\n")
                standard_count = sum(1 for mode in sheet_modes.values() if mode == "standard")
                individual_count = len(sheet_modes) - standard_count
                self.result_text.insert(tk.END, f"  • {standard_count} sheets used Selected Headers\n")
                self.result_text.insert(tk.END, f"  • {individual_count} sheets used Original Headers\n")
            
                self.result_text.insert(tk.END, "\n🎉 Your Access database is ready to use!")
            
                messagebox.showinfo("✅ Success", 
                    f"Conversion completed!\n"
                    f"Selected headers: {len(selected_headers)} items\n"
                    f"Standard mode: {standard_count} sheets\n"
                    f"Individual mode: {individual_count} sheets\n"
                    f"Output: {os.path.basename(access_file)}")
            
            else:
                self.result_text.insert(tk.END, "\n❌ CONVERSION FAILED\n")
                
        except Exception as conv_error:
            error_details = str(conv_error)
            self.result_text.insert(tk.END, f"\n❌ CONVERSION ERROR\n{error_details}\n")
    
        finally:
            self.result_text.see(tk.END)
            self.result_text.update()

    def check_access_drivers(self):
        """Check if Microsoft Access drivers are available"""
        try:
            import pyodbc
            drivers = pyodbc.drivers()
            
            # Check for Access drivers
            access_drivers = [driver for driver in drivers if 'Access' in driver or 'ACE' in driver or 'Jet' in driver]
            
            if not access_drivers:
                # No Access drivers found
                error_msg = """
                ❌ Microsoft Access Driver Not Found
                The system cannot find Microsoft Access Database drivers required for conversion.

                🔧 Solutions:
                1. Download and install Microsoft Access Database Engine:
                    • For 64-bit systems: AccessDatabaseEngine_X64.exe
                    • For 32-bit systems: AccessDatabaseEngine.exe
                    • Download from Microsoft official website

                2. Alternative: Install Microsoft Office with Access component
                3. For developers: Install Microsoft Access Runtime
                4. Try running the program as Administrator"""
                
                messagebox.showerror("❌ Access Driver Error", error_msg)
                return False
            
            print(f"✅ Found Access drivers: {access_drivers}")
            return True
            
        except Exception as e:
            messagebox.showerror("❌ Driver Check Error", 
                f"Cannot check Access drivers:\n\n{str(e)}\n\n"
                f"Please ensure Microsoft Access Database Engine is installed.")
            return False

    def clear_all_data(self):
        """Clear all data"""
        self.pipe_tally_headers = []
        self.all_sheets_data = {}
        self.selected_headers = {}
        self.sheet_processing_mode = {}
        
        # Clear UI
        for widget in self.standard_frame.winfo_children():
            widget.destroy()
        
        for tab in self.comparison_notebook.tabs():
            self.comparison_notebook.forget(tab)
        
        self.pipe_tally_status.config(text="📋 Status: Sheet 'List of Pipe Tally' not found yet", foreground="red")

    def clear_results(self):
        """Clear the results text area"""
        self.result_text.delete(1.0, tk.END)
        welcome_text = """
            🔧 Header Selector Tool - Ready to use

            📋 How to use:
            1. Click "🔍 Browse" to select Excel file with "List of Pipe Tally" sheet
            2. Check Standard Headers loaded from "List of Pipe Tally"
            3. View comparison with other sheets on the right side
            4. Select headers you want to use as standard
            5. Click "📋 Display Results" to view summary
            6. Click "🔄 Convert to Access DB" to convert file
        """
        self.result_text.insert(tk.END, welcome_text)

def main():
    """Main function to run the program"""
    root = tk.Tk()
    
    # Set icon if available
    try:
        # root.iconbitmap('icon.ico')  # Add icon file if available
        pass
    except:
        pass
    
    app = HeaderSelector(root)
    
    # Center on screen
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main()