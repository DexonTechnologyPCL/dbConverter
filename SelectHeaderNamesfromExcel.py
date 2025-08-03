import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import json
from tkinter.scrolledtext import ScrolledText
from datetime import datetime
import traceback
import sys
import threading
import time

class HeaderSelector:
    def __init__(self, root):
        self.root = root
        self.root.title("Header Names Selector")
        self.root.geometry("1000x700")
        self.root.minsize(800, 600)
        
        # Variables for storing data
        self.excel_file = ""
        self.all_sheets_data = {}
        self.selected_sheet = ""                                # Selected sheet in dropdown
        self.current_sheet_headers = []                         # Headers of the selected sheet
        self.standard_headers = self.Set_standard_headers()     # 26 standard columns
        self.excel_header_widgets = {}                          # Store widgets for excel headers
        self.standard_mapping_status = {}                       # Store mapping status of standard headers
        
        self.sheet_mappings = {}                               # Store mapping of all sheets
        self.sheet_selected_headers = {}                       # Store selected headers of all sheets
        
        # Progress tracking variables
        self.conversion_cancelled       = False
        self.progress_var               = None
        self.progress_label             = None
        self.progress_bar               = None
        self.convert_button             = None
        
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

    def extract_headers(self, df):
        """Extract headers from DataFrame"""
        if len(df) < 2: 
            return []
        
        first_row = df.iloc[0].fillna("").tolist()
        second_row = df.iloc[1].fillna("").tolist()
        
        headers = []
        max_cols = max(len(first_row), len(second_row))
        
        for i in range(max_cols):
            second_val = second_row[i] if i < len(second_row) else ""
            first_val = first_row[i] if i < len(first_row) else ""
            
            if str(second_val).strip() != '':
                headers.append(str(second_val).strip())
            elif str(first_val).strip() != '':
                headers.append(str(first_val).strip())
            else:
                headers.append(f"Unnamed_{i}")
        
        return headers

    def setup_styles(self):
        """Set styles for the UI"""
        style = ttk.Style()
        
        # Define colors for tags in Treeview
        style.configure("Match.Treeview", foreground="green")
        style.configure("NoMatch.Treeview", foreground="red")
        style.configure("Missing.Treeview", foreground="orange")
        style.configure("Mapped.Treeview", foreground="blue")
     
    def Set_standard_headers(self):
        """Set standard columns"""
        return [
            "Log distance (m)",
            "Latitude (degree)",
            "Longitude (degree)",
            "Altitude (m)",
            "Feature type",
            "Feature identification",
            "Anomaly identification",
            "Girth weld Nr",
            "Joint manufacturing type",
            "Joint / component length (m)",
            "Nominal internal diameter (mm)",
            "Nominal thickness (mm)",
            "Measure/Reference thickness (mm)",
            "Abs Dist to upstream weld (m)",
            "Clock position seam / anomaly",
            "Surface location",
            "Remaining thickness (mm)",
            "Max depth (mm)",
            "Max depth (%)",
            "Length (mm)",
            "Width (mm)",
            "Metal loss anomaly dimension classification",
            "ERF",
            "Comments"
        ]

    def get_sheet_specific_headers(self, sheet_name):
        """Get headers specific to each sheet type"""
        if "List of Nominal Wall Thickness" in sheet_name:
            return [
                "Log distance (m)",
                "Girth weld Nr",
                "Nominal thickness (mm)",
                "Joint manufacturing type",
                "SMYS (psi)",
                "Design Pressure (psi)",
                "MAOP (psi)"
            ]
        else:
            # Default: return all 26 standard headers
            return self.Set_standard_headers()

    def setup_ui(self):
        """Create a User Interface"""
        self.create_header_section()    # Header Section
        self.create_file_section()      # File Selection Section
        self.create_main_content()      # Main Content Section
        self.create_control_section()   # Control Buttons Section
        self.create_progress_section()  # Progress Section
        self.create_results_section()   # Results Section

    def create_header_section(self):
        """Create the Header section"""
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill="x", padx=10, pady=3)
        
        title_label = ttk.Label(header_frame, text="Header Selector Tool", font=("Arial", 14, "bold"))
        title_label.pack()

    def create_file_section(self):
        """Create the file selection section"""
        file_frame = ttk.LabelFrame(self.root, text="Select Excel File")
        file_frame.pack(fill="x", padx=10, pady=3)
        
        inner_frame = ttk.Frame(file_frame)
        inner_frame.pack(fill="x", padx=10, pady=3)
        
        ttk.Label(inner_frame, text="File:").pack(side="left")
        
        self.file_label = ttk.Label(inner_frame, text="No file selected", foreground="gray", font=("Arial", 8))
        self.file_label.pack(side="left", padx=(10, 0))
        
        ttk.Button(inner_frame, text="🔍 Browse", command=self.browse_file).pack(side="right")
        
        # Sheet selection row
        sheet_row = ttk.Frame(file_frame)
        sheet_row.pack(fill="x", padx=10, pady=(0, 3))
        
        ttk.Label(sheet_row, text="Sheet:").pack(side="left")
        
        self.sheet_combo = ttk.Combobox(sheet_row, state="readonly", width=40)
        self.sheet_combo.pack(side="left", padx=(10, 0))
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_selected)
        
        # Status
        self.status_label = ttk.Label(sheet_row, text="📋 Please select Excel file first", foreground="orange")
        self.status_label.pack(side="left", padx=(20, 0))

    def create_main_content(self):
        """Create the main content section"""
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=3)
    
        paned_window = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned_window.pack(fill="both", expand=True)
    
        # Left Frame - Standard Headers
        left_frame = ttk.LabelFrame(paned_window, text="📂 Excel Sheet Headers")
        paned_window.add(left_frame, weight=1)
    
        # Right Frame - Sheet Comparison  
        right_frame = ttk.LabelFrame(paned_window, text="📊 Standard Headers Mapping")
        paned_window.add(right_frame, weight=1) 

        paned_window.update_idletasks()
        total_width = paned_window.winfo_width()
        if total_width > 1:
            paned_window.sashpos(0, total_width // 2)
        else:
            def set_sash_position():
                paned_window.sashpos(0, paned_window.winfo_width() // 2)
            paned_window.after(100, set_sash_position)
    
        self.setup_excel_headers_section(left_frame)
        self.setup_standard_mapping_section(right_frame)

    def setup_excel_headers_section(self, parent):
        """Excel Headers section"""
        # Info frame
        info_frame = ttk.Frame(parent)
        info_frame.pack(fill="x", padx=3, pady=3)
        
        self.excel_info_label = ttk.Label(info_frame, text="No sheet selected", font=("Arial", 8, "bold"))
        self.excel_info_label.pack()

        # Scrollable frame for excel headers
        canvas = tk.Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        self.excel_frame = ttk.Frame(canvas)
        
        self.excel_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        canvas.create_window((0, 0), window=self.excel_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Add mouse wheel support
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def setup_standard_mapping_section(self, parent):
        """Standard Headers section with Mapping"""
        # Info frame
        info_frame = ttk.Frame(parent)
        info_frame.pack(fill="x", padx=3, pady=3)
       
        self.standard_info_label = ttk.Label(info_frame, text="Standard Headers", font=("Arial", 8, "bold"))
        self.standard_info_label.pack()

        # Treeview for showing standard headers mapping
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill="both", expand=True, padx=3, pady=3)
        
        # Create Treeview
        self.standard_tree = ttk.Treeview(tree_frame, columns=("No", "Standard", "Mapped_To", "Status"), show="headings", height=20)
        
        # Configure columns
        self.standard_tree.heading("No",        text="No.")
        self.standard_tree.heading("Standard",  text="Standard Header")
        self.standard_tree.heading("Mapped_To", text="Mapped To")
        self.standard_tree.heading("Status",    text="Status")
        
        self.standard_tree.column("No",         width=40, anchor="center")
        self.standard_tree.column("Standard",   width=220)
        self.standard_tree.column("Mapped_To",  width=220)
        self.standard_tree.column("Status",     width=100, anchor="center")
        
        # Scrollbars for treeview
        v_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.standard_tree.yview)
        h_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.standard_tree.xview)
        
        self.standard_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Pack treeview and scrollbars
        self.standard_tree.pack(side="left", fill="both", expand=True)
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")
        
        # Initialize with empty data
        self.populate_standard_tree()

    def create_excel_header_widgets(self):
        """Create widgets for Excel headers"""
        # Clear old widgets
        for widget in self.excel_frame.winfo_children():
            widget.destroy()
        
        self.excel_header_widgets = {}

        if not self.current_sheet_headers:
            no_data_label = ttk.Label(self.excel_frame, text="No headers to display", foreground="gray")
            no_data_label.pack(pady=20)
            return
        
        # Group headers by base name (for handling duplicates)
        header_groups = self.group_duplicate_headers(self.current_sheet_headers)
        
        row_index = 0
        for base_name, headers in header_groups.items():
            row_index += 1
            
            # Create frame for each row
            row_frame = ttk.Frame(self.excel_frame)
            row_frame.pack(fill="x", padx=5, pady=2)
            
            # Number label
            num_label = ttk.Label(row_frame, text=f"{row_index:2d}.", width=4, anchor="w")
            num_label.pack(side="left")
            
            if len(headers) == 1:
                # Single header - just show label
                header_label = ttk.Label(row_frame, text=headers[0], width=50, anchor="w")
                header_label.pack(side="left")
                
                # Status label
                status_label = ttk.Label(row_frame, text="Ready", width=15, foreground="blue")
                status_label.pack(side="left")
                
                # Store reference
                self.excel_header_widgets[headers[0]] = {
                    'type': 'single',
                    'widget': header_label,
                    'status': status_label,
                    'selected': headers[0]
                }
                
            else:
                # Multiple headers - show dropdown
                header_var = tk.StringVar(value=headers[0])  # Default to first
                
                # Base name label
                base_label = ttk.Label(row_frame, text=f"{base_name}:", width=25, anchor="w")
                base_label.pack(side="left")
                
                # Dropdown for selecting which duplicate to use
                header_combo = ttk.Combobox(row_frame, textvariable=header_var, values=headers, state="readonly", width=30)
                header_combo.pack(side="left", padx=(5, 0))
                header_combo.bind("<<ComboboxSelected>>", lambda e, base=base_name: self.on_excel_header_selected(base))
                
                # Status label
                status_label = ttk.Label(row_frame, text="Dropdown", width=15, foreground="orange")
                status_label.pack(side="left")
                
                # Store reference
                self.excel_header_widgets[base_name] = {
                    'type': 'multiple',
                    'widget': header_combo,
                    'status': status_label,
                    'headers': headers,
                    'variable': header_var,
                    'selected': headers[0]
                }

    def group_duplicate_headers(self, headers):
        """Group headers with similar base names, maintain Excel order"""
        from collections import defaultdict
        
        groups = defaultdict(list)
        header_order = {}
        
        print(f"📋 Grouping headers from Excel (in order):")
        
        for i, header in enumerate(headers):
            group_key = self.create_grouping_key(header)
            groups[group_key].append(header)
            
            if group_key not in header_order:
                header_order[group_key] = i
            
            print(f"  {i+1:2d}. {header} -> group: {group_key}")
        
        # Convert to ordered dict, maintaining Excel appearance order
        result = {}
        sorted_groups = sorted(groups.items(), key=lambda x: header_order[x[0]])
        
        for group_key, header_list in sorted_groups:
            if len(header_list) == 1:
                result[header_list[0]] = header_list                                    # Single header
            else:
                base_name       = self.extract_base_name(header_list[0])                # Multiple headers - group them
                ordered_headers = sorted(header_list, key=lambda x: headers.index(x))   # Sort by original Excel order
                result[base_name] = ordered_headers
                print(f"📦 Grouped: {base_name} -> {ordered_headers}")
        
        return result

    def create_grouping_key(self, header):
        """Create key for grouping - excluding different units"""
        import re

        # Special cases that should not be combined
        special_cases = [
            ("max. depth", "[%]"),  # Max. depth [%] 
            ("max. depth", "[mm]"), # Max. depth [mm]
        ]

        header_lower = header.lower()
        
        # Check special cases
        for base_term, suffix in special_cases:
            if base_term in header_lower and suffix in header_lower:
                return header  # Use full name as key (no grouping)

        # For normal cases - remove suffix in the last parentheses
        try:
            pattern = r'\s*\([^)]*\)\s*$'
            base = re.sub(pattern, '', header)
            base = base.strip()
            return base if base else header
        except:
            return header

    def extract_base_name(self, header):
        """Extract main name from header"""
        import re
        try:
            pattern = r'\s*\([^)]*\)\s*$'
            base = re.sub(pattern, '', header)
            return base.strip() if base.strip() else header
        except:
            return header

    def on_excel_header_selected(self, base_name):
        """Called when selecting header from dropdown"""
        if base_name in self.excel_header_widgets:
            widget_info = self.excel_header_widgets[base_name]
            if widget_info['type'] == 'multiple':
                selected = widget_info['variable'].get()
                widget_info['selected'] = selected
                widget_info['status'].config(text="Selected", foreground="green")

                self.update_standard_tree()         # Update mapping display
                self.save_current_sheet_mapping()   # Save mapping immediately

    def populate_standard_tree(self):
        """Populate initial data in standard headers tree"""
        # Clear existing items
        for item in self.standard_tree.get_children():
            self.standard_tree.delete(item)
        
        current_headers = self.get_current_standard_headers()   # Get headers for current sheet
        
        # Add standard headers
        for i, header in enumerate(current_headers, 1):
            self.standard_tree.insert("", "end", values=(i, header, "<Not Mapped>", "❌"), tags=("unmapped",))
        
        # Configure tags
        self.standard_tree.tag_configure("unmapped", foreground="red")
        self.standard_tree.tag_configure("mapped", foreground="green")
        self.standard_tree.tag_configure("hardcoded", foreground="blue")

    def get_current_standard_headers(self):
        """Get standard headers for current selected sheet"""
        if self.selected_sheet:
            return self.get_sheet_specific_headers(self.selected_sheet)
        else:
            return self.Set_standard_headers()

    def update_standard_tree(self):
        """Update standard tree according to current mapping"""
        # Clear existing items
        for item in self.standard_tree.get_children():
            self.standard_tree.delete(item)
        
        # Get current headers and available excel headers
        current_headers = self.get_current_standard_headers()
        available_excel_headers = self.get_selected_excel_headers()
        
        # Update info label
        sheet_type = "List of Nominal Wall Thickness" if "List of Nominal Wall Thickness" in self.selected_sheet else "Standard"
        header_count = len(current_headers)
        if sheet_type == "List of Nominal Wall Thickness":
            self.standard_info_label.config(text=f"Standard Headers for {sheet_type} ({header_count} Columns)")
        else:
            self.standard_info_label.config(text=f"Standard Headers ({header_count} Columns)")
        
        # Update each standard header
        for i, standard_header in enumerate(current_headers, 1):
            mapped_to, tag = self.find_mapped_excel_header(standard_header, available_excel_headers)
            
            if mapped_to:
                if tag == "hardcoded":
                    status = "🔗"
                    display_text = f"{mapped_to}"
                else:
                    status = "✅"
                    display_text = mapped_to
            else:
                status = "❌"
                display_text = "<Not Mapped>"
                tag = "unmapped"
            
            self.standard_tree.insert("", "end", values=(i, standard_header, display_text, status), tags=(tag,))
        
        # Configure tag colors
        self.standard_tree.tag_configure("unmapped", foreground="red")
        self.standard_tree.tag_configure("mapped", foreground="green")
        self.standard_tree.tag_configure("hardcoded", foreground="blue")

    def get_selected_excel_headers(self):
        """Get list of selected excel headers"""
        selected_headers = []
        
        for key, widget_info in self.excel_header_widgets.items():
            selected_headers.append(widget_info['selected'])
        
        return selected_headers

    def find_mapped_excel_header(self, standard_header, excel_headers):
        """Find excel header that should map to standard header including hard-coded mapping"""
        
        # Hard-coded mapping for ERF
        if standard_header == "ERF":
            for excel_header in excel_headers:
                if any(erf_pattern in excel_header.lower() for erf_pattern in 
                      ["erf (metal loss)", "erf (modified)", "erf(metal loss)", "erf(modified)"]):
                    return excel_header, "hardcoded"

        # Hard-coded mapping for Comments
        if standard_header == "Comments":
            for excel_header in excel_headers:
                if any(comment_pattern in excel_header.lower() for comment_pattern in 
                    ["comment", "comments"]):
                    return excel_header, "hardcoded"

        # Other auto mappings
        best_match = self.find_best_match(standard_header, excel_headers)
        if best_match:
            return best_match, "mapped"
        
        return None, "unmapped"

    def find_best_match(self, standard_header, excel_headers):
        """Find the most suitable excel header"""
        import re
        
        best_match = None
        best_score = 0
        standard_clean = self.clean_header_for_matching(standard_header)        # Clean standard header
        
        for excel_header in excel_headers:
            excel_clean = self.clean_header_for_matching(excel_header)
            
            # Exact match
            if standard_clean == excel_clean: return excel_header
            score = self.calculate_similarity(standard_clean, excel_clean)      # Calculate similarity
            
            if score > best_score and score > 0.5:  # Threshold
                best_score = score
                best_match = excel_header
        
        return best_match    

    def clean_header_for_matching(self, header):
        """Clean header for matching"""
        import re
        
        clean = header.lower()
        clean = clean.replace('[', '(').replace(']', ')')
        clean = re.sub(r'[.,;:!?]', '', clean)
        clean = clean.replace('/', ' ')
        clean = ' '.join(clean.split())
        
        return clean

    def calculate_similarity(self, str1, str2):
        """Calculate similarity"""
        if not str1 or not str2:
            return 0
        
        words1 = set(str1.split())
        words2 = set(str2.split())
        
        if not words1 or not words2:
            return 0
        
        intersection = len(words1.intersection(words2))
        union = len(words1.union(words2))
        
        return intersection / union if union > 0 else 0

    def create_control_section(self):
        """Create control buttons section"""
        control_frame = ttk.Frame(self.root)
        control_frame.pack(fill="x", padx=10, pady=5)
        
        # Left buttons
        left_buttons = ttk.Frame(control_frame)
        left_buttons.pack(side="left")
        
        ttk.Button(left_buttons, text="🔄 Auto Map", command=self.auto_map_headers, width=15).pack(side="left", padx=(0, 5))
        ttk.Button(left_buttons, text="❌ Clear All", command=self.clear_all_mappings, width=15).pack(side="left", padx=(0, 5))
        
        # Right buttons
        right_buttons = ttk.Frame(control_frame)
        right_buttons.pack(side="right")
        
        # Convert button
        self.convert_button = ttk.Button(right_buttons, text="🔄 Convert to AccessDB", command=self.convert_to_access, width=22)
        self.convert_button.pack(side="right")

        self.progress_frame = ttk.Frame(control_frame)
        ttk.Label(self.progress_frame, text="Progress:", font=("Arial", 9)).pack(side="left", padx=(10, 5))
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.progress_frame, variable=self.progress_var, maximum=100, length=200, mode='determinate')
        self.progress_bar.pack(side="left", padx=(0, 5))
        
        self.progress_percent_label = ttk.Label(self.progress_frame, text="0%", width=5, font=("Arial", 9, "bold"))
        self.progress_percent_label.pack(side="left", padx=(0, 10))
        
        # Progress status label (same line as progress bar)
        self.progress_label = ttk.Label(self.progress_frame, text="⏳ Initializing conversion...", font=("Arial", 6), foreground="blue")
        self.progress_label.pack(side="left")

    def create_progress_section(self):
        """Progress section is now part of control section"""
        pass

    def create_results_section(self):
        """Create results section"""
        result_frame = ttk.LabelFrame(self.root, text="📄 Results and Summary")
        result_frame.pack(fill="both", expand=True, padx=10, pady=3)
        
        self.result_text = ScrolledText(result_frame, height=6, font=("Consolas", 8))
        self.result_text.pack(fill="both", expand=True, padx=3, pady=3)
        
        # Initial message
        welcome_text = """
        🔧 Header Selector Tool - Enhanced Version
        📋 How to use:
        1. Click "🔍 Browse" to select Excel file
        2. Select Sheet from dropdown (sheets are ordered as in Excel file)
        3. Left side: Select desired headers from Excel (dropdown for duplicate columns)
        4. Right side: View mapping with standard headers
        5. Use "🔄 Auto Map" for automatic mapping (selects first occurrence)
        6. Switch between sheets to configure mappings for each sheet
        7. Click "🔄 Convert to AccessDB" to convert all sheets with saved mappings
        💡 Features:
        - Mappings are saved per sheet automatically
        - First occurrence of duplicate columns selected by default
        - Unmapped Excel columns become TempData1, TempData2, etc.
        - All sheets converted together with individual mappings
        """
        self.result_text.insert(tk.END, welcome_text)

    def browse_file(self):
        """Open dialog to select Excel file"""
        file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        
        if file_path:
            self.excel_file = file_path
            filename = os.path.basename(file_path)
            self.file_label.config(text=f"📁 {filename}", foreground="black")
            self.load_excel_data()

    def load_excel_data(self):
        """Load data from Excel file - Added ini file OriginalHeader logging"""
        try:
            self.status_label.config(text="⏳ Loading...", foreground="blue")
            self.root.update()
            
            xls = pd.ExcelFile(self.excel_file)
            self.all_sheets_data = {}
            
            # Clear old mapping
            self.sheet_mappings = {}
            self.sheet_selected_headers = {}

            for i, sheet_name in enumerate(xls.sheet_names, 1):
                print(f"  {i}. {sheet_name}")
            
            # Load all sheets in Excel order
            for sheet_name in xls.sheet_names:
                try:
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=3)
                    headers = self.extract_headers(df)
                    self.all_sheets_data[sheet_name] = headers
                    print(f"  ✅ Sheet '{sheet_name}': {len(headers)} headers")
                    
                except Exception as e:
                    print(f"  ❌ Warning: Cannot read sheet '{sheet_name}': {str(e)}")
                    continue
            
            if not self.all_sheets_data:
                raise Exception("No sheets could be read from the Excel file")

            # Update dropdown in Excel file
            sheet_names = list(self.all_sheets_data.keys())
            self.sheet_combo['values'] = sheet_names

            # Set default to first sheet in Excel order
            if sheet_names:
                default_sheet = sheet_names[0]  # First sheet in Excel order
                self.sheet_combo.set(default_sheet)
                self.selected_sheet = default_sheet
                print(f"🎯 Set default sheet (first in Excel): {default_sheet}")
            
            # Load selected sheet
            if self.selected_sheet:
                self.on_sheet_selected()
            
            self.status_label.config(text=f"✅ Loaded {len(sheet_names)} sheets", foreground="green")
            
            messagebox.showinfo("✅ Success", 
                f"Excel file loaded successfully!\n\n"
                f"📋 Number of Sheets: {len(sheet_names)}\n"
                f"📊 Default Sheet: {default_sheet} (first in Excel order)\n"
                f"📁 File: {os.path.basename(self.excel_file)}")
            
        except Exception as e:
            error_msg = f"Unable to read Excel file\n\n❌ Error: {str(e)}"
            messagebox.showerror("❌ Error", error_msg)
            self.status_label.config(text="❌ Error loading file", foreground="red")

    def on_sheet_selected(self, event=None):
        """Called when selecting sheet from dropdown"""
        # Save mapping of the previous sheet first (if any)
        if self.selected_sheet and self.selected_sheet in self.all_sheets_data:
            self.save_current_sheet_mapping()
        
        # Switch to new sheet
        self.selected_sheet = self.sheet_combo.get()
        if self.selected_sheet and self.selected_sheet in self.all_sheets_data:
            self.current_sheet_headers = self.all_sheets_data[self.selected_sheet]
            
            # Update displays
            self.excel_info_label.config(
                text=f"📊 {self.selected_sheet} ({len(self.current_sheet_headers)} headers)"
            )
            self.create_excel_header_widgets()
            self.load_saved_sheet_mapping()         # Load saved mapping (if any)
            self.update_standard_tree()

    def save_current_sheet_mapping(self):
        """Save mapping of current sheet"""
        if not self.selected_sheet: return
        
        # Store selected headers of this sheet
        selected_headers = self.get_selected_excel_headers()
        self.sheet_selected_headers[self.selected_sheet] = selected_headers
        
        # Store mapping with standard headers
        final_mappings = self.get_final_mappings()
        self.sheet_mappings[self.selected_sheet] = final_mappings
        print(f"💾 Saved mapping for sheet '{self.selected_sheet}': {len(selected_headers)} headers")

    def load_saved_sheet_mapping(self):
        """Load saved mapping for this sheet"""
        if not self.selected_sheet: return
        
        if self.selected_sheet in self.sheet_selected_headers:
            saved_headers = self.sheet_selected_headers[self.selected_sheet]
            self.restore_header_selections(saved_headers)   # Restore header selection
        else:
            self.auto_select_best_options()                 # If no saved data, perform auto-mapping

    def restore_header_selections(self, saved_headers):
        """Restore header selection from saved data"""
        for key, widget_info in self.excel_header_widgets.items():
            if widget_info['type'] == 'multiple':
                # Find what was selected in saved_headers
                for saved_header in saved_headers:
                    if saved_header in widget_info['headers']:
                        widget_info['variable'].set(saved_header)
                        widget_info['selected'] = saved_header
                        widget_info['status'].config(text="Restored", foreground="green")
                        print(f"  🔄 Restored: {key} -> {saved_header}")
                        break

    def auto_map_headers(self):
        """Auto-map similar headers"""
        if not self.current_sheet_headers:
            messagebox.showwarning("⚠️ Warning", "Please select a sheet first")
            return
        
        self.auto_select_best_options()     # Auto-select best options for dropdowns
        self.update_standard_tree()         # Update mapping display
        self.save_current_sheet_mapping()   # Save mapping after auto-map
        
        # Count mappings
        available_headers = self.get_selected_excel_headers()
        current_standards = self.get_current_standard_headers()
        mapped_count = 0
        for standard_header in current_standards:
            mapped_to, _ = self.find_mapped_excel_header(standard_header, available_headers)
            if mapped_to:
                mapped_count += 1
        
        sheet_type = "List of Nominal Wall Thickness" if "List of Nominal Wall Thickness" in self.selected_sheet else "Standard"
        messagebox.showinfo("🔄 Auto Mapping Complete", 
            f"Auto-mapped {mapped_count}/{len(current_standards)} headers\n\n"
            f"📊 Sheet: {self.selected_sheet}\n"
            f"📋 Header Type: {sheet_type}\n"
            f"💾 Mapping saved for this sheet\n"
            f"💡 Please review the mapping results")

    def auto_select_best_options(self):
        """Automatically select the first occurrence for duplicate headers"""
        for key, widget_info in self.excel_header_widgets.items():
            if widget_info['type'] == 'multiple':
                headers         = widget_info['headers']
                first_header    = headers[0]  # First in list = first that appears in Excel
                
                widget_info['variable'].set(first_header)
                widget_info['selected'] = first_header
                widget_info['status'].config(text="First", foreground="blue")
                print(f"Auto-selected first occurrence: {first_header} from {headers}")

    def select_best_header(self, headers):
        """Select the first header from the list (first occurrence in Excel)"""
        return headers[0] if headers else None

    def clear_all_mappings(self):
        """Clear all mappings"""
        # Reset to default selections
        for key, widget_info in self.excel_header_widgets.items():
            if widget_info['type'] == 'multiple':
                widget_info['variable'].set(widget_info['headers'][0])
                widget_info['selected'] = widget_info['headers'][0]
                widget_info['status'].config(text="Dropdown", foreground="orange")
        
        self.update_standard_tree()
        self.save_current_sheet_mapping()       # Save mapping after clear
        
        messagebox.showinfo("❌ Complete", "Reset all selections to default")

    def show_mapping_results(self):
        """Show mapping results"""
        if not self.selected_sheet:
            messagebox.showwarning("⚠️ Warning", "Please select a sheet first")
            return
        
        self.result_text.delete(1.0, tk.END)
        report = self.generate_mapping_report()
        self.result_text.insert(tk.END, report)

    def generate_mapping_report(self):
        """Generate mapping report for all sheets"""
        # Save mapping of current sheet first
        if self.selected_sheet:
            self.save_current_sheet_mapping()
        
        report = f"""
        {'='*80}
        📊 Header Mapping Report - All Sheets
        {'='*80}

        📁 File: {os.path.basename(self.excel_file)}
        📋 Total Sheets: {len(self.all_sheets_data)}
        📊 Current Sheet: {self.selected_sheet}
        """
        
        # Display status of all sheets
        total_configured_sheets = 0
        total_mapped_headers = 0
        total_standard_headers = 0
        
        for sheet_name in self.all_sheets_data.keys():
            sheet_type = "List of Nominal Wall Thickness" if "List of Nominal Wall Thickness" in sheet_name else "Standard"
            standard_headers = self.get_sheet_specific_headers(sheet_name)
            total_standard_headers += len(standard_headers)
            
            if sheet_name in self.sheet_mappings:
                # Has saved mapping
                mappings        = self.sheet_mappings[sheet_name]
                mapped_count    = sum(1 for v in mappings.values() if v is not None)
                total_mapped_headers += mapped_count
                total_configured_sheets += 1
                
                status = f"✅ Configured ({mapped_count}/{len(standard_headers)} mapped)"
            else:
                status = "⚠️ Not configured yet"
            
            report += f"📋 {sheet_name} ({sheet_type}): {status}\n"
        
        report += f"""
        {'─'*80}
        📈 Overall Summary:
        {'─'*80}
        ├─ 📊 Total Sheets: {len(self.all_sheets_data)}
        ├─ ✅ Configured Sheets: {total_configured_sheets}
        ├─ ⚠️ Unconfigured Sheets: {len(self.all_sheets_data) - total_configured_sheets}
        ├─ 📋 Total Standard Headers: {total_standard_headers}
        └─ ✅ Total Mapped Headers: {total_mapped_headers}
        """
        return report

    def convert_to_access(self):
        """Convert to Access Database with current mappings"""
        if not self.excel_file:
            messagebox.showwarning("⚠️ Warning", "Please select Excel file first")
            return
        
        if not self.conversion_available:
            messagebox.showerror("❌ Error", 
                "Excel to Access conversion not available.\n"
                "Please ensure ExcelToAccessDB.py is in the same directory.")
            return
        
        # Save mapping of current sheet first
        if self.selected_sheet:
            self.save_current_sheet_mapping()
        
        # Count configured sheets
        configured_count   = len(self.sheet_mappings)
        unconfigured_count = len(self.all_sheets_data) - configured_count
        
        # Confirm conversion
        result = messagebox.askyesno(
            "🔄 Convert to Access Database",
            f"Convert Excel file to Access Database?\n\n"
            f"📁 Source: {os.path.basename(self.excel_file)}\n"
            f"📊 Total Sheets: {len(self.all_sheets_data)}\n"
            f"✅ Configured Sheets: {configured_count}\n"
            f"📋 Unconfigured Sheets: {unconfigured_count}\n\n"
        )
        
        if result:
            self.start_conversion()

    def start_conversion(self):
        """Start the conversion process with embedded progress bar"""

        self.conversion_cancelled = False               # Reset cancellation flag
        self.show_progress_section()                    # Show progress section
        self.convert_button.config(state="disabled")    # Disable convert button
        
        # Start conversion in separate thread
        self.conversion_thread = threading.Thread(target=self.conversion_worker, daemon=True)
        self.conversion_thread.start()

    def show_progress_section(self):
        """Show the embedded progress section centered between left and right buttons"""
        self.progress_frame.pack(side="left", expand=True, padx=(10, 10))
        
        # Initialize progress
        self.progress_var.set(0)
        self.progress_percent_label.config(text="0%")
        self.progress_label.config(text="⏳ Initializing conversion...")

    def hide_progress_section(self):
        """Hide the embedded progress section"""
        self.progress_frame.pack_forget()
        self.convert_button.config(state="normal")

    def update_progress(self, percent, message, details=None):
        """Update progress bar and message"""
        if not self.conversion_cancelled:
            self.progress_var.set(percent)
            self.progress_percent_label.config(text=f"{percent:.1f}%")
            self.progress_label.config(text=message)
            self.root.update_idletasks()

    def conversion_worker(self):
        """Worker thread for conversion process"""
        try:
            # Save mapping of current sheet first
            if self.selected_sheet:
                self.save_current_sheet_mapping()
            
            # Step 1: Initialize (5%)
            self.root.after(0, self.update_progress, 5, "🚀 Initializing conversion process...")
            if self.conversion_cancelled: return
            time.sleep(0.5)
            
            # Step 2: Prepare mappings (10%)
            self.root.after(0, self.update_progress, 10, "📋 Preparing sheet mappings...")
            
            # Create sheet_modes and sheet_mappings for all sheets
            sheet_modes = {}
            all_sheet_mappings = {}
            configured_count = 0
            unconfigured_count = 0
            
            for i, sheet_name in enumerate(self.all_sheets_data.keys()):
                if self.conversion_cancelled: return
                
                if sheet_name in self.sheet_mappings:
                    sheet_modes[sheet_name] = "standard"        # Has saved mapping -> use standard mode
                    all_sheet_mappings[sheet_name] = self.sheet_mappings[sheet_name]
                    configured_count += 1
                else:
                    sheet_modes[sheet_name] = "individual"      # No mapping -> use individual mode (original headers)
                    unconfigured_count += 1
                
                progress = 10 + (i * 10 / len(self.all_sheets_data))
                self.root.after(0, self.update_progress, progress, f"📋 Processing sheet mappings... ({i+1}/{len(self.all_sheets_data)})")
            
            # Step 3: Create enhanced sheet modes (25%)
            self.root.after(0, self.update_progress, 25, "⚙️ Creating enhanced sheet configurations...")
            if self.conversion_cancelled: return
            time.sleep(0.3)
            
            enhanced_sheet_modes = {}
            for sheet_name, mode in sheet_modes.items():
                if mode == "standard" and sheet_name in all_sheet_mappings:
                    enhanced_sheet_modes[sheet_name] = {
                        'mode': 'standard',
                        'mappings': all_sheet_mappings[sheet_name],
                        'is_nominal_wall': "List of Nominal Wall Thickness" in sheet_name
                    }
                else:
                    enhanced_sheet_modes[sheet_name] = {'mode': mode}
            
            # Step 4: Generate header mapping content (35%)
            self.root.after(0, self.update_progress, 35, "📝 Generating header mapping content...")
            if self.conversion_cancelled: return
            
            header_mapping_content = self.create_header_mapping_content()
            
            # Step 5: Prepare primary mapping (40%)
            primary_mapping = None
            primary_sheet = None
            for sheet_name, mapping in all_sheet_mappings.items():
                if mapping:
                    primary_mapping = list(mapping.keys())
                    primary_sheet = sheet_name
                    break
            
            if not primary_mapping:
                primary_mapping = self.Set_standard_headers()
            
            self.root.after(0, self.update_progress, 40, "🎯 Primary mapping prepared...")
            if self.conversion_cancelled: return
            
            # Step 6: Start actual conversion (50%-90%)
            conversion_steps = [
                (50, "📖 Reading Excel file structure..."),
                (55, "🔍 Analyzing sheet data..."),
                (60, "🏗️ Creating Access database structure..."),
                (65, "📋 Processing sheet configurations..."),
                (70, "🗂️ Creating database tables..."),
                (75, "📊 Converting sheet data..."),
                (80, "🔗 Applying header mappings..."),
                (85, "💾 Writing data to Access database..."),
                (90, "📄 Adding header mapping documentation...")
            ]
            
            for percent, message in conversion_steps:
                if self.conversion_cancelled: return
                self.root.after(0, self.update_progress, percent, message)
                time.sleep(0.3)
            
            # Step 7: Call actual conversion function (95%)
            self.root.after(0, self.update_progress, 95, "⚡ Executing conversion")
            if self.conversion_cancelled: return
            
            # Call the actual conversion function
            result = self.excel_to_access(
                self.excel_file,
                header_file=None,
                selected_headers=primary_mapping,
                sheet_modes=enhanced_sheet_modes,
                header_mapping_content=header_mapping_content
            )
            
            if self.conversion_cancelled: return
            
            # Step 8: Finalization (100%)
            if result:
                access_file = os.path.splitext(self.excel_file)[0] + ".accdb"
                self.root.after(0, self.update_progress, 100, "🎉 Conversion completed successfully!")
                self.root.after(0, self.conversion_success, access_file, configured_count, unconfigured_count)
            else:
                self.root.after(0, self.update_progress, 100, "❌ Conversion failed")
                self.root.after(0, self.conversion_failed, "Conversion function returned False")
                
        except Exception as e:
            if not self.conversion_cancelled:
                error_msg = str(e)
                self.root.after(0, self.update_progress, 100, "❌ Conversion error occurred")
                self.root.after(0, self.conversion_failed, error_msg)
        
        finally:
            # Always hide progress section when done
            self.root.after(0, self.finalize_conversion)

    def finalize_conversion(self):
        """Finalize conversion process"""
        time.sleep(1)  # Show final status for a moment
        self.hide_progress_section()

    def conversion_success(self, access_file, configured_count, unconfigured_count):
        """Handle successful conversion"""
        # Update result text
        self.result_text.delete(1.0, tk.END)
        
        success_report = f"""
        {'='*60}
        ✅ CONVERSION COMPLETED SUCCESSFULLY!
        {'='*60}

        📁 Source: {os.path.basename(self.excel_file)}
        💾 Output: {os.path.basename(access_file)}
        📊 Total Sheets Converted: {len(self.all_sheets_data)}
        ✅ Custom Mappings Applied: {configured_count} sheets
        📋 Original Headers Used: {unconfigured_count} sheets
        📄 HeaderMapping: Added to Access DB

        🎉 Conversion finished successfully!
        All data has been converted and saved to the Access database.
        """
        
        self.result_text.insert(tk.END, success_report)
        
        # Show success message
        if not self.conversion_cancelled:
            messagebox.showinfo("✅ Conversion Successful", 
                f"Conversion completed successfully!\n\n"
                f"Output: {os.path.basename(access_file)}\n"
                f"Total sheets: {len(self.all_sheets_data)}\n"
                f"Custom mappings: {configured_count} sheets\n"
                f"Original headers: {unconfigured_count} sheets\n"
                f"HeaderMapping: Added to Access DB")

    def conversion_failed(self, error_msg, tb=None):
        """Handle failed conversion"""
        # Update result text
        self.result_text.delete(1.0, tk.END)
        
        error_report = f"""
        ❌ CONVERSION FAILED
        {'='*60}

        Error: {error_msg}

        {tb if tb else ''}
        """
        
        self.result_text.insert(tk.END, error_report)
        
        # Show error message
        if not self.conversion_cancelled:
            messagebox.showerror("❌ Conversion Error", f"Conversion failed:\n{error_msg}")

    def get_final_mappings(self):
        """Get final mappings for conversion"""
        mappings = {}
        selected_excel_headers = self.get_selected_excel_headers()
        current_standards = self.get_current_standard_headers()
        
        for standard_header in current_standards:
            mapped_excel, _ = self.find_mapped_excel_header(standard_header, selected_excel_headers)
            mappings[standard_header] = mapped_excel
        
        return mappings

    def create_header_mapping_content(self):
        """Generate HeaderMapping content without adding to Access DB"""
        if not self.excel_file:
            return None

        # Save the mapping of the current sheet first
        if self.selected_sheet:
            self.save_current_sheet_mapping()

        try:
            content = f"""Header Mapping Report
    {'='*80}
    📁 Source File: {os.path.basename(self.excel_file)}
    📊 Total Sheets: {len(self.all_sheets_data)}
    📅 Created: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    {'='*80}

    """
        
            # Loop through each sheet
            for sheet_name in self.all_sheets_data.keys():
                content += f"\n📋 SHEET: {sheet_name}\n"
                content += f"{'─'*60}\n"
                sheet_type = "List of Nominal Wall Thickness" if "List of Nominal Wall Thickness" in sheet_name else "Standard"
                content += f"Sheet Type: {sheet_type}\n"
           
                standard_headers = self.get_sheet_specific_headers(sheet_name)
            
                if sheet_name in self.sheet_mappings:
                    mappings = self.sheet_mappings[sheet_name]
                    content += f"Standard Headers: {len(standard_headers)}\n"
                    content += f"Header Mappings:\n\n"
                
                    mapped_count = 0
                    unmapped_count = 0
                    tempdata_count = 0

                    content += f"{'Standard Header':<50} >> {'Mapping Column'}\n"
                    content += f"{'-'*50} >> {'-'*50}\n"
                
                    for i, standard_header in enumerate(standard_headers, 1):
                        mapped_to = mappings.get(standard_header)
                        if mapped_to:
                            content += f"{standard_header:<50} >> {mapped_to}\n"
                            mapped_count += 1
                        else:
                            content += f"{standard_header:<50} >> [Empty - Will create empty column]\n"
                            unmapped_count += 1
                    
                    # Find the remaining headers from Excel that are not used
                    used_headers = [v for v in mappings.values() if v is not None]
                    all_excel_headers = self.all_sheets_data[sheet_name]
                    remaining_headers = [h for h in all_excel_headers if h not in used_headers]
                
                    if remaining_headers:
                        content += f"\n📁 REMAINING COLUMNS (will become TempData):\n"
                        for i, orig_col in enumerate(remaining_headers, 1):
                            content += f"TempData{i:<3d} << {orig_col}\n"
                            tempdata_count += 1
                    else:
                        content += f"\n📋 No remaining columns - no TempData will be created\n"
                
                    # Show new columns for List of Pipe Tally
                    if sheet_name == "List of Pipe Tally":
                        content += f"\n🆕 NEW COLUMNS:\n"
                        new_columns = ['Velocity (m/s)', 'ImgPath1', 'ImgPath2', 'Timestr']
                        for new_col in new_columns:
                            content += f"{new_col}\n"
                            
                    # Summary for this sheet
                    content += f"\n📊 Summary: {mapped_count} Mapped, {unmapped_count} Unmapped, {tempdata_count} TempData\n"
                            
                else:
                    # Sheet without mapping
                    content += f"Status: ⚠️ Not Configured (Original Headers)\n"
                    available_headers = self.all_sheets_data[sheet_name]
                    content += f"Will use original Excel headers ({len(available_headers)} columns):\n\n"
                
                    for i, header in enumerate(available_headers, 1):
                        content += f"  {i:2d}. {header}\n"
        
            # Overall summary
            configured_count = len(self.sheet_mappings)
            unconfigured_count = len(self.all_sheets_data) - configured_count
        
            content += f"\n{'='*80}\n"
            content += f"📈 OVERALL SUMMARY:\n"
            content += f"{'='*80}\n"
            content += f"📊 Total Sheets: {len(self.all_sheets_data)}\n"
            content += f"✅ Configured Sheets: {configured_count}\n"
            content += f"📋 Unconfigured Sheets: {unconfigured_count}\n"
        
            if configured_count > 0:
                total_mapped = 0
                total_standards = 0
                total_tempdata = 0
            
                for sheet_name, mappings in self.sheet_mappings.items():
                    total_mapped += sum(1 for v in mappings.values() if v is not None)
                    total_standards += len(mappings)
                
                    used_headers = [v for v in mappings.values() if v is not None]
                    all_excel_headers = self.all_sheets_data[sheet_name]
                    remaining_headers = [h for h in all_excel_headers if h not in used_headers]
                    total_tempdata += len(remaining_headers)
            
                content += f"📋 Total Standard Headers: {total_standards}\n"
                content += f"✅ Total Mapped Headers: {total_mapped}\n"
                content += f"📦 Total TempData Columns: {total_tempdata}\n"
            
            content += f"\n{'='*80}\n"
            content += f"🔧 Generated by Header Selector Tool\n"
            content += f"📅 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            content += f"{'='*80}\n"
            
            return content
            
        except Exception as e:
            print(f"❌ Error creating HeaderMapping content: {str(e)}")
            return None

def main():
    """Main function to run the program"""
    root = tk.Tk()
    
    # Set icon if available
    try:
        pass
    except:
        pass
    
    app = HeaderSelector(root)
    
    # Center window on screen
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main()