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
import subprocess
import platform
import glob
import configparser

class HeaderSelector:
    # Feature type configuration - integrated directly into the class
    FEATURE_TYPE_CONFIG = {
        'height_features': ['CRAL'],
        'depth_features': ['COCL', 'CORR', 'LAMI', 'MAIN'],
    }
    
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
        self.excel_header_widgets = {}                          # Store widgets for excel headers
        self.standard_mapping_status = {}                       # Store mapping status of standard headers
        
        self.sheet_mappings = {}                               # Store mapping of all sheets
        self.sheet_selected_headers = {}                       # Store selected headers of all sheets
       
        self.use_imperial_units = tk.BooleanVar(value=False)   # Unit system variables (Default to SI units)
        self.standard_headers = self.Set_standard_headers()    # Initialize standard headers after unit system variable is created
        
        # Mapping file variables
        self.mapping_file_path = ""                            # Path to mapping file
        self.mapping_loaded = False                            # Flag if mapping was loaded
        
        # Progress tracking variables
        self.conversion_cancelled       = False
        self.progress_var               = None
        self.progress_label             = None
        self.progress_bar               = None
        self.convert_button             = None
        
        # Database file tracking
        self.last_converted_db_path = ""                       # Store last converted database path
        self.opening_location = False                          # Flag to prevent multiple opens
        
        # Import conversion functions
        try:
            from ExcelToAccessDB import excel_to_access
            self.excel_to_access = excel_to_access
            self.conversion_available = True
        except ImportError:
            self.conversion_available = False
            print("⚠️ Excel to Access conversion: Not available")

        self.setup_ui()                             # Create UI
        self.setup_styles()                         # Set style
        self.check_and_auto_load_recent_file()      # Check for recently opened Excel file and try to auto-load

    def open_database_location(self, db_path):
        """Open the folder containing the database file"""
        
        # Prevent multiple simultaneous opens
        if hasattr(self, 'opening_location') and self.opening_location:
            return False
            
        self.opening_location = True
        
        try:
            if not os.path.exists(db_path):
                messagebox.showerror("❌ Error", f"Database file not found:\n{db_path}")
                return False
            
            # Get directory path
            dir_path = os.path.dirname(os.path.abspath(db_path))
            os.startfile(dir_path)
            
            return True
            
        except Exception as e:
            messagebox.showerror("❌ Error", f"Cannot open database location:\n{str(e)}")
            return False
            
        finally:
            self.opening_location = False

    def open_database_in_windows_explorer(self, db_path):
        """Redirect to single method"""
        return self.open_database_location(db_path)

    def show_conversion_complete_dialog(self, access_file, configured_count, unconfigured_count):
        """Show completion dialog with option to open database location"""
    
        # Create custom dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("🎉 Conversion Complete")
        dialog.geometry("420x420")
        dialog.resizable(False, False)
        dialog.withdraw()

        # Center the dialog
        dialog.transient(self.root)
    
        # Main container with absolute positioning
        container = tk.Frame(dialog)
        container.pack(fill="both", expand=True)
        container.pack_propagate(False)
    
        # Title section
        title_frame = tk.Frame(container, height=50)
        title_frame.pack(fill="x", padx=10, pady=(10, 5))
        title_frame.pack_propagate(False)
    
        style = ttk.Style()
        style.configure("Small.TLabelframe.Label", font=("Arial", 7))

        success_label = ttk.Label(title_frame, text="🎉", font=("Arial", 11))
        success_label.pack(pady=(5, 0))
    
        title_label = ttk.Label(title_frame, text="Conversion Completed Successfully!", font=("Arial", 7, "bold"), foreground="green")
        title_label.pack()

        # Information section
        info_frame = ttk.LabelFrame(container, text="📊 Conversion Summary", padding="6", style="Small.TLabelframe")
        info_frame.pack(fill="x", padx=10, pady=5)
        info_frame.pack_propagate(False)
        info_frame.configure(height=150)
    
        # Configure grid
        info_frame.grid_columnconfigure(1, weight=1)
    
        # Labels with consistent positioning
        ttk.Label(info_frame, text="📂   Source File:", font=("Arial", 7, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 8), pady=2)
        source_label = ttk.Label(info_frame, text=os.path.basename(self.excel_file), font=("Arial", 7), wraplength=250)
        source_label.grid(row=0, column=1, sticky="w", pady=2)

        ttk.Label(info_frame, text="📋   Database File:", font=("Arial", 7, "bold")).grid(row=1, column=0, sticky="w", padx=(0, 8), pady=2)
        db_label = ttk.Label(info_frame, text=os.path.basename(access_file), font=("Arial", 7), wraplength=250)
        db_label.grid(row=1, column=1, sticky="w", pady=2)
    
        ttk.Label(info_frame, text="📊   Total Sheets:", font=("Arial", 7, "bold")).grid(row=2, column=0, sticky="w", padx=(0, 8), pady=2)
        ttk.Label(info_frame, text=str(len(self.all_sheets_data)), font=("Arial", 7)).grid(row=2, column=1, sticky="w", pady=2)
    
        ttk.Label(info_frame, text="✅   Custom Mappings:", font=("Arial", 7, "bold")).grid(row=3, column=0, sticky="w", padx=(0, 8), pady=2)
        ttk.Label(info_frame, text=f"{configured_count} sheets", font=("Arial", 7)).grid(row=3, column=1, sticky="w", pady=2)
    
        ttk.Label(info_frame, text="📋   Original Headers:", font=("Arial", 7, "bold")).grid(row=4, column=0, sticky="w", padx=(0, 8), pady=2)
        ttk.Label(info_frame, text=f"{unconfigured_count} sheets", font=("Arial", 7)).grid(row=4, column=1, sticky="w", pady=2)
   
        ttk.Label(info_frame, text="📏    Unit System:", font=("Arial", 7, "bold")).grid(row=5, column=0, sticky="w", padx=(0, 8), pady=2)
        unit_system = "Imperial" if self.use_imperial_units.get() else "SI"
        ttk.Label(info_frame, text=unit_system, font=("Arial", 7)).grid(row=5, column=1, sticky="w", pady=2)
        
        ttk.Label(info_frame, text="📄   HeaderMapping:", font=("Arial", 7, "bold")).grid(row=6, column=0, sticky="w", padx=(0, 8), pady=2)
    
        has_mapping = len(self.sheet_mappings) > 0 or len(self.all_sheets_data) > 0    
        if has_mapping:
            ttk.Label(info_frame, text="Added to Access DB", font=("Arial", 7)).grid(row=6, column=1, sticky="w", pady=2)
        else:
            ttk.Label(info_frame, text="Not created", font=("Arial", 7)).grid(row=6, column=1, sticky="w", pady=2)

        # Database Location section
        path_frame = ttk.LabelFrame(container, text="📂 Database Location", padding="6", style="Small.TLabelframe")
        path_frame.pack(fill="x", padx=10, pady=5)
        path_frame.pack_propagate(False)
        path_frame.configure(height=70)

        path_text = tk.Text(path_frame, height=2, wrap=tk.WORD, font=("Consolas", 7), relief="sunken", borderwidth=1, background="#f8f9fa")
        path_text.pack(fill="x", pady=(2, 0))
        path_text.insert("1.0", access_file)
        path_text.config(state="disabled")
   
        # Button section
        button_frame = tk.Frame(container, height=50)
        button_frame.pack(fill="x", side="bottom", padx=10, pady=(5, 10))
        button_frame.pack_propagate(False)

        # Variables to track user choice
        dialog_closed   = threading.Event()
        user_action     = [None]    # Use list to make it mutable in nested function
        action_taken    = [False]   # Prevent multiple actions
    
        def on_open_location():
            if action_taken[0]:  # Prevent multiple clicks
                return
            action_taken[0] = True
            user_action[0]  = 'open_location'
            dialog_closed.set()
            dialog.destroy()
    
        def on_close():
            if action_taken[0]:  # Prevent multiple clicks  
                return
            action_taken[0] = True
            user_action[0]  = 'close'
            dialog_closed.set()
            dialog.destroy()
    
        style.configure("Small.TButton", font=("Arial", 7))

        open_btn = ttk.Button(button_frame, text="📂 Open Location", command=on_open_location, width=18, style="Small.TButton")
        open_btn.pack(side="left", padx=(0, 8))
    
        close_btn = ttk.Button(button_frame, text="❌ Close", command=on_close, width=12, style="Small.TButton")
        close_btn.pack(side="right")
    
        open_btn.focus_set()        # Set focus to Open button
    
        # Add keyboard shortcuts with action protection
        def on_key_press(event):
            if action_taken[0]:     # Prevent multiple keyboard actions
                return
            if event.keysym == 'Return' or event.keysym == 'KP_Enter':
                on_open_location()
            elif event.keysym == 'Escape':
                on_close()
    
        dialog.bind('<Key>', on_key_press)
        dialog.focus_set()
    
        # Handle window close event with action protection
        def on_window_close():
            if not action_taken[0]:
                on_close()
    
        dialog.protocol("WM_DELETE_WINDOW", on_window_close)
    
        # Center dialog on screen with ABSOLUTE positioning
        screen_width  = dialog.winfo_screenwidth()
        screen_height = dialog.winfo_screenheight()
        x = (screen_width // 2) - (420 // 2)
        y = (screen_height // 2) - (420 // 2)
        dialog.geometry(f"420x420+{x}+{y}")
    
        dialog.deiconify()
        dialog.grab_set()

        dialog.wait_window()        # Wait for user choice
    
        # Handle user action
        if user_action[0] == 'open_location':
            success = self.open_database_location(access_file)
        
            if success:
                print("✅ Database location opened successfully")
            else:
                print("❌ Failed to open database location")
        else:
            print("ℹ️ User chose to close dialog without opening location")

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
        """Set standard columns based on unit system"""
        use_imperial = self.use_imperial_units.get()
        
        if use_imperial:
            # Imperial units
            return [
                "Log distance (ft)",
                "Latitude (degree)",
                "Longitude (degree)",
                "Altitude (ft)",
                "Feature type",
                "Feature identification",
                "Anomaly identification",
                "Girth weld Nr",
                "Joint manufacturing type",
                "Joint / component length (ft)",
                "Nominal internal diameter (in)",
                "Nominal thickness (in)",
                "Measure/Reference thickness (in)",
                "Abs Dist to upstream weld (ft)",
                "Clock position seam / anomaly",
                "Surface location",
                "Remaining thickness (in)",
                "Max. depth (in)",
                "Max. depth (%)",
                "Max. Height (in)",
                "Max. Height (%)",
                "Length (in)",
                "Width (in)",
                "Metal loss anomaly dimension classification",
                "ERF",
                "Comments"
            ]
        else:
            # SI units
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
                "Max. depth (mm)",
                "Max. depth (%)",
                "Length (mm)",
                "Width (mm)",
                "Metal loss anomaly dimension classification",
                "ERF",
                "Comments"
            ]

    def get_sheet_specific_headers(self, sheet_name):
        """Get headers specific to each sheet type based on unit system"""
        use_imperial = self.use_imperial_units.get()
        
        if "List of Nominal Wall Thickness" in sheet_name:
            if use_imperial:
                return [
                    "Log distance (ft)",
                    "Girth weld Nr",
                    "Nominal thickness (in)",
                    "Joint manufacturing type",
                    "SMYS (psi)",
                    "Design Pressure (psi)",
                    "MAOP (psi)"
                ]
            else:
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
            # Default: return all standard headers
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
        
        # Buttons container for Browse and Helper
        buttons_frame = ttk.Frame(inner_frame)
        buttons_frame.pack(side="right")
    
        ttk.Button(buttons_frame, text="🔍 Browse", command=self.browse_file).pack(side="left")
        ttk.Button(buttons_frame, text="❓", command=self.show_helper, width=3).pack(side="left")

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

        # Unit system selection row
        unit_row = ttk.Frame(file_frame)
        unit_row.pack(fill="x", padx=10, pady=(0, 3))
        
        ttk.Label(unit_row, text="Units:").pack(side="left")
        
        # Unit system radio buttons
        unit_frame = ttk.Frame(unit_row)
        unit_frame.pack(side="left", padx=(10, 0))
        
        ttk.Radiobutton(unit_frame, text="SI Units", variable=self.use_imperial_units, value=False).pack(side="left", padx=(0, 10))
        ttk.Radiobutton(unit_frame, text="Imperial Units", variable=self.use_imperial_units, value=True).pack(side="left")
        
        # Unit status label
        self.unit_status_label = ttk.Label(unit_row, text="📏 SI Units selected", foreground="blue", font=("Arial", 8))
        self.unit_status_label.pack(side="right")
        
        # Bind unit change event
        self.use_imperial_units.trace('w', self.on_unit_system_changed)

        # Mapping file section
        mapping_row = ttk.Frame(file_frame)
        mapping_row.pack(fill="x", padx=10, pady=(0, 3))
        
        self.mapping_status_label = ttk.Label(mapping_row, text="💾 Mapping: Not loaded", foreground="gray", font=("Arial", 8))
        self.mapping_status_label.pack(side="left")

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
      
        for i, header in enumerate(headers):
            group_key = self.create_grouping_key(header)
            groups[group_key].append(header)
            
            if group_key not in header_order:
                header_order[group_key] = i
            
            print(f"  {i+1:2d}. {header} -> :{group_key}")
        
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

                self.update_standard_tree()                 # Update mapping display
                if self.selected_sheet and self.excel_header_widgets:
                    self.save_current_sheet_mapping()       # Save mapping immediately

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
        """Get current standard headers based on sheet type"""
        if self.selected_sheet:
            base_headers     = self.get_sheet_specific_headers(self.selected_sheet)
            adjusted_headers = self.adjust_log_distance_header(base_headers)         # Check Log distance unit in Excel headers and adjust Standard Headers
            filtered_headers = self.filter_headers_by_feature_type(adjusted_headers) # Filter headers based on feature type analysis
            
            return filtered_headers
        else:
            return self.Set_standard_headers()

    def filter_headers_by_feature_type(self, headers, sheet_name=None):
        """Filter headers based on feature type analysis"""
        if not sheet_name:
            sheet_name = self.selected_sheet
            
        if "List of Nominal Wall Thickness" in sheet_name:
            return headers

        use_imperial = self.use_imperial_units.get()        # Check unit system first
        
        if not use_imperial:
            # SI Units: Always show Max. depth columns
            filtered_headers = []
            max_headers_found = []
       
            for header in headers:
                if not any(max_pattern in header.lower() for max_pattern in ['max. height', 'max. depth']):
                    filtered_headers.append(header)
                else:
                    if 'max. depth' in header.lower():
                        filtered_headers.append(header)
   
            return filtered_headers
        
        else:
            # Imperial Units: Use feature type analysis to determine which Max columns to show
            if "List of Pipe Tally" in sheet_name:
                feature_analysis = self.analyze_feature_types_in_sheet(sheet_name)  # Analyze Feature types in Pipe Tally sheet
                if feature_analysis:
                    self.pipe_tally_feature_analysis = feature_analysis             # Store the analysis result for use in other sheets
            else:
                # Use stored analysis from Pipe Tally if available
                if hasattr(self, 'pipe_tally_feature_analysis') and self.pipe_tally_feature_analysis:
                    feature_analysis = self.pipe_tally_feature_analysis
                else:
                    # Fallback: analyze current sheet if no Pipe Tally analysis available
                    feature_analysis = self.analyze_feature_types_in_sheet(sheet_name)
            
            if feature_analysis:
                # Filter Max columns based on detected feature Identification
                filtered_headers = []
                max_headers_found = []
                
                for header in headers:
                    # Always include non-Max columns
                    if not any(max_pattern in header.lower() for max_pattern in ['max. height', 'max. depth']):
                        filtered_headers.append(header)
                    else:
                        # max_headers_found.append(header)
                        if 'max. height' in header.lower():
                            if feature_analysis['show_max_height']:
                                filtered_headers.append(header)
                        elif 'max. depth' in header.lower():
                            if feature_analysis['show_max_depth']:
                                filtered_headers.append(header)

                # Force show only one type if both are detected
                if feature_analysis['show_max_height'] and feature_analysis['show_max_depth']:
                    # Remove Max. depth columns from filtered_headers
                    filtered_headers = [h for h in filtered_headers if 'max. depth' not in h.lower()]
                
                return filtered_headers
            else:
                # Fallback: If feature analysis fails, show Max. Height columns by default
                filtered_headers = []
                # max_headers_found = []

                for header in headers:
                    if not any(max_pattern in header.lower() for max_pattern in ['max. height', 'max. depth']):
                        filtered_headers.append(header)
                    else:
                        # In fallback mode, show Max. Height columns if they exist
                        if 'max. height' in header.lower():
                            filtered_headers.append(header)

                return filtered_headers

    def update_standard_tree(self):
        """Update standard tree according to current mapping"""
        # Clear existing items
        for item in self.standard_tree.get_children():
            self.standard_tree.delete(item)
        
        # Get current headers and available excel headers
        current_headers         = self.get_current_standard_headers()
        available_excel_headers = self.get_selected_excel_headers()
        
        # Update info label
        sheet_type   = "List of Nominal Wall Thickness" if "List of Nominal Wall Thickness" in self.selected_sheet else "Standard"
        header_count = len(current_headers)
        unit_system  = "Imperial" if self.use_imperial_units.get() else "SI"
        
        # Get feature type analysis for additional info
        feature_info = ""
        if self.selected_sheet and "List of Nominal Wall Thickness" not in self.selected_sheet:
            feature_analysis = self.analyze_feature_types_in_sheet(self.selected_sheet)
            if feature_analysis:
                if feature_analysis['show_max_height'] and feature_analysis['show_max_depth']:
                    feature_info = f" | Max columns: Height & Depth"
                elif feature_analysis['show_max_height']:
                    feature_info = f" | Max columns: Height only (CRACK/CRAL)"
                elif feature_analysis['show_max_depth']:
                    feature_info = f" | Max columns: Depth only (COCL/CORR/LAMI)"
                else:
                    feature_info = f" | Max columns: None"
        
        if sheet_type == "List of Nominal Wall Thickness":
            self.standard_info_label.config(text=f"Standard Headers for {sheet_type} ({header_count} Columns, {unit_system} Units)")
        else:
            self.standard_info_label.config(text=f"Standard Headers ({header_count} Columns, {unit_system} Units){feature_info}")
        
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
            selected_header = widget_info['selected']
            selected_headers.append(selected_header)
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

        # Special mapping for Log distance in Imperial Units
        if standard_header == "Log distance (ft)" and self.use_imperial_units.get():
            for excel_header in excel_headers:
                excel_lower = excel_header.lower()
                if ('log distance' in excel_lower and 
                    ('[ft]' in excel_header or '[mi]' in excel_header or '(ft)' in excel_header or '(mi)' in excel_header)):
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

    def generate_mapping_file_path(self):
        """Generate mapping file path based on Excel file"""
        if not self.excel_file:
            return ""
        
        excel_dir  = os.path.dirname(self.excel_file)
        excel_name = os.path.splitext(os.path.basename(self.excel_file))[0]
        mapping_filename = f"{excel_name}.~p~"
        return os.path.join(excel_dir, mapping_filename)

    def save_mapping_to_file(self):
        """Save current mappings to INI file"""
        try:
            # Save current sheet mapping first
            if self.selected_sheet:
                self.save_current_sheet_mapping()
            
            mapping_file = self.generate_mapping_file_path()

            # Create ConfigParser object
            config = configparser.ConfigParser(interpolation=None)
            
            # Analyze Feature types to determine which type is present
            feature_mapping_type = "NONE"
            if self.all_sheets_data:
                # Analyze the first sheet or a specific sheet to determine feature type
                first_sheet = list(self.all_sheets_data.keys())[0]
                feature_analysis = self.analyze_feature_types_in_sheet(first_sheet)
                
                if feature_analysis:
                    if feature_analysis['show_max_height']:
                        feature_mapping_type = "CRACK"
                    elif feature_analysis['show_max_depth']:
                        feature_mapping_type = "CORROSION"
                    else:
                        feature_mapping_type = "NONE"
            
            # General information section
            config['GENERAL'] = {
                'excel_file': os.path.basename(self.excel_file),
                'created_date': datetime.now().strftime("%Y-%m-%d"), 
                'total_sheets': str(len(self.all_sheets_data)),
                'configured_sheets': str(len(self.sheet_mappings)),
                'Unit System': "Imperial" if self.use_imperial_units.get() else "SI",
                'Feature Iden Mapping': feature_mapping_type
            }
       
            # Save to file
            for sheet_name, mappings in self.sheet_mappings.items():
                section_name = f"MAPPING_{sheet_name}"
                config[section_name] = {}
                
                for standard_header, mapped_excel in mappings.items():
                    # Convert None values to empty strings for INI
                    value = mapped_excel if mapped_excel is not None else ""
                    config[section_name][standard_header] = value

            # Save selected headers
            for sheet_name, headers in self.sheet_selected_headers.items():
                section_name = f"SELECTED_{sheet_name}"
                config[section_name] = {}
                
                # Save headers as numbered keys
                for i, header in enumerate(headers, 1):
                    config[section_name][str(i)] = header

            # Save all sheets data
            for sheet_name, headers in self.all_sheets_data.items():
                section_name = f"SHEETDATA_{sheet_name}"
                config[section_name] = {}
                
                for i, header in enumerate(headers, 1):
                    config[section_name][str(i)] = header
                 
            # Write to INI file
            with open(mapping_file, 'w', encoding='utf-8') as f:
                config.write(f)
            
            self.mapping_file_path = mapping_file
            self.update_mapping_status(f"💾 Saved: {os.path.basename(mapping_file)}", "green")

            return True, mapping_file

        except Exception as e:
            error_msg = f"Error saving mapping: {str(e)}"
            print(f"❌ {error_msg}")
            self.update_mapping_status("💾 Save Failed", "red")
            return False, error_msg

    def load_mapping_from_file(self, mapping_file=None):
        """Load mappings from INI file"""
        try:
            if not mapping_file:
                mapping_file = self.generate_mapping_file_path()
            
            if not os.path.exists(mapping_file):
                return False, "Mapping file not found"
            
            # Load data from file
            config = configparser.ConfigParser(interpolation=None)
            config.read(mapping_file, encoding='utf-8')
            
            # Validate data
            if 'GENERAL' not in config:
                return False, "Invalid mapping file format - missing GENERAL section"
            
            # Check if Excel file matches
            loaded_excel = config['GENERAL'].get('excel_file', '')
            current_excel = os.path.basename(self.excel_file)

            # Load Unit System from saved configuration
            saved_unit_system = config['GENERAL'].get('Unit System', 'SI')
            if saved_unit_system == 'Imperial':
                self.use_imperial_units.set(True)
                self.unit_status_label.config(text="📏 Imperial Units selected", foreground="green")
            else:
                self.use_imperial_units.set(False)
                self.unit_status_label.config(text="📏 SI Units selected", foreground="blue")
            print(f"📏 Loaded Unit System: {saved_unit_system}")
            
            # Load sheet mappings
            self.sheet_mappings = {}
            for section_name in config.sections():
                if section_name.startswith('MAPPING_'):
                    sheet_name = section_name[8:]  # Remove 'MAPPING_' prefix
                    mappings = {}
                    
                    for key, value in config[section_name].items():
                        mapped_value = value if value.strip() != "" else None
                        mappings[key] = mapped_value
                    
                    self.sheet_mappings[sheet_name] = mappings
            
            # Load selected headers
            self.sheet_selected_headers = {}
            for section_name in config.sections():
                if section_name.startswith('SELECTED_'):
                    sheet_name = section_name[9:]  # Remove 'SELECTED_' prefix
                    headers = []

                    numeric_keys = []
                    for key in config[section_name].keys():
                        if key.isdigit():
                            numeric_keys.append(int(key))
                    
                    if numeric_keys:
                        for i in sorted(numeric_keys):
                            key_str = str(i)
                            if key_str in config[section_name]:
                                headers.append(config[section_name][key_str])

                    self.sheet_selected_headers[sheet_name] = headers

            self.mapping_file_path = mapping_file
            self.mapping_loaded = True
            
            loaded_sheets = len(self.sheet_mappings)
            created_date = config['GENERAL'].get('created_date', 'Unknown')
            
            self.update_mapping_status(f"📂 Loaded: {loaded_sheets} sheets", "blue")
            
            return True, f"Loaded mappings for {loaded_sheets} sheets (Created: {created_date[:10]})"

        except Exception as e:
            error_msg = f"Error loading mapping: {str(e)}"
            print(f"❌ {error_msg}")
            self.update_mapping_status("📂 Load Failed", "red")
            return False, error_msg

    def update_mapping_status(self, message, color="black"):
        """Update mapping status label"""
        self.mapping_status_label.config(text=message, foreground=color)
    
    def on_unit_system_changed(self, *args):
        """Called when unit system is changed"""
        use_imperial = self.use_imperial_units.get()
        
        if use_imperial:
            self.unit_status_label.config(text="📏 Imperial Units selected", foreground="green")
        else:
            self.unit_status_label.config(text="📏 SI Units selected", foreground="blue")

        # Update standard headers if sheet is selected
        if self.selected_sheet:
            self.update_standard_tree()

    def auto_load_mapping_if_exists(self):
        """Automatically load mapping file if it exists"""
        mapping_file = self.generate_mapping_file_path()
        if os.path.exists(mapping_file):
            success, message = self.load_mapping_from_file(mapping_file)
            if success:
                loaded_sheets           = len(self.sheet_mappings)
                selected_headers_count  = sum(len(headers) for headers in self.sheet_selected_headers.values())
                self.update_mapping_status(f"📂 Auto-loaded: {loaded_sheets} sheets", "blue")
                return True
            else:
                self.update_mapping_status("📂 Load Failed", "red")
        else:
            self.update_mapping_status("💾 No saved mapping", "gray")
        return False

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
        self.progress_bar = ttk.Progressbar(self.progress_frame, variable=self.progress_var, maximum=100, length=250, mode='determinate')
        self.progress_bar.pack(side="left", padx=(0, 5))
        
        self.progress_percent_label = ttk.Label(self.progress_frame, text="0%", width=5, font=("Arial", 7, "bold"))
        self.progress_percent_label.pack(side="left", padx=(0, 10))
        
        # Progress status label (same line as progress bar)
        self.progress_label = ttk.Label(self.progress_frame, text="⏳ Initializing conversion...", font=("Arial", 6), foreground="blue")
        self.progress_label.pack(side="left", padx=(0, 10))

    def check_and_auto_load_recent_file(self):
        """Check for recently used Excel file and auto-load if found"""
        try:
            # Check current directory for Excel files with mapping
            current_dir = os.getcwd()
            excel_files = []
            for file in os.listdir(current_dir):
                if file.lower().endswith(('.xlsx', '.xls')):
                    excel_path = os.path.join(current_dir, file)
                    
                    # Generate mapping file name for this Excel
                    excel_name   = os.path.splitext(file)[0]
                    mapping_name = f"{excel_name}.~p~"
                    mapping_path = os.path.join(current_dir, mapping_name)
                    
                    if os.path.exists(mapping_path):
                        mtime = os.path.getmtime(excel_path)
                        excel_files.append((excel_path, mapping_path, mtime))
            
            if excel_files:
                # Sort by modification time (most recent first)
                excel_files.sort(key=lambda x: x[2], reverse=True)
                recent_excel, recent_mapping, _ = excel_files[0]             
                self.auto_load_specific_file(recent_excel, recent_mapping)
                
        except Exception as e:
            print(f"❌ Error checking recent file: {str(e)}")

    def auto_load_specific_file(self, excel_path, mapping_path):
        """Auto-load specific Excel file with its mapping"""
        try:
            # Set excel file
            self.excel_file = excel_path
            filename        = os.path.basename(excel_path)
            self.file_label.config(text=f"📁 {filename}", foreground="black")

            mapping_success = self.load_mapping_from_specific_file(mapping_path) # Load mapping first
            
            if mapping_success:
                self.load_excel_with_existing_mapping()                          # Load Excel data
            else:
                self.load_excel_data_silently() 
                
        except Exception as e:
            print(f"❌ Error auto-loading specific file: {str(e)}")
            
    def load_mapping_from_specific_file(self, mapping_path):
        """Load mapping from specific file path"""
        try:
            with open(mapping_path, 'r', encoding='utf-8') as f:
                load_data = json.load(f)
            
             # Load mappings
            self.sheet_mappings         = load_data.get("sheet_mappings", {})
            self.sheet_selected_headers = load_data.get("sheet_selected_headers", {})
            
            loaded_sheets               = len(self.sheet_mappings)
            selected_headers_count      = sum(len(headers) for headers in self.sheet_selected_headers.values())
            
            self.mapping_file_path      = mapping_path
            self.mapping_loaded         = True
            self.update_mapping_status(f"📂 Auto-loaded: {loaded_sheets} sheets", "blue")
            
            return True
            
        except Exception as e:
            print(f"❌ Error loading mapping file: {str(e)}")
            return False
            
    def load_excel_with_existing_mapping(self):
        """Load Excel data when mapping is already loaded"""
        try:
            self.status_label.config(text="⏳ Loading...", foreground="blue")
            self.root.update()
            
            xls = pd.ExcelFile(self.excel_file)
            self.all_sheets_data = {}

            # Load all sheets
            for sheet_name in xls.sheet_names:
                try:
                    df      = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=3)
                    headers = self.extract_headers(df)
                    self.all_sheets_data[sheet_name] = headers
                    
                except Exception as e:
                    continue
            
            if not self.all_sheets_data:
                raise Exception("No sheets could be read from the Excel file")

            # Update dropdown
            sheet_names = list(self.all_sheets_data.keys())
            self.sheet_combo['values'] = sheet_names

            # Set default sheet
            if sheet_names:
                default_sheet = sheet_names[0]
                self.sheet_combo.set(default_sheet)
                self.selected_sheet = default_sheet
            
            # Create UI first
            if self.selected_sheet:
                self.current_sheet_headers = self.all_sheets_data[self.selected_sheet]
                
                # Update display
                self.excel_info_label.config(
                    text=f"📊 {self.selected_sheet} ({len(self.current_sheet_headers)} headers)"
                )

                self.create_excel_header_widgets()      # Create widgets
                self.apply_saved_mapping_to_ui()        # Apply saved mapping
                self.update_standard_tree()             # Update standard tree
            
            self.status_label.config(text=f"✅ Loaded {len(sheet_names)} sheets", foreground="green")
            
            # Get unit system info
            unit_system = "Imperial Units" if self.use_imperial_units.get() else "SI Units"
            
            # Update result text
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, f"""
✅ Auto-loaded recent Excel file with saved mappings!

📁 File: {os.path.basename(self.excel_file)}
📊 Total Sheets: {len(sheet_names)}
💾 Restored mappings for {len(self.sheet_mappings)} sheets
📏 Unit System: {unit_system}

💡 Your previous mapping configurations have been automatically restored.
""")
            
            return True
            
        except Exception as e:
            print(f"❌ Error loading Excel with mapping: {str(e)}")
            traceback.print_exc()
            return False
            
    def apply_saved_mapping_to_ui(self):
        """Apply saved mapping to current UI"""
        try:
            if not self.selected_sheet or self.selected_sheet not in self.sheet_selected_headers:
                return

            saved_headers = self.sheet_selected_headers[self.selected_sheet]
            restored_count = 0
            
           # Apply to each widget
            for key, widget_info in self.excel_header_widgets.items():
                if widget_info['type'] == 'multiple':
                    # Find matching saved header
                    for saved_header in saved_headers:
                        if saved_header in widget_info['headers']:
                            widget_info['variable'].set(saved_header)
                            widget_info['selected'] = saved_header
                            widget_info['status'].config(text="Restored", foreground="green")
                            restored_count += 1
                            break
                    else:
                        # Set to first option as default
                        first_option = widget_info['headers'][0]
                        widget_info['variable'].set(first_option)
                        widget_info['selected'] = first_option
                        widget_info['status'].config(text="Default", foreground="blue")
                        
                elif widget_info['type'] == 'single':
                    if widget_info['selected'] in saved_headers:
                        widget_info['status'].config(text="Restored", foreground="green")
                        restored_count += 1
                    else:
                        widget_info['status'].config(text="Ready", foreground="blue")

            # Force UI update
            self.root.update_idletasks()
            
        except Exception as e:
            print(f"❌ Error applying saved mapping: {str(e)}")

    def get_recent_excel_file(self):
        """This function is no longer used"""
        pass

    def auto_load_recent_file_with_mapping(self, excel_file_path):
        """This function is no longer used"""
        pass

    def load_excel_data_silently(self):
        """Load Excel data without showing message dialog"""
        try:
            self.status_label.config(text="⏳ Loading...", foreground="blue")
            self.root.update()
            
            xls = pd.ExcelFile(self.excel_file)
            self.all_sheets_data = {}
            
            # Clear old mapping
            self.sheet_mappings = {}
            self.sheet_selected_headers = {}
            self.mapping_loaded = False

            # Load all sheets in Excel order
            for sheet_name in xls.sheet_names:
                try:
                    df      = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=3)
                    headers = self.extract_headers(df)
                    self.all_sheets_data[sheet_name] = headers
                    
                except Exception as e:
                    continue
            
            if not self.all_sheets_data:
                raise Exception("No sheets could be read from the Excel file")

            # Try to auto-load mapping file BEFORE setting up UI
            mapping_auto_loaded = self.auto_load_mapping_if_exists()

            # Update dropdown in Excel file
            sheet_names = list(self.all_sheets_data.keys())
            self.sheet_combo['values'] = sheet_names

            # Set default to first sheet in Excel order
            if sheet_names:
                default_sheet = sheet_names[0]  # First sheet in Excel order
                self.sheet_combo.set(default_sheet)
                self.selected_sheet = default_sheet
            
            # Load selected sheet AFTER loading mapping
            if self.selected_sheet:
                self.on_sheet_selected()
            
            # Force restore mapping after UI is ready
            if mapping_auto_loaded and self.selected_sheet:
                self.root.after(100, self.force_restore_mapping)
            
            self.status_label.config(text=f"✅ Loaded {len(sheet_names)} sheets", foreground="green")
            
            # Get unit system info
            unit_system = "Imperial Units" if self.use_imperial_units.get() else "SI Units"
            
            # Update result text if mapping was loaded
            if mapping_auto_loaded:
                self.result_text.delete(1.0, tk.END)
                self.result_text.insert(tk.END, f"""
✅ Auto-loaded recent Excel file with mappings!

📁 File: {os.path.basename(self.excel_file)}
📊 Total Sheets: {len(sheet_names)}
💾 Auto-loaded mappings for {len(self.sheet_mappings)} sheets
📏 Unit System: {unit_system}

💡 The previous mapping configurations have been automatically restored.
You can switch between sheets to see the saved configurations.
""")
            
            return True
            
        except Exception as e:
            print(f"❌ Error loading Excel data silently: {str(e)}")
            import traceback
            traceback.print_exc()
            self.status_label.config(text="❌ Error loading file", foreground="red")
            # Reset
            self.excel_file = ""
            self.file_label.config(text="No file selected", foreground="gray")
            return False

    def force_restore_mapping(self):
        """Force restore mapping after UI is fully ready"""
        try:
            if not self.selected_sheet or self.selected_sheet not in self.sheet_selected_headers:
                return

            saved_headers = self.sheet_selected_headers[self.selected_sheet]

            # Ensure widgets are created
            if not self.excel_header_widgets:
                self.create_excel_header_widgets()
                self.root.after(200, self.force_restore_mapping)    # Try again after a delay
                return
            
            self.restore_header_selections(saved_headers)           # Force restore selections
            
        except Exception as e:
            print(f"❌ Error in force restore: {str(e)}")

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
    🔧 Header Selector Tool
    📋 How to use:
    1. Click "🔍 Browse" to select Excel file
    2. Select Unit System: SI Units (m, mm, bar) or Imperial Units (ft, in, psi)
    3. Select Sheet from dropdown (sheets are ordered as in Excel file)
    4. Left side: Select desired headers from Excel (dropdown for duplicate columns)
    5. Right side: View mapping with standard headers (updated based on unit system)
    6. Use "🔄 Auto Map" for automatic mapping (selects first occurrence)
    7. Switch between sheets to configure mappings for each sheet
    8. Click "🔄 Convert to AccessDB" to convert all sheets with selected units
    
    💡 Features:
    - Unit System Selection: Choose between SI and Imperial units
    - Automatic Unit Conversion: Values and column names converted automatically
    - Mappings are saved per sheet automatically
    - Auto-save mapping when converting to Access DB
    - Auto-load existing mappings when opening same Excel file
    - First occurrence of duplicate columns selected by default
    - Unmapped Excel columns become TempData1, TempData2, etc.
    - All sheets converted together with individual mappings
    
    📏 Unit Conversion:
    - SI Units: m, mm, bar, m/s, m/min
    - Imperial Units: ft, in, psi, ft/s, ft/min
    - Automatic conversion of values and column names
    - Supported columns: distance, altitude, diameter, thickness, velocity
    
    💾 Auto Mapping Files:
    - Auto-saved as: MappingHeader_[ExcelFileName].~p~
    - Stored in the same directory as Excel file
    - Auto-loaded when opening the same Excel file again
    - Contains all sheet configurations and mappings
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

    def show_helper(self):
        """Show Helper dialog with Word Software Manual"""
        try:
            manual_files = self.find_manual_files()     # Look for Word manual files in current directory
        
            if not manual_files:                        # If no manual file found, show error message
                messagebox.showwarning(
                    "📄 Manual Not Found", 
                    "Software Manual not found!\n\n"
                    "Please ensure 'SOFTWARE MANUAL for Selected header.docx'\n"
                    "is placed in the same directory as this Python script"
                )
            else:                                       # If found, show selection or open directly
                self.open_word_manual(manual_files[0])  # Open the first (or only) manual file found
                
        except Exception as e:
            print(f"❌ Error opening helper: {str(e)}")
            messagebox.showerror("❌ Error", f"Cannot open Software Manual:\n{str(e)}")

    def find_manual_files(self):
        """Find Word manual files in current directory"""
        manual_files = []
        current_dir  = os.getcwd()
    
        # Common manual file patterns
        manual_patterns = [
            "*manual*.docx", "*manual*.doc",
            "*Manual*.docx", "*Manual*.doc", 
        ]
    
        for pattern in manual_patterns:
            found_files = glob.glob(os.path.join(current_dir, pattern))
            manual_files.extend(found_files)
    
        # Remove duplicates and sort
        manual_files = list(set(manual_files))
        manual_files.sort()

        return manual_files

    def open_word_manual(self, manual_path):
        """Open Word manual file"""
        try:
            if not os.path.exists(manual_path):
                messagebox.showerror("❌ Error", f"Manual file not found:\n{manual_path}")
                return

            # Try to open with system default application
            if platform.system() == "Windows":
                os.startfile(manual_path)
        
            # Update result text
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, f"""
    📖 Software Manual Opened Successfully!

    📄 Manual File: {os.path.basename(manual_path)}
    📂 File Location: {os.path.dirname(manual_path)}
    📊 File Size: {os.path.getsize(manual_path) / 1024:.1f} KB
    📅 Opened: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

    💡 The Software Manual has been opened in your default Word application.
    Please refer to the manual for detailed instructions on using this tool.

    🔧 If you need quick help while using the tool, you can always click the Helper button again.
    """)
        
        except Exception as e:
            error_msg = f"Cannot open manual file:\n{str(e)}"
            print(f"❌ {error_msg}")
            messagebox.showerror("❌ Error", error_msg)

    def load_excel_data(self):
        """Load data from Excel file - Added auto-load mapping"""
        try:
            self.status_label.config(text="⏳ Loading...", foreground="blue")
            self.root.update()
            
            xls = pd.ExcelFile(self.excel_file)
            self.all_sheets_data = {}
            
            # Clear old mapping
            self.sheet_mappings = {}
            self.sheet_selected_headers = {}
            self.mapping_loaded = False

            for i, sheet_name in enumerate(xls.sheet_names, 1):
                print(f"  {i}. {sheet_name}")
            
            # Load all sheets in Excel order
            for sheet_name in xls.sheet_names:
                try:
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=3)
                    headers = self.extract_headers(df)
                    self.all_sheets_data[sheet_name] = headers

                except Exception as e:
                    continue
            
            if not self.all_sheets_data:
                raise Exception("No sheets could be read from the Excel file")

            mapping_auto_loaded = self.auto_load_mapping_if_exists()    # Try to auto-load mapping file BEFORE setting up UI
            sheet_names = list(self.all_sheets_data.keys())             # Update dropdown in Excel file
            self.sheet_combo['values'] = sheet_names

            # Set default to first sheet in Excel order
            if sheet_names:
                default_sheet = sheet_names[0]  # First sheet in Excel order
                self.sheet_combo.set(default_sheet)
                self.selected_sheet = default_sheet
            
            # Load selected sheet AFTER loading mapping
            if self.selected_sheet:
                self.on_sheet_selected()
            
            # Force restore mapping after UI is ready (for manual browse)
            if mapping_auto_loaded and self.selected_sheet:
                self.root.after(100, self.force_restore_mapping)
            
            self.status_label.config(text=f"✅ Loaded {len(sheet_names)} sheets", foreground="green")
            
            # Get unit system info
            unit_system = "Imperial Units" if self.use_imperial_units.get() else "SI Units"
            
            # Show success message
            if mapping_auto_loaded:
                message_text = (f"Excel file loaded successfully!\n\n"
                    f"📋 Number of Sheets: {len(sheet_names)}\n"
                    f"📊 Default Sheet: {default_sheet} (first in Excel order)\n"
                    f"💾 Mappings: Auto-loaded from saved file\n"
                    f"📏 Unit System: {unit_system}\n"
                    f"📁 File: {os.path.basename(self.excel_file)}")
                
                # Update result text to show auto-load info
                self.result_text.delete(1.0, tk.END)
                self.result_text.insert(tk.END, f"""
✅ Excel file and mappings loaded successfully!

📁 File: {os.path.basename(self.excel_file)}
📊 Total Sheets: {len(sheet_names)}
💾 Auto-loaded mappings for {len(self.sheet_mappings)} sheets
📏 Unit System: {unit_system}

💡 You can now review the loaded mappings by switching between sheets.
The saved configurations have been automatically applied.
""")
            else:
                message_text = (f"Excel file loaded successfully!\n\n"
                    f"📋 Number of Sheets: {len(sheet_names)}\n"
                    f"📊 Default Sheet: {default_sheet} (first in Excel order)\n"
                    f"💾 Mappings: Starting fresh (no saved file found)\n"
                    f"📏 Unit System: {unit_system}\n"
                    f"📁 File: {os.path.basename(self.excel_file)}")
            
            messagebox.showinfo("✅ Success", message_text)
            
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

            # Analyze feature types for Max column filtering
            if "List of Nominal Wall Thickness" not in self.selected_sheet:
                feature_analysis = self.analyze_feature_types_in_sheet(self.selected_sheet)

            self.create_excel_header_widgets()      # Create header widgets first
            self.load_saved_sheet_mapping()         # Then load saved mapping if any
            self.update_standard_tree()             # Finally update standard tree

    def save_current_sheet_mapping(self):
        """Save mapping of current sheet"""
        if not self.selected_sheet or not self.excel_header_widgets: 
            return
        
        # Store selected headers of this sheet
        selected_headers = self.get_selected_excel_headers()
        if selected_headers:                                # Only save if we have headers
            self.sheet_selected_headers[self.selected_sheet] = selected_headers
            
            # Store mapping with standard headers
            final_mappings = self.get_final_mappings()
            self.sheet_mappings[self.selected_sheet] = final_mappings

    def load_saved_sheet_mapping(self):
        """Load saved mapping for this sheet"""
        if not self.selected_sheet: 
            return
        
        if self.selected_sheet in self.sheet_selected_headers:
            saved_headers = self.sheet_selected_headers[self.selected_sheet]
            self.restore_header_selections(saved_headers)   # Restore header selection
        else:
            self.auto_select_best_options()                  # If no saved data, perform auto-mapping

    def restore_header_selections(self, saved_headers):
        """Restore header selection from saved data"""
        restored_count = 0
        
        for key, widget_info in self.excel_header_widgets.items():
            if widget_info['type'] == 'multiple':
                for saved_header in saved_headers:
                    if saved_header in widget_info['headers']:
                        widget_info['variable'].set(saved_header)
                        widget_info['selected'] = saved_header
                        widget_info['status'].config(text="Restored", foreground="green")
                        restored_count += 1
                        break
                    
            elif widget_info['type'] == 'single':
                # For single headers, make sure it's in the saved list
                if widget_info['selected'] in saved_headers:
                    widget_info['status'].config(text="Restored", foreground="green")
                    restored_count += 1

        self.root.update_idletasks() # Force update the UI
        self.update_standard_tree()  # Update the standard tree to reflect restored mappings

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
        unit_system = "Imperial Units" if self.use_imperial_units.get() else "SI Units"
        messagebox.showinfo("🔄 Auto Mapping Complete", 
            f"Auto-mapped {mapped_count}/{len(current_standards)} headers\n\n"
            f"📊 Sheet: {self.selected_sheet}\n"
            f"📋 Header Type: {sheet_type}\n"
            f"📏 Unit System: {unit_system}\n"
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
        
        unit_system = "Imperial Units" if self.use_imperial_units.get() else "SI Units"
        
        report = f"""
        {'='*80}
        📊 Header Mapping Report - All Sheets
        {'='*80}

        📁 File: {os.path.basename(self.excel_file)}
        📋 Total Sheets: {len(self.all_sheets_data)}
        📊 Current Sheet: {self.selected_sheet}
        📏 Unit System: {unit_system}
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
            
            report += f"📋 {sheet_name} ({sheet_type}, {unit_system.split()[0]}): {status}\n"
        
        report += f"""
        {'─'*80}
        📈 Overall Summary:
        {'─'*80}
        ├─ 📊 Total Sheets: {len(self.all_sheets_data)}
        ├─ ✅ Configured Sheets: {total_configured_sheets}
        ├─ ⚠️ Unconfigured Sheets: {len(self.all_sheets_data) - total_configured_sheets}
        ├─ 📋 Total Standard Headers: {total_standard_headers}
        ├─ ✅ Total Mapped Headers: {total_mapped_headers}
        ├─ 📏 Unit System: {unit_system}
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
        
        # Ensure we have current sheet mapping with proper widgets
        if self.selected_sheet and self.excel_header_widgets:
            self.save_current_sheet_mapping()
        
        # Auto-save mapping before conversion (always save if there are mappings)
        if self.sheet_mappings and any(self.sheet_selected_headers.values()):
            success, result = self.save_mapping_to_file()
            if success:
                self.update_mapping_status(f"💾 Auto-saved: {os.path.basename(result)}", "green")
        
        # Count configured sheets
        configured_count   = len(self.sheet_mappings)
        unconfigured_count = len(self.all_sheets_data) - configured_count
        
        # Get unit system info
        unit_system = "Imperial Units" if self.use_imperial_units.get() else "SI Units"
        
        # Confirm conversion
        result = messagebox.askyesno(
            "🔄 Convert to Access Database",
            f"Convert Excel file to Access Database?\n\n"
            f"📁 Source: {os.path.basename(self.excel_file)}\n"
            f"📊 Total Sheets: {len(self.all_sheets_data)}\n"
            f"✅ Configured Sheets: {configured_count}\n"
            f"📋 Unconfigured Sheets: {unconfigured_count}\n"
            f"📏 Unit System: {unit_system}\n"
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
        self.progress_frame.pack(side="left", expand=True, fill="x", padx=(20, 20))
        
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
            self.progress_percent_label.config(text=f"{percent:.0f}%")
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
                header_mapping_content=header_mapping_content,
                use_imperial=self.use_imperial_units.get()
            )
            
            if self.conversion_cancelled: return
            
            # Step 8: Finalization (100%)
            if result:
                access_file = os.path.splitext(self.excel_file)[0] + ".accdb"
                self.last_converted_db_path = access_file  # Store database path
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
        
        # Get unit system info
        unit_system = "Imperial Units" if self.use_imperial_units.get() else "SI Units"
        
        success_report = f"""
        {'='*60}
        ✅ CONVERSION COMPLETED SUCCESSFULLY!
        {'='*60}

        📁 Source: {os.path.basename(self.excel_file)}
        💾 Output: {os.path.basename(access_file)}
        📊 Total Sheets Converted: {len(self.all_sheets_data)}
        ✅ Custom Mappings Applied: {configured_count} sheets
        📋 Original Headers Used: {unconfigured_count} sheets
        📏 Unit System: {unit_system}
        📄 HeaderMapping: Added to Access DB

        🎉 Conversion finished successfully!
        All data has been converted and saved to the Access database.
        """
        
        self.result_text.insert(tk.END, success_report)
        
        # Show custom completion dialog with option to open database location
        if not self.conversion_cancelled:
            self.show_conversion_complete_dialog(access_file, configured_count, unconfigured_count)

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

    def save_mapping_dialog(self):
        """This function is no longer used - auto save on convert"""
        pass

    def load_mapping_dialog(self):
        """This function is no longer used - auto load on file open"""
        pass

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
                unit_system = "Imperial Units" if self.use_imperial_units.get() else "SI Units"
                content += f"Sheet Type: {sheet_type}\n"
                content += f"Unit System: {unit_system}\n"
           
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
            content += f"📏 Unit System: {unit_system}\n"
            content += f"💾 Mapping Config File: {os.path.basename(self.mapping_file_path) if self.mapping_file_path else 'Not saved'}\n"
        
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

    def analyze_feature_types_in_sheet(self, sheet_name):
        """Analyze Feature type column in the specified sheet to determine which Max columns to show"""
        try:
            if not sheet_name or sheet_name not in self.all_sheets_data:
                return None

            xls = pd.ExcelFile(self.excel_file)
            
            # First, read just a few rows to find the Feature identification column
            df_sample = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=5)
            
            # Use the same logic as extract_headers to find the correct header row
            if len(df_sample) < 2: 
                return None
            
            first_row  = df_sample.iloc[0].fillna("").tolist()
            second_row = df_sample.iloc[1].fillna("").tolist()
            
            # Find Feature identification column in the correct header row
            feature_type_col_idx  = None
            feature_type_col_name = None
            
            # Check second row first (as per extract_headers logic)
            for i, val in enumerate(second_row):
                if str(val).strip() != '':
                    col_str = str(val).strip().lower()
                    if 'feature identification' in col_str or 'featureidentification' in col_str:
                        feature_type_col_idx = i
                        feature_type_col_name = str(val).strip()
                        break
            
            if feature_type_col_idx is None:
                return None
            
            chunk_size              = 1000                              # Process 1000 rows at a time
            feature_types           = set()                             # Use set for efficient deduplication
            target_features_found   = set()
            target_features = ['CRACK', 'CRAL', 'COCL', 'CORR', 'LAMI', 'MAIN'] # Define target feature types
            
            try:
                # Read Excel in chunks
                for chunk_start in range(0, 1000, chunk_size):          # Limit to 1000 rows max
                    df_chunk = pd.read_excel(xls, sheet_name=sheet_name, header=None, skiprows=chunk_start, nrows=chunk_size)
                    
                    if len(df_chunk) == 0:
                        break  # No more data

                    start_row = 2 if chunk_start == 0 else 0            # Skip header rows in first chunk

                    if feature_type_col_idx < len(df_chunk.columns):    # Extract values from the Feature identification column
                        values = df_chunk.iloc[start_row:, feature_type_col_idx].dropna()
                        values = values.astype(str).str.strip().str.upper()
                       
                        feature_types.update(values.unique())           # Add all unique values to feature_types

                        for target in target_features:                  # Check for target features
                            if target in values.values:
                                target_features_found.add(target)
                        
                        if len(target_features_found) >= 2:             # Found both height and depth features
                            break

                feature_types = list(feature_types)                     # Convert set to list

            except Exception as e:
                df_simple = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=100)
                feature_types = []
                
                for row_idx in range(2, len(df_simple)):
                    if feature_type_col_idx < len(df_simple.iloc[row_idx]):
                        val = df_simple.iloc[row_idx, feature_type_col_idx]
                        if pd.notna(val) and str(val).strip() != '':
                            feature_types.append(str(val).strip().upper())

                feature_types = list(set(feature_types))
            
            # Determine which Max columns should be shown (show only one type)
            height_features = self.FEATURE_TYPE_CONFIG['height_features']
            depth_features  = self.FEATURE_TYPE_CONFIG['depth_features']
            
            has_height_features = any(ft in height_features for ft in feature_types)
            has_depth_features  = any(ft in depth_features for ft in feature_types)

            result = {
                'feature_types': feature_types,
                'show_max_height': has_height_features,
                'show_max_depth': has_depth_features
            }
            
            return result
            
        except Exception as e:
            print(f"⚠️ Error analyzing Feature types in sheet '{sheet_name}': {str(e)}")
            import traceback
            traceback.print_exc()
            return None

    def adjust_log_distance_header(self, base_headers):
        """Adjust Log distance header based on Excel headers to match the unit"""
        if not self.selected_sheet or not self.current_sheet_headers:
            return base_headers
        
        # Check if we're using Imperial units
        if not self.use_imperial_units.get():
            return base_headers
        
        # Look for Log distance headers in Excel
        log_distance_excel = None
        for header in self.current_sheet_headers:
            if 'log distance' in header.lower():
                log_distance_excel = header
                break
        
        if not log_distance_excel:
            return base_headers
        
        # Determine the unit from Excel header
        excel_unit = None
        if '[mi]' in log_distance_excel:
            excel_unit = 'mi'
        elif '[ft]' in log_distance_excel:
            excel_unit = 'ft'
        
        if not excel_unit:
            return base_headers
        
        # Adjust Standard Headers to match Excel unit
        adjusted_headers = []
        for header in base_headers:
            if 'log distance' in header.lower():
                if excel_unit == 'mi':
                    adjusted_headers.append("Log distance [mi]")
                elif excel_unit == 'ft':
                    adjusted_headers.append("Log distance [ft]")
            else:
                adjusted_headers.append(header)
        
        return adjusted_headers

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