import pandas as pd
import numpy as np
import pyodbc
import os
import sys
import openpyxl
import re
import subprocess
import ctypes
from ctypes import wintypes

# Unit conversion factors
UNIT_CONVERSION_FACTORS = {
    'm_to_ft': 3.28084,      # meters to feet
    'm_to_mi': 0.000621371,  # meters to miles
    'ft_to_mi': 0.000189394, # feet to miles
    'mi_to_ft': 5280.0,      # miles to feet
    'mi_to_mi': 1.0,         # miles to miles (no conversion)
    'ft_to_ft': 1.0,         # feet to feet (no conversion)
    'mm_to_in': 0.0393701,   # millimeters to inches
    'bar_to_psi': 14.5038,   # bar to psi
    'ft_to_m': 0.3048,       # feet to meters
    'in_to_mm': 25.4,        # inches to millimeters
    'psi_to_bar': 0.0689476  # psi to bar
}

# Feature type configuration - integrated directly into the file
FEATURE_TYPE_CONFIG = {
    'height_features': ['CRAL', 'CRACK'],
    'depth_features': ['COCL', 'CORR', 'LAMI'],
    'priority': 'height'  # If both types exist, prefer height features
}

def create_empty_accdb(file_path):
    """Create an empty .accdb file using JET/ACE engine via ctypes"""
    try:
        # Ensure the directory exists
        directory = os.path.dirname(os.path.abspath(file_path))
        if not os.path.exists(directory):
            os.makedirs(directory)

        # Make sure file doesn't exist
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
                print(f"Removed existing database: {file_path}")
            except Exception as e:
                print(f"Warning: Could not remove existing file: {str(e)}")
                return False

        # Load the ACE DLL
        try:
            acedb = ctypes.windll.LoadLibrary("acecore.dll")
            print("Loaded acecore.dll")
        except:
            try:
                # Try alternate DLL names that might be available
                acedb = ctypes.windll.LoadLibrary("msjetoledb40.dll")
                print("Loaded msjetoledb40.dll")
            except:
                try:
                    acedb = ctypes.windll.LoadLibrary("msjet40.dll")
                    print("Loaded msjet40.dll")
                except:
                    # Final fallback: try to create it by running a VBA script 
                    # through an external process
                    return create_access_with_vbscript(file_path)

        # Define error checking function
        def check_result(result, func, args):
            if result != 0:
                raise ctypes.WinError(result)
            return result

        # Get JetCreateDatabase function
        create_db_func = acedb.JetCreateDatabase
        create_db_func.argtypes = [ctypes.c_wchar_p, ctypes.c_void_p]
        create_db_func.restype = ctypes.c_ulong
        create_db_func.errcheck = check_result

        # Call the function to create the database
        result = create_db_func(file_path, None)
        print(f"Successfully created database: {file_path}")
        return True

    except Exception as e:
        print(f"Error creating database with ctypes: {str(e)}")
        
        # Try fallback approach
        return create_access_with_vbscript(file_path)

def create_access_with_vbscript(file_path):
    """Create an Access database using VBScript as a fallback"""
    try:
        # Create a temporary VBScript file
        vbs_path = os.path.join(os.path.dirname(file_path), "create_db.vbs")
        
        # Write the VBScript content
        with open(vbs_path, "w") as f:
            f.write(f'''
Set objCatalog = CreateObject("ADOX.Catalog")
objCatalog.Create "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={file_path}"
Set objCatalog = Nothing
''')
        
        # Run the VBScript
        subprocess.run(["cscript", "//NoLogo", vbs_path], check=True)
        
        # Remove the temporary script
        os.remove(vbs_path)
        
        print(f"Successfully created database with VBScript: {file_path}")
        return True
    
    except Exception as e:
        print(f"Error creating database with VBScript: {str(e)}")
        
        # Try one more approach
        return create_access_with_powershell(file_path)

def create_access_with_powershell(file_path):
    """Create an Access database using PowerShell as a final fallback"""
    try:
        # Create a temporary PowerShell script
        ps_path = os.path.join(os.path.dirname(file_path), "create_db.ps1")
        
        # Write the PowerShell script content
        with open(ps_path, "w") as f:
            f.write(f'''
$connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={file_path};"
$adoxCatalog = New-Object -ComObject ADOX.Catalog
$adoxCatalog.Create($connectionString)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($adoxCatalog) | Out-Null
''')
        
        # Run the PowerShell script
        subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", ps_path], check=True)
        
        # Remove the temporary script
        os.remove(ps_path)
        
        print(f"Successfully created database with PowerShell: {file_path}")
        return True
    
    except Exception as e:
        print(f"Error creating database with PowerShell: {str(e)}")
        return False

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def set_specific_headers(df, sheetname):
    """Set column headers from the first and second rows of the sheet."""
    df = df.copy()
    
    # Extract the first two rows which may contain header information
    first_row   = df.iloc[0].ffill().tolist()
    second_row  = df.iloc[1].ffill().tolist()

    # Combine the rows to create the headers
    headers = []
    for i in range(len(second_row)):
        if pd.notna(second_row[i]) and second_row[i].strip() != '':
            headers.append(second_row[i].strip())
        elif pd.notna(first_row[i]) and first_row[i].strip() != '':
            headers.append(first_row[i].strip())
        else:
            headers.append(f"Unnamed_{i}")  # Assign a placeholder for truly empty headers

    # Set the last column name to 'Comments' unless it's the 'List of Nominal Wall Thickness' sheet.
    if sheetname != "List of Nominal Wall Thickness":
        has_comment_column = any(
            str(header).lower().strip() in ['comment', 'comments'] 
            for header in headers
        )
        
        if not has_comment_column and len(headers) > 1:
            headers[-1] = "Comments"        # Forcefully assign if missing
            print(f"📝 Added 'Comments' as last column in sheet '{sheetname}'")
        else:
            print(f"📋 Sheet '{sheetname}' already has comment column, skipping auto-assignment")
    
    # Ensure all headers are strings and unique
    unique_headers = []
    for i, header in enumerate(headers):
        header_str = str(header)  # Convert to string
        if header_str in unique_headers:
            unique_headers.append(f"{header_str}_{i}")
        else:
            unique_headers.append(header_str)

    df.columns = unique_headers
    df = df.drop([0, 1]).reset_index(drop=True)   # Drop the first two rows after setting headers and reset index
    return df

def GetHeaderColumn(df):
    headers = df.iloc[0].tolist()       # Create a list for new headers
    headers = pd.Series(headers).fillna("Unnamed")
    
    # Ensure all headers are strings and unique
    unique_headers = []
    for i, header in enumerate(headers):
        header_str = str(header)  # Convert to string
        if header_str in unique_headers:
            unique_headers.append(f"{header_str}_{i}")
        else:
            unique_headers.append(header_str)
    
    df.columns = unique_headers
    
    return df

def compare_arrays_with_alert(temp, data):
    # Convert both arrays to sets
    temp_set = set(temp)
    data_set = set(data)
    
    potentially_missing = temp_set - data_set    # Find elements in temp that are not in data (potentially missing)
    extra_in_data       = data_set - temp_set    # Find elements in data that are not in temp (extra)
    
    # Initialize variables
    missing = set()
    misspelled = []  
    true_extra = []
    
     # Check for potential misspellings and true extra data
    for word in extra_in_data:
        if any(sum((c1 != c2) for c1, c2 in zip(word, temp_word)) <= 2 and abs(len(word) - len(temp_word)) <= 2 for temp_word in temp_set):
            misspelled.append(word)
        else:
            true_extra.append(word)
    
    # Check if potentially missing columns are truly missing or just misspelled
    for temp_word in potentially_missing:
        if not any(sum((c1 != c2) for c1, c2 in zip(temp_word, data_word)) <= 2 and abs(len(temp_word) - len(data_word)) <= 2 for data_word in data_set):
            missing.add(temp_word)
    
    # Check data is OK
    if len(missing) == 0  and len(misspelled) == 0 and len(true_extra) == 0:
        message  = "OK"
    else:
        message = "HAVE ERROR"
      
    return  message, misspelled, true_extra, list(missing)

def clean_column_name(column_name):
    """Modify column names to be compatible with Access"""
    clean_name = column_name.replace('[', '(').replace(']', ')')    # Replace square brackets [] with parentheses ()
    clean_name = clean_name.replace('.', '')                        # Remove '.'
    clean_name = clean_name[:64]                                    # Trim names if they exceed the length limit (Access restriction)
   
    return clean_name

def get_access_data_type(df_column, column_name=None, table_name=None):
    """ Assign appropriate data types for columns in Access"""
    
    # For List of Pipe Tally: Log distance as DOUBLE, all others as TEXT
    if table_name == "List of Pipe Tally":
        if column_name and "Log distance" in column_name:
            return "DOUBLE"
        else:
            return "TEXT(255)"
    
    # For other tables: use original logic
    if pd.api.types.is_float_dtype(df_column):
        return "DOUBLE"
    elif pd.api.types.is_integer_dtype(df_column):
        return "LONG"
    elif pd.api.types.is_bool_dtype(df_column):
        return "BIT"
    elif pd.api.types.is_datetime64_dtype(df_column):
        return "DATETIME"
    else:
        # Check the maximum length of data in the column
        max_length = 0
        for val in df_column:
            if isinstance(val, str) and len(val) > max_length:
                max_length = len(val)
        
        if max_length > 255:
            return "MEMO"
        else:
            return "TEXT(255)"

def create_access_table(cursor, table_name, df):
    """Create a table in Access with appropriate data types"""
    try:
        try:
            # Delete the existing table if it exists
            cursor.execute(f"DROP TABLE [{table_name}]")
            cursor.commit()
            print(f"Dropped existing table: {table_name}")
        except:
            pass  # No table found, proceed with the process
        
        # Create a dictionary to store column name replacements
        column_map = {}
        for col in df.columns:
            clean_col = clean_column_name(col)
            column_map[col] = clean_col

        # Generate SQL statement for creating the table
        columns = []
        for col in df.columns:
            clean_col = column_map[col]
            data_type = get_access_data_type(df[col], column_name=col, table_name=table_name)
            columns.append(f"[{clean_col}] {data_type}")
        
        # Generate SQL statement
        create_table_sql = f"CREATE TABLE [{table_name}] ({', '.join(columns)})"
        
        # Execute table creation
        print(f"Creating table with SQL: {create_table_sql}")
        cursor.execute(create_table_sql)
        cursor.commit()
        
        # Rename columns in the DataFrame
        df.columns = [column_map[col] for col in df.columns]

        return True, df

    # Method 2: Set all columns as TEXT
    except Exception as e:
        print(f"Error creating table: {str(e)}")
        print(f"SQL: {create_table_sql}")
        
        try:
            columns = []
            for col in df.columns:
                clean_col = column_map[col]
                if table_name == "List of Pipe Tally" and "Log distance" in col:
                    columns.append(f"[{clean_col}] DOUBLE")
                else:
                    columns.append(f"[{clean_col}] TEXT(255)")
            
            create_table_sql = f"CREATE TABLE [{table_name}] ({', '.join(columns)})"
            cursor.execute(create_table_sql)
            cursor.commit()
            
            # Rename columns in the DataFrame table
            df.columns = [column_map[col] for col in df.columns]
            
            print(f"Table '{table_name}' created successfully with alternate method")
            return True, df
        except Exception as e2:
            raise Exception(f"Failed to create table: {str(e2)}")

def insert_data_to_access(cursor, table_name, df):
    """Import data from the DataFrame into the Access table"""
    try:
        # Process data in batches
        batch_size = 50
        total_rows = len(df)
        
        if total_rows == 0:
            print(f"Warning: No data to insert into table '{table_name}'")
            return True
        
        column_names = ", ".join([f"[{col}]" for col in df.columns])
        placeholders = ", ".join(["?" for _ in range(len(df.columns))])
        sql_insert = f"INSERT INTO [{table_name}] ({column_names}) VALUES ({placeholders})"
        
        # print(f"Insert SQL: {sql_insert}")
        
        inserted_count = 0
        skipped_count = 0
        
        for start_idx in range(0, total_rows, batch_size):
            end_idx = min(start_idx + batch_size, total_rows)
            batch_df = df.iloc[start_idx:end_idx]
            
            for _, row in batch_df.iterrows():
                
                # Convert None and empty values to None for all columns
                values = []
                for val in row:
                    if pd.isna(val) or val == "" or val == "None" or val is None:
                        values.append(None)
                    elif isinstance(val, (int, float, bool)):
                        values.append(val)
                    else:
                        # Convert to string and truncate length
                        values.append(str(val)[:255] if val is not None else None)
                
                try:
                    cursor.execute(sql_insert, values)
                    inserted_count += 1
                except Exception as e:
                    print(f"Error inserting row: {str(e)}")
                    skipped_count += 1
                    continue
            
            cursor.commit()
        return True
    
    except Exception as e:
        print(f"Error in insert_data_to_access: {str(e)}")
        raise

def create_access_database(file_path):
    """Create a new Access database"""
    
    # Use the ctypes implementation to create the database
    return create_empty_accdb(file_path)

def convert_value_to_imperial(value, factor_key):
    """Convert a single value to Imperial units"""
    if pd.isna(value) or value is None:
        return value
    
    try:
        numeric_value   = float(value)
        factor          = UNIT_CONVERSION_FACTORS[factor_key]
        converted_value = numeric_value * factor
        return converted_value

    except (ValueError, TypeError, KeyError):
        return value

def handle_max_columns_by_feature_type(df, use_imperial=False):
    """Handle Max columns based on Feature type using vectorized operations"""
    
    # Check if Feature type column exists
    feature_type_col = None
    for col in df.columns:
        if 'Feature identification' in col.lower() or 'Feature iden' in col.lower():
            feature_type_col = col
            break
    
    if feature_type_col is None:
        print("⚠️ Feature type column not found, using default Max column handling")
        # Default handling - process all Max columns
        max_columns = [col for col in df.columns if 'max' in col.lower()]
        for col in max_columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        return df
    
    # Get Max columns
    max_height_cols = [col for col in df.columns if 'max' in col.lower() and 'height' in col.lower()]
    max_depth_cols = [col for col in df.columns if 'max' in col.lower() and 'depth' in col.lower()]
    
    # Convert Feature type to uppercase for comparison
    df[feature_type_col] = df[feature_type_col].astype(str).str.strip().str.upper()
    
    # Create masks for different feature types
    height_features = FEATURE_TYPE_CONFIG['height_features']
    depth_features  = FEATURE_TYPE_CONFIG['depth_features']
    
    height_mask = df[feature_type_col].isin(height_features)
    depth_mask  = df[feature_type_col].isin(depth_features)
    
    # Process Max. Height columns
    for col in max_height_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')   # Convert to numeric first
            df.loc[depth_mask, col] = None                      # Set to None for depth features
    
    # Process Max. Depth columns  
    for col in max_depth_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')   # Convert to numeric first
            df.loc[height_mask, col] = None                     # Set to None for height features
    
    # Process other feature types (keep both columns)
    other_mask = ~(height_mask | depth_mask)
    for col in max_height_cols + max_depth_cols:
        if col in df.columns:
            df.loc[other_mask, col] = pd.to_numeric(df.loc[other_mask, col], errors='coerce')
    
    # Print column types
    for col in max_height_cols + max_depth_cols:
        if col in df.columns:
            print(f"  - {col}: {df[col].dtype}")
    
    return df

def convert_data_types(df, use_imperial=False):
    """Convert data types and handle unit conversions using dynamic pattern matching"""
    
    for col in df.columns:
        if "Log distance" in col or "distance" in col.lower():
            print(f"  - {col}: {df[col].dtype}")

    # Handle Max columns based on Feature type
    df = handle_max_columns_by_feature_type(df, use_imperial)
 
    return df

def add_erf_type(df):
    """Add ERF flag based on ERF column."""

    if 'ERF (Modified)' in df.columns and 'ERF (metal loss)' in df.columns:

        position = df.columns.get_loc('ERF (Modified)')                     #It finds the position of 'ERF (Modified)' column.

        #Creates a new 'ERF' column, using 'ERF (Modified)' values if they're not null, otherwise using 'ERF (metal loss)' values.
        df['ERF'] = df.apply(lambda row: row['ERF (Modified)'] if pd.notnull(row['ERF (Modified)']) else row['ERF (metal loss)'], axis=1)
        df.insert(position, 'ERF', df.pop('ERF'))                           #Inserts the new 'ERF' column at the position of 'ERF (Modified)'.

        df['isNormalERF'] = df['ERF (metal loss)'].notnull()                #Creates an 'isNormalERF' column, which is True where 'ERF (metal loss)' is not null.
        df = df.drop(columns = ['ERF (Modified)', 'ERF (metal loss)'])      #Drops the original 'ERF (Modified)' and 'ERF (metal loss)' columns.

    elif 'ERF (Modified)' in df.columns:
        position = df.columns.get_loc('ERF (Modified)')
        df['ERF'] = df['ERF (Modified)']
        df.insert(position, 'ERF', df.pop('ERF'))

        df['isNormalERF'] = False
        df = df.drop(columns=['ERF (Modified)'])

    elif 'ERF (metal loss)' in df.columns:
        position = df.columns.get_loc('ERF (metal loss)')
        df['ERF'] = df['ERF (metal loss)']
        df.insert(position, 'ERF', df.pop('ERF'))

        df['isNormalERF'] = True
        df = df.drop(columns=['ERF (metal loss)'])

    else:
        df['isNormalERF'] = True
    return df

def create_new_tables(cursor):
    """Create 4 new tables that do not exist in Excel"""
    # Create table Velocity
    try:
        cursor.execute("""CREATE TABLE [Velocity] (
                                        [Log distance (m)] DOUBLE,
                                        [Velocity (m/min)] DOUBLE,
                                        [Velocity (m/sec)] DOUBLE,
                                        [Feature] TEXT(255))
                                        """)
        cursor.commit()
        print("Successfully created table Velocity")
    except Exception as e:
        print(f"Unable to create table Velocity: {str(e)}")
    
    # Create table ClientInspection
    try:
        cursor.execute("""CREATE TABLE [ClientInspection] (
                                        [Project no] TEXT(255),
                                        [Project] TEXT(255),
                                        [Client] TEXT(255),
                                        [Inspection Date] TEXT(255),
                                        [Launcher] TEXT(255),
                                        [Receiver] TEXT(255),
                                        [Pipeline Designation] TEXT(255),
                                        [Product] TEXT(255),
                                        [Revision] TEXT(255))
                                        """) 
        cursor.commit()
        print("Successfully created table ClientInspection")
    except Exception as e:
        print(f"Unable to create table ClientInspection: {str(e)}")
    
    # Create table PipelineParameters
    try:
        cursor.execute("""CREATE TABLE [PipelineParameters] (
                                        [Outside Diameter] TEXT(255),
                                        [pipelineMaterial] TEXT(255),
                                        [NW Thickness] TEXT(255),
                                        [pipeline Class] TEXT(255),
                                        [Internal Diameter] TEXT(255),
                                        [pipe length] TEXT(255),
                                        [const Code] TEXT(255),
                                        [Max AlloOperPres] TEXT(255),
                                        [Design Press] TEXT(255),
                                        [SMYS] TEXT(255),
                                        [Design Factor] TEXT(255),
                                        [Construction Year] TEXT(255))
                                        """)
        cursor.commit()
        print("Successfully created table PipelineParameters")
    except Exception as e:
        print(f"Unable to create table PipelineParameters: {str(e)}")
    
    # Create table DataQuality
    try:
        cursor.execute("""CREATE TABLE [DataQuality] (
                                        [Inspection Direction] TEXT(255),
                                        [Launching Date] TEXT(255),
                                        [Receiving Date] TEXT(255),
                                        [Duration] TEXT(255),
                                        [Inspect Medium] TEXT(255),
                                        [Pres during run] TEXT(255),
                                        [Flowrate] TEXT(255),
                                        [Disc CupWear] TEXT(255),
                                        [Debris] TEXT(255),
                                        [Damage] TEXT(255),
                                        [Start Data Record] TEXT(255),
                                        [End Data Record] TEXT(255),
                                        [Min Velocity Record] TEXT(255),
                                        [Max Velocity Record] TEXT(255),
                                        [Size Record] TEXT(255),
                                        [Date Received Headquarters] TEXT(255))
                                        """)
        cursor.commit()
        print("Successfully created table DataQuality")
    except Exception as e:
        print(f"Unable to create table DataQuality: {str(e)}")

def add_header_mapping_to_access(cursor, mapping_content):
    """Add HeaderMapping table with structured data to Access database"""
    try:
        # Create HeaderMapping table (drop if exists)
        try:
            cursor.execute("DROP TABLE HeaderMapping")
            print("🔄 Dropped existing HeaderMapping table")
        except:
            pass  # Table doesn't exist
        
        # Create structured HeaderMapping table
        create_table_sql = """
        CREATE TABLE HeaderMapping (
            SheetName TEXT(100),
            StandardHeader TEXT(255),
            ExcelHeader TEXT(255)
        )
        """
        
        cursor.execute(create_table_sql)
        print("✅ Created HeaderMapping table in Access DB")
        
        # Parse mapping_content to extract structured data
        records = []
        current_sheet = None
        
        lines = mapping_content.split('\n')
        
        for line in lines:
            line = line.strip()
            
            # Detect sheet name
            if line.startswith('📋 SHEET:'):
                current_sheet = line.replace('📋 SHEET:', '').strip()
                continue
                
            # Skip if no current sheet
            if not current_sheet:
                continue
                
            # Parse standard header mappings (format: "Standard Header >> Excel Header")
            if '>>' in line and not line.startswith('-'):
                parts = line.split('>>')
                if len(parts) == 2:
                    standard_header = parts[0].strip()
                    excel_header = parts[1].strip()
                    
                    # Handle empty mappings
                    if excel_header == "[Empty - Will create empty column]":
                        excel_header = None
                    
                    records.append({
                        'SheetName': current_sheet,
                        'StandardHeader': standard_header,
                        'ExcelHeader': excel_header
                    })
            
            # Parse TempData mappings (format: "TempData1 << Original Column")
            elif '<<' in line and line.startswith('TempData'):
                parts = line.split('<<')
                if len(parts) == 2:
                    temp_data_name = parts[0].strip()
                    original_column = parts[1].strip()
                    
                    records.append({
                        'SheetName': current_sheet,
                        'StandardHeader': temp_data_name,
                        'ExcelHeader': original_column
                    })
            
            # Parse new columns for List of Pipe Tally
            elif current_sheet == "List of Pipe Tally" and line in ['Velocity (m/s)', 'ImgPath1', 'ImgPath2', 'Timestr']:
                records.append({
                    'SheetName': current_sheet,
                    'StandardHeader': line,
                    'ExcelHeader': None
                })
            
            # Parse original headers for unconfigured sheets
            elif line.startswith('  ') and '. ' in line and current_sheet:
                # Check if this is from an unconfigured sheet
                if any('Not Configured' in prev_line for prev_line in lines[max(0, lines.index(line)-10):lines.index(line)]):
                    # Extract header name (format: "  1. Header Name")
                    try:
                        header_part = line.split('. ', 1)
                        if len(header_part) == 2:
                            header_name = header_part[1].strip()
                            
                            records.append({
                                'SheetName': current_sheet,
                                'StandardHeader': header_name,
                                'ExcelHeader': header_name
                            })
                    except:
                        pass
        
        # Insert mapping records
        insert_sql = """
        INSERT INTO HeaderMapping (SheetName, StandardHeader, ExcelHeader) 
        VALUES (?, ?, ?)
        """
        
        for record in records:
            cursor.execute(insert_sql, (
                record['SheetName'],
                record['StandardHeader'], 
                record['ExcelHeader']
            ))
        
        print(f"✅ Added {len(records)} HeaderMapping records to Access DB")
        print(f"   📋 Structured mapping information stored in table")
        
        return True
        
    except Exception as e:
        print(f"❌ Error creating HeaderMapping table: {str(e)}")
        return False

def excel_to_access(excel_file, header_file=None, selected_headers=None, sheet_modes=None, header_mapping_content=None, use_imperial=False):
    check_List_Pipe = False
    check_List_Nominal = False
    pipeTallyColumns = []  
    nomThickColumns = []
    
    # Use selected_headers if provided, otherwise use header_file, otherwise use default values
    if selected_headers:
        pipeTallyColumns = selected_headers        # Use selected headers from UI
    elif header_file is not None:
        # Use header_file if specified
        try:
            xlsHead = pd.ExcelFile(header_file)
            for sheet_name in xlsHead.sheet_names:
                dfheader = pd.read_excel(xlsHead, sheet_name=sheet_name, header=None)
                
                if sheet_name == "List of Pipe Tally":
                    dfheader = GetHeaderColumn(dfheader)
                    pipeTallyColumns = dfheader.columns
                
                if sheet_name == "List of Nominal Wall Thickness":
                    dfheader = GetHeaderColumn(dfheader)
                    nomThickColumns = dfheader.columns
        except Exception as e:
            print(f"Warning: Error reading header file: {str(e)}")
    else:
        # Use default header file
        header_file = resource_path("resoure\\header.xlsx")
        try:
            if os.path.exists(header_file):
                xlsHead = pd.ExcelFile(header_file)
                for sheet_name in xlsHead.sheet_names:
                    dfheader = pd.read_excel(xlsHead, sheet_name=sheet_name, header=None)
                    
                    if sheet_name == "List of Pipe Tally":
                        dfheader = GetHeaderColumn(dfheader)
                        pipeTallyColumns = dfheader.columns
                    
                    if sheet_name == "List of Nominal Wall Thickness":
                        dfheader = GetHeaderColumn(dfheader)
                        nomThickColumns = dfheader.columns
        except Exception as e:
            print(f"Warning: Error reading default header file: {str(e)}")

    access_file = os.path.splitext(excel_file)[0] + ".accdb"    # Create or connect to the Access database
    if not create_access_database(access_file):                 # Create a new Access database
        print("Error: Failed to create Access database.")
        return False
    
    # # Connect to the Access database
    conn        = None
    conn_str    = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={access_file};'
    conn        = pyodbc.connect(conn_str, autocommit=False)
    cursor      = conn.cursor()
    
    # Read the Excel file and transform the data
    try:
        xls = pd.ExcelFile(excel_file)
        total_sheets = len(xls.sheet_names)
        
        # Process each sheet in the Excel file
        for i, sheet_name in enumerate(xls.sheet_names, 1):
            print(f"\n{'='*60}")
            print(f"Processing sheet: {sheet_name}")
            print(f"{'='*60}")
            
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            df = set_specific_headers(df, sheet_name)

            # Check sheet processing mode
            sheet_mode = "individual"  # default
            sheet_mappings = None
            is_nominal_wall = False
            
            if sheet_modes and sheet_name in sheet_modes:
                mode_data = sheet_modes[sheet_name]
                if isinstance(mode_data, dict):
                    # Enhanced mode data
                    sheet_mode      = mode_data.get('mode', 'individual')
                    sheet_mappings  = mode_data.get('mappings')
                    is_nominal_wall = mode_data.get('is_nominal_wall', False)
                else:
                    sheet_mode = mode_data      # Simple mode string

            if sheet_mappings:
                mapped_count = sum(1 for v in sheet_mappings.values() if v is not None)
                print(f"Sheet '{sheet_name}' has {mapped_count} mapped headers")

            # Apply header management based on sheet mode
            if sheet_mode == "standard":
                if sheet_mappings:
                    df = apply_selected_headers_to_dataframe(df, sheet_mappings, sheet_name)        # Using mappings from UI
                elif is_nominal_wall:
                    # Sheet "List of Nominal Wall Thickness" use specific headers
                    nominal_headers = [
                        "Log distance (m)",
                        "Girth weld Nr",
                        "Nominal thickness (mm)", 
                        "Joint manufacturing type",
                        "SMYS (psi)",
                        "Design Pressure (psi)",
                        "MAOP (psi)"
                    ]
                    df = apply_selected_headers_to_dataframe(df, nominal_headers, sheet_name)
                elif selected_headers:
                    df = apply_selected_headers_to_dataframe(df, selected_headers, sheet_name)      # Use normal selected headers   
            else:
                print(f"📋 Using original headers for {sheet_name}: {len(df.columns)} columns")

            df = add_erf_type(df)
   
            if sheet_name == "List of Pipe Tally":
                check_List_Pipe = True
        
                if 'isNormalERF' in df.columns:
                    df = df.drop(columns=['isNormalERF'])
        
                df = convert_data_types(df, use_imperial=use_imperial)

                # Check against `pipeTallyColumns` if defined
                if len(pipeTallyColumns) > 0:
                    message, misspelled, true_extra, missing = compare_arrays_with_alert(pipeTallyColumns, df.columns)
                    if message != 'OK':
                        if len(misspelled) > 0 or len(missing) > 0:
                            print(f"Warning: Issues with sheet '{sheet_name}':")
                            if len(misspelled) > 0:
                                print(f"  - Misspelled columns: {', '.join(misspelled)}")
                            if len(missing) > 0:
                                print(f"  - Missing columns: {', '.join(missing)}")
                        if len(true_extra) > 0:
                            print(f"Info: Extra columns in '{sheet_name}': {', '.join(true_extra)}")
    
            if sheet_name == "List of Nominal Wall Thickness":
                check_List_Nominal = True

                # Check against `nomThickColumns` if defined
                if len(nomThickColumns) > 0:
                    message, misspelled, true_extra, missing = compare_arrays_with_alert(nomThickColumns, df.columns)
                    if message != 'OK':
                        if len(misspelled) > 0 or len(missing) > 0:
                            print(f"Warning: Issues with sheet '{sheet_name}':")
                            if len(misspelled) > 0:
                                print(f"  - Misspelled columns: {', '.join(misspelled)}")
                            if len(missing) > 0:
                                print(f"  - Missing columns: {', '.join(missing)}")
                        if len(true_extra) > 0:
                            print(f"Info: Extra columns in '{sheet_name}': {', '.join(true_extra)}")
    
            # Create tables and import data
            success, df = create_access_table(cursor, sheet_name, df)
            if success:
                insert_data_to_access(cursor, sheet_name, df)
   
            progress = int((i / total_sheets) * 100)
            print(f"PROGRESS:{progress}", flush=True)
       
        create_new_tables(cursor)

        if header_mapping_content:
            add_header_mapping_to_access(cursor, header_mapping_content)

        conn.commit()
        print("Excel to Access conversion completed successfully")
        return True
    
    except Exception as e:
        print(f"Error during Excel to Access conversion: {str(e)}")
        if conn:
            conn.rollback()
        
        # Delete the created database if an error occurs
        if os.path.exists(access_file):
            try:
                if conn:
                    conn.close()
                os.remove(access_file)
                print(f"Removed incomplete database: {access_file}")
            except Exception as e2:
                print(f"Warning: Could not remove database file: {str(e2)}")
        return False
    
    finally:
        if conn:
            try:
                conn.close()
            except Exception as e:
                print(f"Warning: Error closing connection: {str(e)}")

def apply_selected_headers_to_dataframe(df, selected_headers_or_mappings, sheet_name=None):
    """Apply selected headers or UI mappings to DataFrame and add TempData after new columns"""
    try:
        current_columns = list(df.columns)
        new_df_data = {}
        used_original_columns = []

        # Check if it's a dict (mappings) or a list (headers)
        if isinstance(selected_headers_or_mappings, dict):
            ui_mappings = selected_headers_or_mappings
            
            for std_header, excel_header in ui_mappings.items():
                if excel_header:
                    print(f"  {std_header} <-- {excel_header}")
            
            # Apply mappings
            for std_header, excel_header in ui_mappings.items():
                if excel_header and excel_header in current_columns:
                    new_df_data[std_header] = df[excel_header].copy()
                    used_original_columns.append(excel_header)
                    print(f"    ✅ {std_header} <- {excel_header} (Type: {df[excel_header].dtype})")
                else:
                    new_df_data[std_header] = [None] * len(df)
                    if excel_header:
                        print(f"    ⚠️ {std_header} <- [Empty - Excel column '{excel_header}' not found]")
                    else:
                        print(f"    ⚠️ {std_header} <- [Empty - Not mapped]")          
        else:
            selected_headers = selected_headers_or_mappings
           
            for header in selected_headers:     # Apply headers
                matched_column = None
                for orig_col in current_columns:
                    if orig_col not in used_original_columns:
                        if header == orig_col or header.lower() == orig_col.lower():
                            matched_column = orig_col
                            break
                        elif header.lower() in orig_col.lower() or orig_col.lower() in header.lower():
                            matched_column = orig_col
                            break
                
                if matched_column:
                    # Copy data without converting data type
                    new_df_data[header] = df[matched_column].copy()
                    used_original_columns.append(matched_column)
                    print(f"  ✅ '{header}' <-- '{matched_column}' (Type: {df[matched_column].dtype})")
                else:
                    new_df_data[header] = [None] * len(df)
                    print(f"  ⚠️ '{header}' <-- [EMPTY - No Match Found]")
        
        # Calculate remaining columns
        remaining_columns = [col for col in current_columns if col not in used_original_columns]
        
        # Add new column only in Sheet "List of Pipe Tally"
        if sheet_name == "List of Pipe Tally":
            print(f"📌 Adding 5 new columns (Sheet: {sheet_name})...")
            new_df_data['Velocity (m/s)'] = [None] * len(df)
            new_df_data['ImgPath1'] = [None] * len(df)
            new_df_data['ImgPath2'] = [None] * len(df)
            new_df_data['Timestr'] = [None] * len(df)
            print(f"✅ Added: Velocity (m/s), ImgPath1, ImgPath2, Timestr")
        else:
            print(f"📋 Skipping new columns (Sheet: {sheet_name} - not 'List of Pipe Tally')")
        
        # Add TempData
        if remaining_columns:
            temp_counter = 1
            for orig_col in remaining_columns:
                temp_name = f"TempData{temp_counter}"
                new_df_data[temp_name] = df[orig_col].copy()  # Copy without conversion
                print(f"    📦 {temp_name} <- {orig_col} (Type: {df[orig_col].dtype})")
                temp_counter += 1
        else:
            print("📋 No remaining columns - no TempData to add")
        
        new_df = pd.DataFrame(new_df_data)  # Create DataFrame
        return new_df
        
    except Exception as e:
        print(f"❌ Error applying headers: {str(e)}")
        return df

def main():
    if len(sys.argv) < 2:
        print(f"No file path provided, using default: {excel_file}")
    else:
        excel_file = sys.argv[1]
    
    # Call the function to convert Excel to Access
    if excel_to_access(excel_file):
        print("Msg: Conversion completed successfully.")
    else:
        print("Msg: Conversion failed.")

if __name__ == "__main__":
    main()