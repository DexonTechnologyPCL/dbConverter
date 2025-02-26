import pandas as pd
import numpy as np
import pyodbc
import os
import sys
import openpyxl
import win32com.client
import re

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
        if "Comments" not in headers and len(headers) > 1:
            headers[-1] = "Comments"        # Forcefully assign if missing
    
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

def custom_round(x):
    """Custom rounding function: round down at 0.49 and up at 0.5."""
    if pd.isna(x):
        return pd.NA
    return round(x)

def custom_round_max_depth(x):
    if pd.isna(x):
        return None
    try:
        float_x = float(x)
        # Round only if the value has more than 1 decimal place
        if abs(float_x - round(float_x, 1)) > 0.00001:
            return str(round(float_x))
        else:
            return str(float_x) # Keep original precision for 1 or 0 decimal places
    except ValueError:
        return str(x)           # Keep as is if it's not a number
    
def custom_round_two_decimal(x):
    """Custom rounding function to two decimal places."""
    if pd.isna(x):
        return pd.NA
    rounded = round(x * 100) / 100
    if round(x, 3) - rounded >= 0.001:
        rounded += 0.01
    return round(rounded, 2)

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

def get_access_data_type(df_column):
    """ Assign appropriate data types for columns in Access"""

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
            data_type = get_access_data_type(df[col])
            columns.append(f"[{clean_col}] {data_type}")
        
        # Generate SQL statement
        create_table_sql = f"CREATE TABLE [{table_name}] ({', '.join(columns)})"
        
        # Execute table creation
        print(f"Creating table with SQL: {create_table_sql}")
        cursor.execute(create_table_sql)
        cursor.commit()
        
        # Rename columns in the DataFrame
        df.columns = [column_map[col] for col in df.columns]
        
        print(f"Table '{table_name}' created successfully")
        return True, df

    # Method 2: Set all columns as TEXT
    except Exception as e:
        print(f"Error creating table: {str(e)}")
        print(f"SQL: {create_table_sql}")
        
        try:
            columns = []
            for col in df.columns:
                clean_col = column_map[col]
                columns.append(f"[{clean_col}] TEXT(255)")
            
            create_table_sql = f"CREATE TABLE [{table_name}] ({', '.join(columns)})"
            print(f"Trying alternative approach with SQL: {create_table_sql}")
            cursor.execute(create_table_sql)
            cursor.commit()
            
            # Rename columns in the DataFrame table
            df.columns = [column_map[col] for col in df.columns]
            
            print(f"Table '{table_name}' created successfully with alternate method")
            return True, df
        except Exception as e2:
            print(f"Error with alternative approach: {str(e2)}")
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
        
        print(f"Insert SQL: {sql_insert}")
        
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
                    print(f"Row values: {values}")
                    skipped_count += 1
                    continue
            
            cursor.commit()
            print(f"Inserted rows {start_idx+1} to {end_idx} of {total_rows}")
        
        print(f"Total rows inserted: {inserted_count}, skipped: {skipped_count}")
        return True
    
    except Exception as e:
        print(f"Error in insert_data_to_access: {str(e)}")
        raise

def create_access_database(file_path):
    """Create a new Access database"""

    # Delete the existing file if it exists
    if os.path.exists(file_path):
        try:
            os.remove(file_path)
            print(f"Removed existing database: {file_path}")
        except Exception as e:
            print(f"Warning: Could not remove existing file: {str(e)}")
            return False
    
    try:
        # Method 1: Use ADOX
        cat = win32com.client.Dispatch('ADOX.Catalog')
        conn_str = f'Provider=Microsoft.ACE.OLEDB.12.0;Data Source={file_path};'
        cat.Create(conn_str)
        cat = None
        print(f"Created database successfully using ADOX: {file_path}")
        return True
    except Exception as e:
        print(f"Error creating Access database: {str(e)}")
        
        # Method 2: Create an empty database
        try:
            # Create an empty Access file
            conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={file_path};'
            conn = pyodbc.connect(conn_str, autocommit=True)
            conn.close()

            print(f"Created database using alternative method: {file_path}")
            return True

        except Exception as e2:
            print(f"Error with alternative approach: {str(e2)}")
            return False

def convert_data_types(df):
    """Convert data types of specific columns."""

    if "Log distance (m)" in df.columns:
        df["Log distance (m)"] = pd.to_numeric(df["Log distance (m)"], errors='coerce').round(3)
    
    # List of columns to process for three decimal places
    columns_three_decimal = ["Altitude (m)", "Joint / component length [m]", "Abs Dist to upstream weld (m)", "Remaining thickness (mm)"]
    
    for col in columns_three_decimal:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').round(3).apply(lambda x: f"{x:.3f}" if pd.notnull(x) else None)

    columns_two_decimal = ["Nominal Internal diameter (mm)", "Max depth (mm)"]
    
    for col in columns_two_decimal:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').apply(custom_round_two_decimal).apply(lambda x: f"{x:.2f}" if pd.notnull(x) else None)

    if "Max depth (%)" in df.columns:
        df["Max depth (%)"] = df["Max depth (%)"].apply(custom_round_max_depth)

    numeric_columns_to_round = ["Length (mm)", "Width (mm)"]

    for col in numeric_columns_to_round:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').apply(custom_round).apply(lambda x: str(int(x)) if pd.notnull(x) else None)

    for col in df.columns:
        if col not in ["Log distance (m)"] + columns_three_decimal + columns_two_decimal + numeric_columns_to_round + ["Max. depth (%)"]:
            df[col] = df[col].astype(str).replace({'nan': None, 'None': None, '': None}).where(pd.notnull(df[col]), None)
 
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

def excel_to_access(excel_file, header_file=None):
    check_List_Pipe = False
    check_List_Nominal = False
    pipeTallyColumns = []  
    nomThickColumns = []
    
    # Use `header_file` if specified; otherwise, use default values
    if header_file is None:
        header_file = resource_path("resoure\\header.xlsx")
    
    # Verify the source Excel file
    if not os.path.exists(excel_file):
        print(f"Error: The file {excel_file} does not exist.")
        return False
    
    # Create or connect to the Access database
    access_file = os.path.splitext(excel_file)[0] + ".accdb"
    
    # Create a new Access database
    if not create_access_database(access_file):
        print("Error: Failed to create Access database.")
        return False
    
    # Connect to the Access database
    conn = None
    try:
        conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={access_file};'
        conn = pyodbc.connect(conn_str, autocommit=False)
        cursor = conn.cursor()
    except Exception as e:
        print(f"Error connecting to Access database: {str(e)}")
        return False
    
    # Read the Excel file and transform the data
    try:
        xls = pd.ExcelFile(excel_file)
        total_sheets = len(xls.sheet_names)
        
        try:
            # Read the header file if it exists
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
            print(f"Warning: Error reading header file: {str(e)}")
        
        # Process each sheet in the Excel file
        for i, sheet_name in enumerate(xls.sheet_names, 1):
            print(f"Processing sheet: {sheet_name}")
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            df = set_specific_headers(df, sheet_name)
            
            if sheet_name == "List of Pipe Tally":
                check_List_Pipe = True
                df = add_erf_type(df)
                
                if 'isNormalERF' in df.columns:
                    df = df.drop(columns=['isNormalERF'])
                
                df = convert_data_types(df)

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
        
        # Check if all required sheets are present
        if not check_List_Pipe:
            print("Warning: The sheet 'List of Pipe Tally' is missing")
        
        if not check_List_Nominal:
            print("Warning: The sheet 'List of Nominal Wall Thickness' is missing")
        
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

def main():
    if len(sys.argv) < 2:
        excel_file = "D:\\PlusPetrol_Test.xlsx"
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