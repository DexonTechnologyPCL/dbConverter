import pandas as pd
import numpy as np
import sqlite3
import sys
import os
import re
import openpyxl

def set_specific_headers(df):
    """Set specific columns as headers for the DataFrame, preserving original structure."""
    df = df.copy()  # Work on a copy of the DataFrame

    # Extract the first two rows which may contain header information
    first_row = df.iloc[0].fillna(method='ffill').tolist()
    second_row = df.iloc[1].fillna(method='ffill').tolist()

    # Combine the rows to create the headers
    headers = []
    for i in range(len(second_row)):
        if pd.notna(second_row[i]) and second_row[i].strip() != '':
            headers.append(second_row[i].strip())
        elif pd.notna(first_row[i]) and first_row[i].strip() != '':
            headers.append(first_row[i].strip())
        else:
            headers.append(f"Unnamed_{i}")  # Assign a placeholder for truly empty headers

    if "Comments" not in headers and len(headers) > 1:
        headers[-1] = "Comments"  # Forcefully assign if missing

    # Ensure all headers are strings and unique
    unique_headers = []
    for i, header in enumerate(headers):
        header_str = str(header)  # Convert to string
        if header_str in unique_headers:
            unique_headers.append(f"{header_str}_{i}")
        else:
            unique_headers.append(header_str)
    
    df.columns = unique_headers
    df = df.drop([0, 1]).reset_index(drop=True)  # Drop the first two rows after setting headers and reset index
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
            return str(float_x)  # Keep original precision for 1 or 0 decimal places
    except ValueError:
        return str(x)  # Keep as is if it's not a number
    
    
def custom_round_two_decimal(x):
    """Custom rounding function to two decimal places."""
    if pd.isna(x):
        return pd.NA
    rounded = round(x * 100) / 100
    if round(x, 3) - rounded >= 0.001:
        rounded += 0.01
    return round(rounded, 2)

def convert_data_types(df):
    """Convert data types of specific columns."""

    if "Log distance [m]" in df.columns:
        df["Log distance [m]"] = pd.to_numeric(df["Log distance [m]"], errors='coerce').round(3)
    
    # List of columns to process for three decimal places
    columns_three_decimal = ["Altitude [m]", "Joint / component length [m]", "Abs. Dist. to upstream weld [m]", "Remaining thickness [mm]"]
    
    for col in columns_three_decimal:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').round(3).apply(lambda x: f"{x:.3f}" if pd.notnull(x) else None)

    columns_two_decimal = ["Nominal Internal diameter [mm]", "Max. depth [mm]"]
    
    for col in columns_two_decimal:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').apply(custom_round_two_decimal).apply(lambda x: f"{x:.2f}" if pd.notnull(x) else None)

    if "Max. depth [%]" in df.columns:
        df["Max. depth [%]"] = df["Max. depth [%]"].apply(custom_round_max_depth)

    numeric_columns_to_round = ["Length [mm]", "Width [mm]"]
    for col in numeric_columns_to_round:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').apply(custom_round).apply(lambda x: str(int(x)) if pd.notnull(x) else None)
    for col in df.columns:
        if col not in ["Log distance [m]"] + columns_three_decimal + columns_two_decimal + numeric_columns_to_round + ["Max. depth [%]"]:
            df[col] = df[col].astype(str).replace({'nan': None, 'None': None, '': None}).where(pd.notnull(df[col]), None)
 
    return df

def add_erf_type(df):
    """Add ERF flag based on ERF column."""
    if 'ERF (Modified)' in df.columns and 'ERF (metal loss)' in df.columns:
        #It finds the position of 'ERF (Modified)' column.
        position = df.columns.get_loc('ERF (Modified)')
        #Creates a new 'ERF' column, using 'ERF (Modified)' values if they're not null, otherwise using 'ERF (metal loss)' values.
        df['ERF'] = df.apply(lambda row: row['ERF (Modified)'] if pd.notnull(row['ERF (Modified)']) else row['ERF (metal loss)'], axis=1)
        #Inserts the new 'ERF' column at the position of 'ERF (Modified)'.
        df.insert(position, 'ERF', df.pop('ERF'))
        #Creates an 'isNormalERF' column, which is True where 'ERF (metal loss)' is not null.
        df['isNormalERF'] = df['ERF (metal loss)'].notnull()
        #Drops the original 'ERF (Modified)' and 'ERF (metal loss)' columns.
        df = df.drop(columns=['ERF (Modified)', 'ERF (metal loss)'])
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

def excel_to_sqlite(excel_file):
    # Check if the Excel file exists
    if not os.path.exists(excel_file):
        print(f"Error: The file {excel_file} does not exist.")
        return False

    # Create a connection to the SQLite database
    db_file = os.path.splitext(excel_file)[0] + ".db"
    conn = sqlite3.connect(db_file)
    
    # Read the Excel file
    xls = pd.ExcelFile(excel_file)
    # Loop through each sheet in the Excel file
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
     
        # Set specific columns as headers
        df = set_specific_headers(df)
 
        # if (sheet_name == "List of Pipe Tally") :
        df = add_erf_type(df) # Add the ERF flag

        # Drop the isNormalERF column if it exists
        if 'isNormalERF' in df.columns:
            df = df.drop(columns=['isNormalERF'])
            
        df.to_sql(sheet_name, conn, if_exists='replace', index=False) # Insert data into SQLite in bulk
    
    # Commit and close the connection
    conn.commit()
    conn.close()
    
    print(f"Processing sheet: {sheet_name}")
    print(f"Excel file {excel_file} has been successfully converted to {db_file}.")
    return True

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

def compare_arrays_with_alert(array1, array2):
    # Find elements in array1 that are not in array2
    missing_in_array2 = set(array1) - set(array2)
    
    # Find elements in array2 that are not in array1
    missing_in_array1 = set(array2) - set(array1)
    
    if missing_in_array2 or missing_in_array1:
        print("ALERT: The arrays are different!")
        
        if missing_in_array2:
            print(f"Elements in array1 but not in array2: {', '.join(missing_in_array2)}")
        
        if missing_in_array1:
            print(f"Elements in array2 but not in array1: {', '.join(missing_in_array1)}")
    else:
        print("The arrays contain the same elements.")
    
    return missing_in_array2, missing_in_array1
        
if __name__ == "__main__":

    # folder_path = "D:/"
    # for filename in os.listdir(folder_path):
    #     if filename.endswith('.xlsx'):
    #         if  excel_to_sqlite(folder_path + filename):
    #             print("Conversion completed successfully.")
    #         else:
    #             print("Conversion completed with errors.")

    excel_file = "D:\dbtest\PlusPetrol_Argentina_12inch_82km_UTMC List of Pipe Tally_Rev01 1.xlsx"
    if  excel_to_sqlite(excel_file):
        print("Conversion completed successfully.")
    else:
        print("Conversion completed with errors.")

    
