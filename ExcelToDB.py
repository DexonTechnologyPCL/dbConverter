import pandas as pd
import numpy as np
import sqlite3
import sys
import os
import re

def set_specific_headers(df):
    """Set specific columns as headers for the DataFrame."""
    df = df.copy()  # Ensure we are working on a copy of the DataFrame

    headers = df.iloc[1].tolist()       # Create a list for new headers
    headers[0] = "Log distance [m]"     # Set specific headers for columns 0 and 25
    headers[25] = "Comments"
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
    df["Log distance [m]"] = pd.to_numeric(df["Log distance [m]"], errors='coerce').round(3)
    
    # Convert specific columns to numeric and round decimal 
    columns_three_decimal = [
        "Altitude [m]", "Joint / component length [m]", 
        "Abs. Dist. to upstream weld [m]", "Remaining thickness [mm]"
    ]
    for col in columns_three_decimal:
        df[col] = pd.to_numeric(df[col], errors='coerce').round(3).apply(lambda x: f"{x:.3f}" if pd.notnull(x) else None)

    columns_two_decimal = ["Nominal Internal diameter [mm]", "Max. depth [mm]"]
    for col in columns_two_decimal:
         df[col] = pd.to_numeric(df[col], errors='coerce').apply(custom_round_two_decimal).apply(lambda x: f"{x:.2f}" if pd.notnull(x) else None)

    # Special handling for Max. depth [%]
    df["Max. depth [%]"] = df["Max. depth [%]"].apply(custom_round_max_depth)

    numeric_columns_to_round = ["Length [mm]", "Width [mm]"]
    for col in numeric_columns_to_round:
        df[col] = pd.to_numeric(df[col], errors='coerce').apply(custom_round).apply(lambda x: str(int(x)) if pd.notnull(x) else None)

    for col in df.columns:
        if col not in ["Log distance [m]"] + columns_three_decimal + columns_two_decimal + numeric_columns_to_round + ["Max. depth [%]"]:
            df[col] = df[col].astype(str).replace({'nan': None, 'None': None, '': None}).where(pd.notnull(df[col]), None)

    return df

def add_erf_type(df):
    """Add ERF flag based on ERF column."""
    if 'ERF (Modified)' in df.columns and 'ERF (metal loss)' in df.columns:
        position = df.columns.get_loc('ERF (Modified)')
        df['ERF'] = df.apply(lambda row: row['ERF (Modified)'] if pd.notnull(row['ERF (Modified)']) else row['ERF (metal loss)'], axis=1)
        df.insert(position, 'ERF', df.pop('ERF'))
        df['isNormalERF'] = df['ERF (metal loss)'].notnull()
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
        df = convert_data_types(df)
        df = add_erf_type(df) # Add the ERF flag
        df = df.drop(columns=['isNormalERF'])  # Drop the isNormalERF column before saving
        df.to_sql(sheet_name, conn, if_exists='replace', index=False) # Insert data into SQLite in bulk
    
    # Commit and close the connection
    conn.commit()
    conn.close()
    
    print(f"Excel file {excel_file} has been successfully converted to {db_file}.")
    return True

if __name__ == "__main__":

    # folder_path = "D:/"
    # for filename in os.listdir(folder_path):
    #     if filename.endswith('.xlsx'):
    #         if  excel_to_sqlite(folder_path + filename):
    #             print("Conversion completed successfully.")
    #         else:
    #             print("Conversion completed with errors.")

    excel_file = "D:\dbtest\YPF 8in x 10km Jet Fuel Pipeline Poliducto La Matanza to Aeroplanta Ezeiza UTMC List Pipe Tally_Rev.03.xlsx"
    if  excel_to_sqlite(excel_file):
        print("Conversion completed successfully.")
    else:
        print("Conversion completed with errors.")
