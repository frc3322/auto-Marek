import sqlite3
import pandas as pd
import os
import time  # For adding a small delay

# Example usage: List of database and excel files
db_files = {
    'R1': r'C:\ScoutingData\Milford\Data\ScoutingTableRed1\match_data.db',
    'R2': r'C:\ScoutingData\Milford\Data\ScoutingTableRed2\match_data.db',
    'R3': r'C:\ScoutingData\Milford\Data\ScoutingTableRed3\match_data.db',
    'B1': r'C:\ScoutingData\Milford\Data\ScoutingTableBlu1\match_data.db',
    'B2': r'C:\ScoutingData\Milford\Data\ScoutingTableBlu2\match_data.db',
    'B3': r'C:\ScoutingData\Milford\Data\ScoutingTableBlu3\match_data.db',
}

# Excel files for individual databases
excel_files = {
    'R1': r'C:\ScoutingData\Milford\Data\ScoutingTableRed1\outputRed1.xlsx',
    'R2': r'C:\ScoutingData\Milford\Data\ScoutingTableRed2\outputRed2.xlsx',
    'R3': r'C:\ScoutingData\Milford\Data\ScoutingTableRed3\outputRed3.xlsx',
    'B1': r'C:\ScoutingData\Milford\Data\ScoutingTableBlu1\outputBlu1.xlsx',
    'B2': r'C:\ScoutingData\Milford\Data\ScoutingTableBlu2\outputBlu2.xlsx',
    'B3': r'C:\ScoutingData\Milford\Data\ScoutingTableBlu3\outputBlu3.xlsx',
}

# Output file for the combined data
combined_output_file = r'C:\ScoutingData\Milford\Data\combined_output.xlsx'

def db_to_excel(db_files, excel_files, combined_output_file):
    # Initialize an empty DataFrame to hold all combined data
    combined_df = pd.DataFrame()

    for key in db_files:
        # Check if the individual Excel file already exists, and if so, delete it
        if os.path.exists(excel_files[key]):
            print(f"Deleting existing file: {excel_files[key]}")
            os.remove(excel_files[key])

        # Connect to the SQLite database
        conn = sqlite3.connect(db_files[key])
        
        # Table name (assuming the table name is always "MatchData")
        table_name = "MatchData"
        
        # Read the data from the table into a pandas DataFrame
        
        df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
        
        # Optionally, add a column indicating the source (for identification)
        df['Source'] = key
        
        # Append the data to the combined DataFrame
        combined_df = pd.concat([combined_df, df], ignore_index=True)

        # Create an Excel writer object for each individual database
        with pd.ExcelWriter(excel_files[key], engine='openpyxl') as writer:
            # Write the DataFrame to a sheet in the Excel file
            df.to_excel(writer, sheet_name=table_name, index=False)
        
        # Close the connection
        conn.close()

        # Open the individual Excel file after saving it
        os.startfile(excel_files[key])  # This will open the file with the default associated app (Excel)

    # Check if the combined Excel file exists, and if so, delete it
    if os.path.exists(combined_output_file):
        print(f"Deleting existing file: {combined_output_file}")
        os.remove(combined_output_file)

    # Write the combined data to a single Excel file (combined sheet)
    with pd.ExcelWriter(combined_output_file, engine='openpyxl') as writer:
        # Write the combined DataFrame to a sheet in the combined Excel file
        combined_df.to_excel(writer, sheet_name='CombinedData', index=False)

    # Adding a delay before attempting to open the combined Excel file
    time.sleep(1)  # Wait for 1 second to ensure the file is fully written

    # Open the combined Excel file after saving it
    if os.path.exists(combined_output_file):
        os.startfile(combined_output_file)  # This will open the file with the default associated app (Excel)
    else:
        print(f"Error: The file {combined_output_file} was not created successfully.")

# Call the function to process all databases and save to individual and combined Excel files
db_to_excel(db_files, excel_files, combined_output_file)
