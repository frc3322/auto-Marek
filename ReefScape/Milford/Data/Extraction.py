import sqlite3
import pandas as pd
import os

# Example usage: List of database and excel files
db_files = {
    'R1': r'C:\ScoutingData\Milford\Data\ScoutingTableRed1\match_data.db',
    'R2': r'C:\ScoutingData\Milford\Data\ScoutingTableRed2\match_data.db',
    'R3': r'C:\ScoutingData\Milford\Data\ScoutingTableRed3\match_data.db',
    'B1': r'C:\ScoutingData\Milford\Data\ScoutingTableBlu1\match_data.db',
    'B2': r'C:\ScoutingData\Milford\Data\ScoutingTableBlu2\match_data.db',
    'B3': r'C:\ScoutingData\Milford\Data\ScoutingTableBlu3\match_data.db',
}

excel_files = {
    'R1': r'C:\ScoutingData\Milford\Data\ScoutingTableRed1\outputRed1.xlsx',
    'R2': r'C:\ScoutingData\Milford\Data\ScoutingTableRed2\outputRed2.xlsx',
    'R3': r'C:\ScoutingData\Milford\Data\ScoutingTableRed3\outputRed3.xlsx',
    'B1': r'C:\ScoutingData\Milford\Data\ScoutingTableBlu1\outputBlu1.xlsx',
    'B2': r'C:\ScoutingData\Milford\Data\ScoutingTableBlu2\outputBlu2.xlsx',
    'B3': r'C:\ScoutingData\Milford\Data\ScoutingTableBlu3\outputBlu3.xlsx',
}

def db_to_excel(db_files, excel_files):
    for key in db_files:
        # Check if the Excel file already exists, and if so, delete it
        if os.path.exists(excel_files[key]):
            print(f"Deleting existing file: {excel_files[key]}")
            os.remove(excel_files[key])

        # Connect to the SQLite database
        conn = sqlite3.connect(db_files[key])
        
        # Create a cursor to interact with the database
        cursor = conn.cursor()
        
        # Table name (assuming the table name is always "MatchData")
        table_name = "MatchData"
        
        # Create an Excel writer object for each database
        with pd.ExcelWriter(excel_files[key], engine='openpyxl') as writer:
            # Read the data from the table into a pandas DataFrame
            df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
            
            # Write the DataFrame to a sheet in the Excel file
            df.to_excel(writer, sheet_name=table_name, index=False)
        
        # Close the connection
        conn.close()
        
        # Open the Excel file after saving it
        os.startfile(excel_files[key])  # This will open the file with the default associated app (Excel)

# Call the function to process all databases and save to Excel
db_to_excel(db_files, excel_files)
