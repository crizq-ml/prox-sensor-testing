import pandas as pd
import os
import csv
from typing import Dict, List, Tuple

# Configuration
EXCEL_FILE = 'Prox Sensor Calibration Grounding Data Only (Trial 2).xlsx'
COLUMN_INDEX = 0
OUTPUT_CSV = 'grounding_test_data_trial2.csv'

# Define the text keys to look for
KEY_TEXTS = [
    "PHASE_3_STARTUP_OFFSET value (decimal): ",         # Will be value_1
    "use_3 average at float: ",                         # Will be value_2
    "use_3 average at detected: ",                      # Will be value_3
    "PHASE_3_STARTUP_THRESHOLD: ",                      # Will be value_4
    "Calculated PHASE_3_STARTUP_CONFIG: "               # Will be value_5
]

OUTPUT_HEADERS = ['Serial Number']
OUTPUT_HEADERS.extend(KEY_TEXTS)

# Extraction Functions
def read_and_extract_excel_file(file_path: str) -> Tuple[List[str], Dict[str, List]]:
    """
    Reads an Excel file, processes every sheet, and extracts the first column's data.

    Returns:
        A tuple containing:
        1. A list of all sheet names discovered.
        2. A dictionary mapping sheet names to the extracted data (List of values).
    """
    
    # Use pandas ExcelFile object to get sheet names
    try:
        xls = pd.ExcelFile(file_path)
    except FileNotFoundError:
        print(f"Error: File not found at {file_path}")
        return [], {}
    except Exception as e:
        print(f"An error occurred reading the Excel file: {e}")
        return [], {}
        
    sheet_names = xls.sheet_names
    extracted_data_by_sheet = {}
    
    # Iterate through every sheet name found in the file
    for sheet_name in xls.sheet_names:
        
        # Read the data from the current sheet into a DataFrame
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        
        # Check if the DataFrame is empty or doesn't have the desired column
        if df.empty or COLUMN_INDEX >= df.shape[1]:
            print(f"Skipping sheet: '{sheet_name}' is empty or too small.")
            continue
            
        # Extract the data from the specified column index (e.g., column 0/A)
        column_data = df.iloc[:, COLUMN_INDEX].astype(str).tolist() # Convert to string for consistent searching
        
        # Store the list of data, keyed by the sheet name
        extracted_data_by_sheet[sheet_name] = column_data
        
    return sheet_names, extracted_data_by_sheet

def process_sheet(sheet_name: str, data: List[str]):
    """
    Takes the extracted data for one sheet, looks for 5 key values 
    based on text matching, and prints them.
    """
    print(f"\n **Processing Sheet: {sheet_name}**")
    
    found_values = {}
    
    # Loop through the list of strings (data) for the current sheet
    for i, item in enumerate(data):
        # Look for the 5 key texts in the current item
        for key in KEY_TEXTS:
            if key in item:
                # Assuming the key text is in column A, and the value we want 
                # is the remainder of the cell OR is in the cell below.
                
                # --- Scenario 1: Value is in the same cell after the key ---
                # This tries to strip the key and trim whitespace
                value = item.replace(key, '').strip() 
                
                # Assign the value to the key's label (e.g., "Test Date:" -> "value_1")
                found_values[key] = value
                
                # Print immediately for verification
                print(f"    - Found **{key}** -> Value: {value}")
                
    # Optional: Check if all 5 values were found
    if len(found_values) < len(KEY_TEXTS):
        print(f"    - Warning: Only {len(found_values)}/{len(KEY_TEXTS)} values were found.")
    
    # Return the dictionary of found values for later use (e.g., writing to CSV)
    return found_values

def write_to_csv(filename, data_rows):
    """
    Writes a list of data rows to a specified CSV file.

    Args:
        filename (str): The name of the CSV file to create/write to.
        data_rows (list of lists): The data, where each inner list
                                   represents a row (3 values).
    """
    try:
        # Open the file in write mode ('w') with a newline='' argument
        # to prevent extra blank rows on Windows.
        with open(filename, 'w', newline='') as csvfile:
            # Create a writer object
            csv_writer = csv.writer(csvfile)

            # Write a header row (optional)
            csv_writer.writerow(OUTPUT_HEADERS)

            # Write all data rows
            csv_writer.writerows(data_rows)

        print(f"Successfully wrote {len(data_rows)} rows to **{filename}**.")

    except Exception as e:
        print(f"An error occurred: {e}")

def main():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(current_dir, EXCEL_FILE)
    final_csv_rows = []

    # Read & extract data
    all_sheet_names, all_sheet_data = read_and_extract_excel_file(file_path)

# --- New Logic: Loop through each sheet, process, and collect data ---
    print("\n--- Starting Detailed Sheet Processing ---")
    
    # all_sheet_data is a dictionary where keys are sheet names and values are the List[str] data
    for sheet_name, data_array in all_sheet_data.items():
        # 1. Call the function. It returns a dictionary of found values.
        found_data_dict = process_sheet(sheet_name, data_array)
        
        # 2. Start the row with the sheet name (used as 'Serial Number' column)
        output_row = [sheet_name]
        
        # 3. Iterate through the KEYS in the pre-defined order (KEY_TEXTS)
        #    and append the corresponding value from the dictionary.
        for key in KEY_TEXTS:
            # We use .get() to safely retrieve the value, defaulting to 'N/A' if the key wasn't found
            # in this particular sheet's data.
            value = found_data_dict.get(key, 'N/A')
            output_row.append(value)
            
        # 4. Add the complete structured row to the final list
        final_csv_rows.append(output_row)
        
    print("\n--- Summary of Extracted Data ---")
    for sheet, data_array in all_sheet_data.items():
        print(f"\nSheet **'{sheet}'** (Total Rows: {len(data_array)}):")
        print(f"  First 5 Entries: {data_array[:5]}")
    
    # --- Final Step: Function Call to write data to CSV ---
    output_filename = OUTPUT_CSV
    write_to_csv(output_filename, final_csv_rows) 
    # The final_csv_rows list is correctly structured as a list of lists:
    # [['Sheet1', 'val1', 'val2', 'val3', 'val4', 'val5'], ['Sheet2', 'val1', ...], ...]
    
if __name__ == '__main__':

    main()
