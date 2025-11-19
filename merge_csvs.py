import pandas as pd
import os

# --- Configuration ---
# List the names of your input CSV files
input_files = ['day1.0_extracted_sensor_data.csv', 'day1.1_extracted_sensor_data.csv', 'day2.0_extracted_sensor_data.csv']

# Name for the merged output file
output_file = 'merged_sensor_data.csv'

# List to hold the dataframes read from each file
all_data = []

# --- Main Logic ---

print(f"Starting merge of {len(input_files)} CSV files...")

# Loop through the list of files
for file_name in input_files:
    if os.path.exists(file_name):
        print(f"Reading file: {file_name}")
        
        # Read the CSV file into a pandas DataFrame
        # The 'header=0' tells pandas that the first row is the header
        try:
            df = pd.read_csv(file_name, header=0)
            all_data.append(df)
        except Exception as e:
            print(f"Error reading {file_name}: {e}")
    else:
        print(f"Warning: File not found: {file_name} - Skipping.")

# Check if any data was successfully loaded
if all_data:
    # Concatenate all DataFrames in the list into a single DataFrame
    # This automatically aligns columns based on the header names
    merged_df = pd.concat(all_data, ignore_index=True)
    
    # Save the resulting DataFrame to a new CSV file
    # 'index=False' prevents pandas from writing the DataFrame index as a column
    # The header will be included by default
    merged_df.to_csv(output_file, index=False)
    
    print(f"\nSuccessfully merged {len(all_data)} files into {output_file}")
    print(f"Total rows in the merged file (excluding header): {len(merged_df)}")
else:
    print("\nNo data was successfully loaded. The merged file was not created.")