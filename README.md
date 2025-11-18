# Excel Test Data Extractor & Compiler

## 💡 Repository Summary

This repository contains a simple, yet robust **Python script** designed to automate the extraction of specific data points from multi-tab Excel workbooks. It is optimized for processing sensor test data where results are isolated in the first column of each sheet.

* **Primary Goal:** Consolidate data from dozens of Excel sheets into one structured CSV file.

* **Key Feature:** Uses sheet names (e.g., unit serial numbers) as the primary identifier for each row in the final output.

* **Extraction Method:** Locates and extracts 5 critical values by matching predefined text keys within the first column of the Excel sheets.

## 📄 Overview

This script processes a multi-tab Microsoft Excel file containing proximity sensor test data. It automates the extraction of five specific key performance indicator (KPI) values from the **first column (Column A)** of every sheet/tab and compiles the results into a single, structured CSV file.

Each sheet in the input Excel file is treated as a record for one tested unit, with the **sheet name used as the Serial Number** in the output.

## 🛠️ Prerequisites

To run this script, you must have **Python 3** installed, along with the following libraries:

### 1. Install Dependencies

Install the required libraries using pip:

```bash
pip install pandas openpyxl
```

## ⚙️ Configuration

Before running, ensure your input file and key phrases match the script's configuration:

### 1. Input File

* Place your Excel file, named **`Proximity Sensor Test - Sample units.xlsx`**, in the same directory as the Python script.

### 2. Extracted Data Keys (`KEY_TEXTS`)

The script searches for these specific text strings in **Column A** of each sheet. The value immediately following the text (in the same cell, after the colon) is extracted.

| Output Column Header | Source Text Key | 
| :--- | :--- | 
| **Serial Number** | *(Sheet Name)* | 
| PHASE_3_STARTUP_OFFSET value (decimal): | `PHASE_3_STARTUP_OFFSET value (decimal): ` | 
| use_3 average at float: | `use_3 average at float: ` | 
| use_3 average at detected: | `use_3 average at detected: ` | 
| PHASE_3_STARTUP_THRESHOLD: | `PHASE_3_STARTUP_THRESHOLD: ` | 
| Calculated PHASE_3_STARTUP_CONFIG: | `Calculated PHASE_3_STARTUP_CONFIG: ` | 

## ▶️ How to Run the Script

1. Ensure the prerequisite libraries are installed.

2. Place the Excel file in the same directory as the Python script.

3. Open your terminal or command prompt, navigate to the script's directory, and execute:

``` bash
  python your_script_name.py
```


*(Replace `your_script_name.py` with the actual name of your Python file).*

## 📤 Output

The script will generate a new file named **`extracted_sensor_data.csv`** in the same directory.

### Output CSV Structure

The CSV file will contain one data row for every sheet found in the input Excel file. If a specific key text is **not found** in a sheet, the corresponding cell in the CSV output will be filled with the string **`N/A`**.

| **Serial Number** | **PHASE_3_STARTUP_OFFSET value (decimal):** | **use_3 average at float:** | **...** | 
| :--- | :--- | :--- | :--- | 
| `Sheet_Name_1` | `1234` | `5.67` | ... | 
| `Sheet_Name_2` | `2345` | `N/A` | ... | 
