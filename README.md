# Inventory Data Processing Script

## Overview
This script automates the process of cleaning, organizing, and summarizing inventory data extracted from Excel spreadsheets. It is designed to help different departments efficiently manage their inventory by filtering, parsing, and aggregating data while ensuring proper formatting for further analysis.

## Features
- **Filters Data**: Removes records where the 'Stored / frozen on' date falls in the current year.
- **Extracts and Standardizes Units**: Parses `Units remarks` to extract quantity and unit information.
- **Determines Final Amounts and Units**: Prioritizes stock volume and weight to ensure consistent measurements.
- **Aggregates Inventory Data**: Groups records by `Stock name` and calculates total quantities per unit type.
- **Computes Total Prices**: Estimates total price based on quantity and unit information.
- **Applies Conditional Formatting**: Highlights rows where `Total Price` is zero for easy identification.
- **Exports to Excel**: Saves the processed and formatted data to an output Excel file.

## Dependencies
Ensure the following Python libraries are installed before running the script:

```bash
pip install pandas openpyxl
```

## Usage
1. **Prepare the Input File**:
   - Place the Excel file in the `vivo` or `vitro` directory.
   - The expected input should contain columns: `Stored / frozen on`, `Units remarks`, `Stock volume`, `Stock weight`, `Volume units`, `Weight units`, `Price`, `Stock name`, `Manufacturer`, `Catalog no.`.

2. **Run the Script**:

```bash
python main.py
```

3. **Check the Output**:
   - The processed file will be saved in the `vivo_results` or `vitro_results` directory as `[inventory_type]_results.xlsx`.
   - The output file will include cleaned and formatted data, with grouped totals and conditional formatting applied.

## File Structure
```
project_directory/
│-- vivo/
│   ├── vivo_culture.xlsx  # Input file
│-- vivo_results/
│   ├── Culture_results.xlsx  # Processed output file
│-- main.py  # Script for processing inventory data
│-- README.md  # Documentation
```




