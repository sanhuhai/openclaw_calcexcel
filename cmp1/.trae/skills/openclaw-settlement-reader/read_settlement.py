#!/usr/bin/env python3
"""
Read the 外包结算单测试模版.xlsx file, find the cell with content "{钉钉编号}",
and save that cell and the data after it.
"""

import pandas as pd
import os
from datetime import datetime

# Define the file path
file_path = "外包结算单测试模版.xlsx"

try:
    # Read all sheets from the Excel file
    xl = pd.ExcelFile(file_path)
    print(f"File: {file_path}")
    print(f"Sheets: {xl.sheet_names}")
    
    # Create ExcelWriter to save all sheets to one file
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"外包结算单测试模版_extracted_{timestamp}.xlsx"
    
    with pd.ExcelWriter(output_file) as writer:
        # Process each sheet
        for sheet_name in xl.sheet_names:
            print(f"\nProcessing sheet: {sheet_name}")
            
            # Read the sheet
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            print(f"Sheet shape: {df.shape}")
            
            # Search for the cell with "{钉钉编号}"
            found = False
            for row_idx, row in df.iterrows():
                for col_idx, value in enumerate(row):
                    if str(value) == "钉钉编号":
                        print(f"Found '{value}' at row {row_idx}, column {col_idx}")
                        
                        # Extract data from this cell onwards
                        # Get the remaining columns from this cell
                        remaining_columns = df.columns[col_idx:]
                        # Get the remaining rows including current row
                        # Note: row_idx is the header row, data starts from row_idx+1
                        extracted_data = df.iloc[row_idx:, col_idx:]
                        
                        # Remove rows from "合计" onwards
                        total_row_idx = None
                        for i, row in enumerate(extracted_data.iterrows()):
                            _, row_data = row
                            for value in row_data:
                                value_str = str(value).strip()
                                # Check if the value contains "合计" (with or without spaces)
                                if "合" in value_str and "计" in value_str:
                                    total_row_idx = i
                                    break
                            if total_row_idx is not None:
                                break
                        
                        if total_row_idx is not None:
                            print(f"Found '合计' at row {total_row_idx} in extracted data")
                            extracted_data = extracted_data.iloc[:total_row_idx]
                            print(f"Data shape after removing '合计' and beyond: {extracted_data.shape}")
                        
                        # Set the first row as column names and remove it from data
                        if not extracted_data.empty:
                            # Get the first row as column names
                            new_columns = extracted_data.iloc[0]
                            # Remove the first row from data
                            extracted_data = extracted_data.iloc[1:]
                            # Set the new column names
                            extracted_data.columns = new_columns
                            print("Set first row as column names and removed it from data")
                        
                        # Process merged cells: fill empty cells with previous values
                        if not extracted_data.empty:
                            # Create a copy to avoid modifying the original
                            processed_data = extracted_data.copy()
                            # Initialize variables to store previous values
                            previous_values = {}
                            
                            # Iterate through each row
                            for i, row in processed_data.iterrows():
                                # Check if this row has a钉钉编号
                                has_dingding = pd.notna(row.get('钉钉编号')) and str(row.get('钉钉编号')).strip() != ''
                                
                                if has_dingding:
                                    # Save all non-empty values for future rows
                                    for col in processed_data.columns:
                                        if pd.notna(row[col]) and str(row[col]).strip() != '':
                                            previous_values[col] = row[col]
                                else:
                                    # Fill empty cells with previous values
                                    for col in processed_data.columns:
                                        if not (pd.notna(row[col]) and str(row[col]).strip() != ''):
                                            if col in previous_values:
                                                processed_data.at[i, col] = previous_values[col]
                            
                            # Replace with processed data
                            extracted_data = processed_data
                            print("Processed merged cells: filled empty cells with previous values")
                        
                        # Reset index to ensure all rows are included
                        extracted_data = extracted_data.reset_index(drop=True)
                        
                        print(f"Extracted data shape: {extracted_data.shape}")
                        print("Extracted data:")
                        print(extracted_data)
                        
                        # Save extracted data to the current ExcelWriter with original sheet name
                        extracted_data.to_excel(writer, sheet_name=sheet_name, index=False)
                        print(f"\nSuccessfully saved extracted data to sheet '{sheet_name}' in {output_file}")
                        found = True
                        break
                if found:
                    break
            
            if not found:
                print(f"'钉钉编号' not found in sheet {sheet_name}")
    
    print(f"\nAll sheets processed. Final output file: {output_file}")
            
except Exception as e:
    print(f"Error processing file: {str(e)}")