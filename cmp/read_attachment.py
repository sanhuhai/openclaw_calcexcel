#!/usr/bin/env python3
"""
Read and display the content of the "土建+杆塔" sheet in the attachment file, and output as md and json files.
"""

import pandas as pd
import json

# Define the attachment file path
attachment_file = "附件五：中国铁塔云南公司2024-2025年土建杆塔施工集中采购项目施工模块安全生产费明细表.xlsx"

try:
    # Read the "土建+杆塔" sheet from the Excel file
    sheet_name = "土建+杆塔"
    df = pd.read_excel(attachment_file, sheet_name=sheet_name)
    
    # Print file information
    print(f"File: {attachment_file}")
    print(f"Sheet: {sheet_name}")
    print(f"Shape: {df.shape}")
    print(f"Columns: {list(df.columns)}")
    print("\nContent:")
    
    # Print the content
    print(df)
    
    # Also print each row with index
    print("\nDetailed content by row:")
    for i, row in df.iterrows():
        print(f"\nRow {i}:")
        for j, value in enumerate(row):
            print(f"  {df.columns[j]}: {value}")
    
    # Output as markdown file
    md_file = "附件五_土建+杆塔内容.md"
    with open(md_file, 'w', encoding='utf-8') as f:
        # Write header
        f.write(f"# 附件五：中国铁塔云南公司2024-2025年土建杆塔施工集中采购项目施工模块安全生产费明细表\n\n")
        f.write(f"## 土建+杆塔表格\n\n")
        
        # Write table
        f.write("| " + " | ".join(df.columns) + " |\n")
        f.write("| " + " | ".join(["---"] * len(df.columns)) + " |\n")
        
        for _, row in df.iterrows():
            row_str = "| " + " | ".join([str(value) for value in row]) + " |\n"
            f.write(row_str)
    
    print(f"\nSuccessfully wrote content to {md_file}")
    
    # Output as json file
    json_file = "附件五_土建+杆塔内容.json"
    # Convert DataFrame to list of dictionaries
    data = df.to_dict('records')
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print(f"Successfully wrote content to {json_file}")
    
except Exception as e:
    print(f"Error reading file: {str(e)}")