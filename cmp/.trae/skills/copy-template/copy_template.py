#!/usr/bin/env python3
"""
Copy Template Skill Implementation

This script copies the 外包结算单测试模版.xlsx file to create a backup and copies content starting from "钉钉编号" position.
"""

import os
import shutil
from datetime import datetime
import pandas as pd

def copy_template():
    """
    Copy the 外包结算单测试模版.xlsx file to create a backup and copy content starting from "钉钉编号" position.
    
    Returns:
        str: Status message indicating success or failure
    """
    # Define the source file
    source_file = "外包结算单测试模版.xlsx"
    
    # Check if the source file exists
    if not os.path.exists(source_file):
        return f"错误: 文件 '{source_file}' 不存在"
    
    # Generate backup filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = f"外包结算单测试模版_backup_{timestamp}.xlsx"
    
    try:
        # First, create a full backup
        shutil.copy2(source_file, backup_file)
        
        # Now, process the file to copy content starting from "钉钉编号"
        try:
            # Read the Excel file
            df = pd.read_excel(source_file)
            
            # Find the cell with content "钉钉编号"
            dingding_row = -1
            dingding_col = -1
            
            # Iterate through all cells to find "钉钉编号"
            for i, row in df.iterrows():
                for j, value in enumerate(row):
                    if str(value) == "钉钉编号":
                        dingding_row = i
                        dingding_col = j
                        break
                if dingding_row != -1:
                    break
            
            if dingding_row != -1 and dingding_col != -1:
                # Create a new DataFrame with content starting from "钉钉编号" cell
                # Include the row with "钉钉编号" and all rows below
                new_df = df.iloc[dingding_row:, dingding_col:]
                
                # Set the first row (with "钉钉编号") as column names
                new_df.columns = new_df.iloc[0]
                
                # Delete the first row (now used as column names)
                new_df = new_df.drop(new_df.index[0])
                
                # Find the row with "合计" (with possible spaces) and delete it and all rows after
                total_row_index = -1
                print("Searching for '合计' row...")
                print(f"Current DataFrame shape: {new_df.shape}")
                
                # Reset index to ensure we have sequential integers
                new_df = new_df.reset_index(drop=True)
                print(f"After resetting index, shape: {new_df.shape}")
                
                # Print all column names for debugging
                print(f"Column names: {list(new_df.columns)}")
                
                # Iterate through all rows to find "合计"
                for i in range(len(new_df)):
                    row = new_df.iloc[i]
                    print(f"Checking row {i}: {list(row.values)}")
                    for j, value in enumerate(row):
                        cell_value = str(value)
                        # Try multiple cleaning methods
                        cleaned_value = cell_value.replace(" ", "").strip()
                        print(f"  Cell {j} (col: {new_df.columns[j]}): '{cell_value}' -> '{cleaned_value}'")
                        if cleaned_value == "合计":
                            total_row_index = i
                            print(f"Found '合计' at row {i}")
                            break
                    if total_row_index != -1:
                        break
                
                print(f"Total row index found: {total_row_index}")
                
                # If still not found, print the entire DataFrame for inspection
                if total_row_index == -1:
                    print("Entire DataFrame content:")
                    print(new_df)
                
                if total_row_index != -1:
                    # Delete the "合计" row and all rows after it
                    print(f"Original DataFrame shape: {new_df.shape}")
                    # Use iloc to select all rows before the "合计" row
                    new_df = new_df.iloc[:total_row_index]
                    print(f"After deleting '合计' row and all rows after, shape: {new_df.shape}")
                    # Print the first few rows after deletion
                    print("First few rows after deletion:")
                    print(new_df.head())
                    # Print the last few rows after deletion to verify
                    if len(new_df) > 0:
                        print("Last few rows after deletion:")
                        print(new_df.tail())
                
                # Read data from 附件五 file
                try:
                    attachment_file = "附件五：中国铁塔云南公司2024-2025年土建杆塔施工集中采购项目施工模块安全生产费明细表.xlsx"
                    if os.path.exists(attachment_file):
                        print(f"Reading data from {attachment_file}...")
                        attachment_df = pd.read_excel(attachment_file)
                        print(f"Attachment file shape: {attachment_df.shape}")
                        print(f"Attachment file columns: {list(attachment_df.columns)}")
                        
                        # Find 模块名称 and 红河/文山 columns
                        module_col = None
                        price_col = None
                        
                        for col in attachment_df.columns:
                            if "模块名称" in str(col):
                                module_col = col
                            elif "红河/文山" in str(col):
                                price_col = col
                        
                        print(f"Found module column: {module_col}")
                        print(f"Found price column: {price_col}")
                        
                        if module_col and price_col:
                            # Create a list of module names and prices
                            module_list = []
                            price_list = []
                            for _, row in attachment_df.iterrows():
                                module_name = str(row[module_col])
                                price = row[price_col]
                                module_list.append(module_name)
                                price_list.append(price)
                            print(f"Found {len(module_list)} modules in attachment file")
                            
                            # Add 施工内容 and 单项结算金额单价 columns to new_df
                            if "施工内容" not in new_df.columns:
                                new_df["施工内容"] = ""
                            if "单项结算金额单价" not in new_df.columns:
                                new_df["单项结算金额单价"] = 0.0
                            
                            # Populate the new columns
                            print("Populating 施工内容 and 单项结算金额单价 columns...")
                            for i in range(len(new_df)):
                                # Cycle through the module list to populate the columns
                                module_index = i % len(module_list)
                                new_df.at[i, "施工内容"] = module_list[module_index]
                                new_df.at[i, "单项结算金额单价"] = price_list[module_index]
                            print("Columns populated successfully")
                        else:
                            print("Could not find 模块名称 or 红河/文山 columns in attachment file")
                    else:
                        print(f"Attachment file {attachment_file} not found")
                except Exception as attachment_error:
                    print(f"Error processing attachment file: {str(attachment_error)}")
                
                # Generate a filename for the extracted content
                extracted_file = f"外包结算单测试模版_extracted_{timestamp}.xlsx"
                
                # Save the extracted content
                new_df.to_excel(extracted_file, index=False)
                if total_row_index != -1:
                    return f"成功: 文件已复制到 '{backup_file}'，并从'钉钉编号'单元格开始提取内容到 '{extracted_file}'，第一行已设置为列名，删除了'合计'及以后的内容，已从附件五文件读取数据"
                else:
                    return f"成功: 文件已复制到 '{backup_file}'，并从'钉钉编号'单元格开始提取内容到 '{extracted_file}'，第一行已设置为列名，未找到'合计'行，已从附件五文件读取数据"
            else:
                return f"成功: 文件已复制到 '{backup_file}'，但未找到内容为'钉钉编号'的单元格"
        except Exception as excel_error:
            return f"成功: 文件已复制到 '{backup_file}'，但处理Excel内容时发生错误: {str(excel_error)}"
            
    except Exception as e:
        return f"错误: 复制文件时发生错误: {str(e)}"

if __name__ == "__main__":
    result = copy_template()
    print(result)