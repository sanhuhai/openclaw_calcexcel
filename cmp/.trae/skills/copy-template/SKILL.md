---
name: "copy-template"
description: "Copies the 外包结算单测试模版.xlsx file to create a backup and extracts content starting from the '钉钉编号' position. Invoke when user needs to duplicate the template file and extract specific content."
---

# Copy Template Skill

This skill copies the 外包结算单测试模版.xlsx file to create a backup version and extracts content starting from the "钉钉编号" position.

## Usage

1. Ensure the 外包结算单测试模版.xlsx file exists in the current directory
2. Invoke this skill to create a copy of the file and extract content from "钉钉编号" position

## Example

When you run this skill, it will:
1. Create a copy of 外包结算单测试模版.xlsx with a timestamp suffix (e.g., 外包结算单测试模版_backup_20260316_100317.xlsx)
2. Create a new file with content starting from the "钉钉编号" column (e.g., 外包结算单测试模版_extracted_20260316_100317.xlsx)

## Implementation

The skill uses Python to:
1. Create a full backup of the original file
2. Read the Excel file using pandas
3. Find the "钉钉编号" column
4. Extract content starting from that column
5. Save the extracted content to a new file

## Dependencies

- Python 3.x
- pandas library
- openpyxl library (for Excel file handling)