---
name: "settlement-reader"
description: "Reads 外包结算单测试模版.xlsx file, finds cells with 钉钉编号, extracts data, compares with 附件五 file, and saves results with color markings. Invoke when user needs to process settlement files."
---

# Settlement Reader

This skill processes settlement Excel files, extracts data, and compares with attachment files.

## Usage

### What it does
1. Reads `外包结算单测试模版.xlsx` file
2. Finds cells with content "钉钉编号" and extracts data
3. Reads `附件五` file for comparison
4. Compares data between the two files
5. Adds color markings to indicate comparison results
6. Saves extracted data to new Excel files
7. Creates template files with pre-filled data

### When to invoke
- When you need to process settlement Excel files
- When you need to extract data from settlement templates
- When you need to compare settlement data with attachment files
- When you need to generate colored comparison reports

### How it works
1. The skill searches for the input files in the current directory
2. It extracts data starting from cells with "钉钉编号"
3. It processes merged cells and removes summary rows
4. It compares extracted data with attachment data
5. It saves results with color-coded cells:
   - Green: Values match
   - Red: Values don't match
   - Yellow: Items don't exist in attachment
6. It generates both a complete report and a template file

### Example

To process settlement files in the current directory:

```bash
python read_settlement.py
```

To process files in a specific directory:

```bash
python read_settlement.py /path/to/directory
```

## Dependencies

- pandas
- openpyxl

## Output

The skill generates two files:
1. `外包结算单测试模版_extracted_<timestamp>.xlsx` - Complete report with color markings
2. `外包结算单测试模版_template_<timestamp>.xlsx` - Template file with pre-filled data