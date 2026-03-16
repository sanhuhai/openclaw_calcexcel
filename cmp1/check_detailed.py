import pandas as pd

file_path = "外包结算单测试模版.xlsx"

# Read 工程结算金额 sheet
print("=== 工程结算金额详细数据 ===")
df = pd.read_excel(file_path, sheet_name="工程结算金额")

# 查看第6-17行的详细数据
print("\n第6-17行的详细数据:")
for i in range(6, 17):
    print(f"\n第{i}行:")
    row_data = df.iloc[i]
    for j, value in enumerate(row_data):
        if pd.notna(value) and str(value).strip() != 'nan':
            print(f"  列{j} ({df.columns[j]}): {value}")
