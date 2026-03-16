import pandas as pd

file_path = "外包结算单测试模版.xlsx"

# Read 工程结算金额 sheet
print("=== 原始文件所有数据 ===")
df = pd.read_excel(file_path, sheet_name="工程结算金额")

# 查看所有行的数据
print(f"总行数: {len(df)}")
print("\n所有行的数据:")
for i, row in df.iterrows():
    print(f"第{i}行: {row.tolist()[:5]}")  # 只显示前5列