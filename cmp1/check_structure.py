import pandas as pd

file_path = "外包结算单测试模版.xlsx"

# Read 工程结算金额 sheet
print("=== 工程结算金额 ===")
df = pd.read_excel(file_path, sheet_name="工程结算金额")
print(f"Shape: {df.shape}")
print("\nFirst 15 rows:")
print(df.head(15))

print("\n\n=== 搬运结算金额 ===")
df2 = pd.read_excel(file_path, sheet_name="搬运结算金额")
print(f"Shape: {df2.shape}")
print("\nAll rows:")
print(df2)