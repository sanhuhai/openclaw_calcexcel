import pandas as pd

file_path = "外包结算单测试模版.xlsx"

# 读取Excel文件
xl = pd.ExcelFile(file_path)
print(f"工作表: {xl.sheet_names}")

# 读取第一个工作表
df = pd.read_excel(file_path, sheet_name='工程结算金额')
print(f"\n数据形状: {df.shape}")
print(f"\n前5行数据:")
print(df.head())
print(f"\n列名: {df.columns.tolist()}")

# 查找钉钉编号
print("\n\n查找钉钉编号:")
for col in df.columns:
    if '钉钉' in str(col):
        print(f"找到列: {col}")

# 查看所有数据
print("\n\n所有数据:")
print(df)