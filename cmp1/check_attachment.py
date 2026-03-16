import pandas as pd

file_path = "附件五：中国铁塔云南公司2024-2025年土建杆塔施工集中采购项目施工模块安全生产费明细表.xlsx"

# 读取Excel文件
xl = pd.ExcelFile(file_path)
print(f"工作表: {xl.sheet_names}")

# 读取第一个工作表
df = pd.read_excel(file_path, sheet_name=xl.sheet_names[0])
print(f"\n数据形状: {df.shape}")
print(f"\n列名:")
for i, col in enumerate(df.columns):
    print(f"{i}: {col}")

print(f"\n前10行数据:")
print(df.head(10))