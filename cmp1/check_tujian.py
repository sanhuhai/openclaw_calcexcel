import pandas as pd

file_path = "附件五：中国铁塔云南公司2024-2025年土建杆塔施工集中采购项目施工模块安全生产费明细表.xlsx"

# 读取"土建+杆塔"工作表
df = pd.read_excel(file_path, sheet_name='土建+杆塔')
print(f"数据形状: {df.shape}")
print(f"\n列名:")
for i, col in enumerate(df.columns):
    print(f"{i}: {repr(col)}")

print(f"\n前20行数据:")
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
print(df.head(20))