import pandas as pd

file_path = "附件五：中国铁塔云南公司2024-2025年土建杆塔施工集中采购项目施工模块安全生产费明细表.xlsx"

# 读取"土建+杆塔"工作表
df = pd.read_excel(file_path, sheet_name='土建+杆塔')

print("所有列名:")
for i, col in enumerate(df.columns):
    print(f"{i}: {repr(col)}")

# 查找包含红河或文山的列
print("\n\n查找包含'红河'或'文山'的列:")
for i, col in enumerate(df.columns):
    col_str = str(col)
    if '红河' in col_str or '文山' in col_str:
        print(f"找到: {i}: {repr(col)}")

# 查看第3行（可能是标题行）
print("\n\n第3行数据（可能是标题行）:")
print(df.iloc[2].tolist())