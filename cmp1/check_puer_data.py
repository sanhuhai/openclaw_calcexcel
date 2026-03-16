import pandas as pd

try:
    # 读取附件五文件
    attachment_file = "附件五：中国铁塔云南公司2024-2025年土建杆塔施工集中采购项目施工模块安全生产费明细表.xlsx"
    xl = pd.ExcelFile(attachment_file)
    
    # 读取"土建+杆塔"工作表
    if '土建+杆塔' in xl.sheet_names:
        df = pd.read_excel(attachment_file, sheet_name='土建+杆塔')
        print(f"读取工作表: 土建+杆塔")
    else:
        df = pd.read_excel(attachment_file, sheet_name=xl.sheet_names[0])
        print(f"读取工作表: {xl.sheet_names[0]}")
    
    print(f"数据形状: {df.shape}")
    print(f"\n所有列名:")
    for i, col in enumerate(df.columns):
        print(f"{i+1}. '{col}'")
    
    # 查找包含"普洱"的列
    print(f"\n查找包含'普洱'的列:")
    puer_col = None
    for col in df.columns:
        if '普洱' in str(col):
            print(f"找到: '{col}'")
            puer_col = col
            break
    
    if puer_col:
        print(f"\n普洱列数据（前10行）:")
        print(df[puer_col].head(10))
        
        # 检查普洱列的数据类型和空值情况
        print(f"\n普洱列数据类型: {df[puer_col].dtype}")
        print(f"普洱列非空值数量: {df[puer_col].notna().sum()}")
        print(f"普洱列空值数量: {df[puer_col].isna().sum()}")
    else:
        print("未找到普洱列")
    
except Exception as e:
    print('错误:', str(e))
    import traceback
    traceback.print_exc()
