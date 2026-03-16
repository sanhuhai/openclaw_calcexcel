import pandas as pd

try:
    # 读取生成的模板文件
    df = pd.read_excel('外包结算单测试模版_template_20260317_003917.xlsx', sheet_name='工程结算金额')
    print('工程结算金额工作表读取成功，数据形状:', df.shape)
    print('\n列名:', df.columns.tolist())
    
    # 检查是否没有额外的列
    if '红河文山单价' not in df.columns and '普洱单价' not in df.columns:
        print('\n✅ 没有添加额外的"红河文山单价"和"普洱单价"列')
    else:
        print('\n❌ 仍然存在额外的列')
        if '红河文山单价' in df.columns:
            print('  - 红河文山单价列存在')
        if '普洱单价' in df.columns:
            print('  - 普洱单价列存在')
    
    # 显示前5行数据
    print('\n前5行数据（施工内容、单项结算金额单价、施工量）:')
    display_cols = ['施工内容', '单项结算金额单价', '施工量']
    available_cols = [col for col in display_cols if col in df.columns]
    print(df[available_cols].head(5))
    
    # 显示有施工量的行
    if '施工量' in df.columns:
        non_null_quantity = df['施工量'].notna().sum()
        print(f'\n非空施工量数量: {non_null_quantity}')
        if non_null_quantity > 0:
            print('\n有施工量的行:')
            print(df[df['施工量'].notna()][available_cols])
    
except Exception as e:
    print('错误:', str(e))
    import traceback
    traceback.print_exc()
