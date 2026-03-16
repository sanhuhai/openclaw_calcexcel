import pandas as pd

try:
    # 测试模板文件
    df = pd.read_excel('外包结算单测试模版_template_20260317_001618.xlsx', sheet_name='工程结算金额')
    print('工程结算金额工作表读取成功，数据形状:', df.shape)
    
    # 检查施工量列
    if '施工量' in df.columns:
        print('\n施工量列数据（前20行）:')
        print(df['施工量'].head(20).tolist())
        
        # 统计非空施工量的数量
        non_null_quantity = df['施工量'].notna().sum()
        print(f'\n非空施工量数量: {non_null_quantity}')
        
        # 显示有施工量的行
        print('\n有施工量的行:')
        quantity_rows = df[df['施工量'].notna()][['施工内容', '单项结算金额单价', '施工量']]
        print(quantity_rows)
    else:
        print('\n警告: 未找到施工量列')
    
except Exception as e:
    print('读取错误:', str(e))
    import traceback
    traceback.print_exc()