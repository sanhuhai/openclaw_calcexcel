import pandas as pd

try:
    # 读取生成的模板文件
    df = pd.read_excel('外包结算单测试模版_template_20260317_010539.xlsx', sheet_name='工程结算金额')
    print('工程结算金额工作表读取成功，数据形状:', df.shape)
    print('\n列名:', df.columns.tolist())
    
    # 显示数据总行数
    print(f'\n总行数: {len(df)}')
    
    # 显示前15行数据（红河/文山区域的所有县名）
    print('\n前15行数据（红河/文山区域，所有县名）:')
    display_cols = ['区域', '施工内容', '单项结算金额单价', '施工量']
    available_cols = [col for col in display_cols if col in df.columns]
    print(df[available_cols].head(15))
    
    # 显示第95-108行数据（交界区域，普洱区域的所有县名）
    print('\n第95-108行数据（红河/文山和普洱交界，普洱区域所有县名）:')
    print(df[available_cols].iloc[94:108])
    
    # 显示最后10行数据（普洱区域）
    print('\n最后10行数据（普洱区域）:')
    print(df[available_cols].tail(10))
    
    # 统计区域分布
    if '区域' in df.columns:
        print('\n区域分布统计:')
        print(df['区域'].value_counts().sort_index())
    
    # 统计有施工量的行
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
