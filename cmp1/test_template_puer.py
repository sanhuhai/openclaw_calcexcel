import pandas as pd

try:
    # 测试模板文件
    df = pd.read_excel('外包结算单测试模版_template_20260317_002238.xlsx', sheet_name='工程结算金额')
    print('工程结算金额工作表读取成功，数据形状:', df.shape)
    
    # 检查施工内容和单项结算金额单价列
    if '施工内容' in df.columns and '单项结算金额单价' in df.columns:
        print('\n施工内容和单项结算金额单价列数据（前10行）:')
        for i in range(min(10, len(df))):
            content = df['施工内容'].iloc[i]
            price = df['单项结算金额单价'].iloc[i]
            print(f"{i+1}. {content}: {price}")
        
        # 检查施工量列
        if '施工量' in df.columns:
            print('\n施工量列数据（前20行）:')
            print(df['施工量'].head(20).tolist())
            
            # 统计非空施工量的数量
            non_null_quantity = df['施工量'].notna().sum()
            print(f'\n非空施工量数量: {non_null_quantity}')
            
            # 显示有施工量的行
            if non_null_quantity > 0:
                print('\n有施工量的行:')
                quantity_rows = df[df['施工量'].notna()][['施工内容', '单项结算金额单价', '施工量']]
                print(quantity_rows)
            else:
                print('\n没有匹配到施工量数据（这是正常的，因为extracted表格中的价格是红河/文山的，而模板现在填充的是普洱的价格）')
    else:
        print('\n警告: 未找到施工内容或单项结算金额单价列')
    
except Exception as e:
    print('读取错误:', str(e))
    import traceback
    traceback.print_exc()