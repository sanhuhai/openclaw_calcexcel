import pandas as pd

try:
    # 读取生成的模板文件
    df = pd.read_excel('外包结算单测试模版_template_20260317_012934.xlsx', sheet_name='工程结算金额')
    print('工程结算金额工作表读取成功，数据形状:', df.shape)
    
    # 显示列名
    print(f'\n所有列名: {df.columns.tolist()}')
    
    # 统计有施工量的行
    if '施工量' in df.columns:
        non_null_quantity = df['施工量'].notna().sum()
        print(f'\n非空施工量数量: {non_null_quantity}')
        if non_null_quantity > 0:
            print('\n有施工量的行（显示计算后的列）:')
            # 选择要显示的列
            display_cols = ['区域', '施工内容', '单项结算金额单价', '施工量', 
                          '单项结算金额', '工程结算金额31%', '外包结算比例', 
                          '外包结算单价', '外包结算金额', '外包结算金额31%',
                          '本次外包结算比例', '本次外包结算金额', '外包结算剩余金额',
                          '利润率比例']
            available_cols = [col for col in display_cols if col in df.columns]
            pd.set_option('display.float_format', lambda x: f'{x:,.4f}')
            print(df[df['施工量'].notna()][available_cols])
    
except Exception as e:
    print('错误:', str(e))
    import traceback
    traceback.print_exc()
