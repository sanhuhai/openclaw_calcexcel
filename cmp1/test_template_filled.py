import pandas as pd

try:
    # 测试模板文件
    df = pd.read_excel('外包结算单测试模版_template_20260317_000801.xlsx', sheet_name='工程结算金额')
    print('工程结算金额工作表读取成功，数据形状:', df.shape)
    print('列名:', df.columns.tolist())
    print('\n数据内容（前10行）:')
    print(df.head(10))
    
    # 检查施工内容和单项结算金额单价列
    if '施工内容' in df.columns and '单项结算金额单价' in df.columns:
        print('\n施工内容列数据（前10行）:')
        print(df['施工内容'].head(10).tolist())
        print('\n单项结算金额单价列数据（前10行）:')
        print(df['单项结算金额单价'].head(10).tolist())
    else:
        print('\n警告: 未找到施工内容或单项结算金额单价列')
    
    # 测试搬运结算金额工作表
    df2 = pd.read_excel('外包结算单测试模版_template_20260317_000801.xlsx', sheet_name='搬运结算金额')
    print('\n搬运结算金额工作表读取成功，数据形状:', df2.shape)
    print('数据内容:')
    print(df2)
    
except Exception as e:
    print('读取错误:', str(e))
    import traceback
    traceback.print_exc()