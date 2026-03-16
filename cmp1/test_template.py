import pandas as pd

try:
    # 测试模板文件
    df = pd.read_excel('外包结算单测试模版_template_20260317_000352.xlsx', sheet_name='工程结算金额')
    print('工程结算金额工作表读取成功，数据形状:', df.shape)
    print('列名:', df.columns.tolist())
    print('\n数据内容:')
    print(df)
    
    # 测试搬运结算金额工作表
    df2 = pd.read_excel('外包结算单测试模版_template_20260317_000352.xlsx', sheet_name='搬运结算金额')
    print('\n搬运结算金额工作表读取成功，数据形状:', df2.shape)
    print('列名:', df2.columns.tolist())
    print('\n数据内容:')
    print(df2)
    
except Exception as e:
    print('读取错误:', str(e))
    import traceback
    traceback.print_exc()