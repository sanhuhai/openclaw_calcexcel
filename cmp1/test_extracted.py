import pandas as pd

try:
    # 测试新生成的文件
    df = pd.read_excel('外包结算单测试模版_extracted_20260316_231316.xlsx', sheet_name='工程结算金额')
    print('工程结算金额工作表读取成功，数据形状:', df.shape)
    print('列名:', df.columns.tolist())
    print('前5行数据:')
    print(df.head())
    
    # 测试搬运结算金额工作表
    df2 = pd.read_excel('外包结算单测试模版_extracted_20260316_231316.xlsx', sheet_name='搬运结算金额')
    print('\n搬运结算金额工作表读取成功，数据形状:', df2.shape)
    print('前2行数据:')
    print(df2.head())
except Exception as e:
    print('读取错误:', str(e))