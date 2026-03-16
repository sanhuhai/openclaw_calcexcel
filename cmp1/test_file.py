import pandas as pd

try:
    df = pd.read_excel('结算比较结果_20260316_230023.xlsx')
    print('文件读取成功，数据形状:', df.shape)
    print('列名:', df.columns.tolist())
    print('前5行数据:')
    print(df.head())
except Exception as e:
    print('读取错误:', str(e))