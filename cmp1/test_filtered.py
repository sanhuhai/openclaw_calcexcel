import pandas as pd

try:
    # 测试新生成的文件
    df = pd.read_excel('外包结算单测试模版_extracted_20260316_235611.xlsx', sheet_name='工程结算金额')
    print('工程结算金额工作表读取成功，数据形状:', df.shape)
    print('列名:', df.columns.tolist())
    
    # 显示所有钉钉编号
    print('\n所有钉钉编号:')
    print(df['钉钉编号'].tolist())
    
    # 检查是否有空值
    print('\n钉钉编号列中的空值数量:', df['钉钉编号'].isna().sum())
    
    # 显示前10行
    print('\n前10行数据:')
    print(df.head(10))
    
    # 测试搬运结算金额工作表
    df2 = pd.read_excel('外包结算单测试模版_extracted_20260316_235611.xlsx', sheet_name='搬运结算金额')
    print('\n搬运结算金额工作表读取成功，数据形状:', df2.shape)
    print('所有钉钉编号:', df2['钉钉编号'].tolist())
    
except Exception as e:
    print('读取错误:', str(e))
    import traceback
    traceback.print_exc()