import pandas as pd

try:
    # 读取生成的模板文件
    df = pd.read_excel('外包结算单测试模版_template_20260317_003353.xlsx', sheet_name='工程结算金额')
    print('工程结算金额工作表读取成功，数据形状:', df.shape)
    print('\n列名:', df.columns.tolist())
    
    # 检查是否有普洱单价列
    if '普洱单价' in df.columns:
        print('\n普洱单价列存在！')
        print('普洱单价数据（前10行）:')
        print(df['普洱单价'].head(10))
    else:
        print('\n普洱单价列不存在')
        print('只有这些列:', df.columns.tolist())
    
    # 检查是否有红河文山单价列
    if '红河文山单价' in df.columns:
        print('\n红河文山单价列存在！')
        print('红河文山单价数据（前10行）:')
        print(df['红河文山单价'].head(10))
    else:
        print('\n红河文山单价列不存在')
    
except Exception as e:
    print('错误:', str(e))
    import traceback
    traceback.print_exc()
