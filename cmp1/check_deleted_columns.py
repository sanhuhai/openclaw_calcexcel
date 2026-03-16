import pandas as pd

try:
    # 读取生成的模板文件
    df = pd.read_excel('外包结算单测试模版_template_20260317_013750.xlsx', sheet_name='工程结算金额')
    print('工程结算金额工作表读取成功，数据形状:', df.shape)
    
    # 显示列名
    print(f'\n所有列名: {df.columns.tolist()}')
    
    # 检查是否还有比较结果和附件五单价列
    print(f'\n比较结果列是否存在: {"比较结果" in df.columns}')
    print(f'附件五单价列是否存在: {"附件五单价" in df.columns}')
    
except Exception as e:
    print('错误:', str(e))
    import traceback
    traceback.print_exc()
