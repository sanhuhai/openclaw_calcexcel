import pandas as pd

try:
    # 读取生成的模板文件
    df = pd.read_excel('外包结算单测试模版_template_20260317_003501.xlsx', sheet_name='工程结算金额')
    print('工程结算金额工作表读取成功，数据形状:', df.shape)
    print('\n列名:', df.columns.tolist())
    
    # 检查是否有普洱单价列
    if '普洱单价' in df.columns:
        print('\n✅ 普洱单价列存在！')
        print('普洱单价数据（前10行）:')
        print(df['普洱单价'].head(10).tolist())
    else:
        print('\n❌ 普洱单价列不存在')
    
    # 检查是否有红河文山单价列
    if '红河文山单价' in df.columns:
        print('\n✅ 红河文山单价列存在！')
        print('红河文山单价数据（前10行）:')
        print(df['红河文山单价'].head(10).tolist())
    else:
        print('\n❌ 红河文山单价列不存在')
    
    # 显示完整的几行数据
    print('\n前5行完整数据:')
    display_cols = ['施工内容', '单项结算金额单价', '红河文山单价', '普洱单价', '施工量']
    available_cols = [col for col in display_cols if col in df.columns]
    print(df[available_cols].head(5))
    
except Exception as e:
    print('错误:', str(e))
    import traceback
    traceback.print_exc()
