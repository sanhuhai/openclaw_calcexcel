# 示例代码
import pandas as pd
import os
from datetime import datetime

DEBUG = True

def filter_excel(file_path, sheet_name, output_dir=None):
    # 读取Excel文件
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # 识别"模块名称"和"备注"列
    columns = df.columns.tolist()
    module_name_index = columns.index("模块名称") if "模块名称" in columns else -1
    remark_index = columns.index("备注") if "备注" in columns else -1
    
    # 确定开始筛选的列索引
    start_index = max(module_name_index, remark_index) + 1 if max(module_name_index, remark_index) >= 0 else 0
    
    # 只保留"模块名称"以及"备注"以后的列（不包含"备注"列）
    if module_name_index >= 0:
        # 包含模块名称列，加上备注之后的列
        if remark_index >= 0:
            columns_to_keep = [columns[module_name_index]] + columns[remark_index + 1:]
        else:
            # 如果没有备注列，只保留模块名称列
            columns_to_keep = [columns[module_name_index]]
        df = df[columns_to_keep]
    elif remark_index >= 0:
        # 如果没有模块名称列，只保留备注之后的列
        df = df[columns[remark_index + 1:]]
    
    # 生成输出文件名
    if output_dir is None:
        output_dir = os.path.dirname(file_path)
    base_name = os.path.basename(file_path)
    name_without_ext = os.path.splitext(base_name)[0]
    if DEBUG:
        output_file = os.path.join(output_dir, datetime.now().strftime("%Y%m%d%H%M%S")+ f"{name_without_ext}selectdata.xlsx")
    else:
        output_file = os.path.join(output_dir, f"{name_without_ext}selectdata.xlsx")
    
    # 保存筛选结果
    df.to_excel(output_file, index=False)
    
    return output_file

# 使用示例
file_path = "E:\\test\\openclaw\\附件五：中国铁塔云南公司2024-2025年土建杆塔施工集中采购项目施工模块安全生产费明细表.xlsx"
sheet_name = "土建+杆塔"
# 假设"模块名称"和"备注"之后的列有相关筛选字段
output_file = filter_excel(file_path, sheet_name)
print(f"筛选结果已保存到: {output_file}")