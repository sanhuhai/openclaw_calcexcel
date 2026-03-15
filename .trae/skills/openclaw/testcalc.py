# 示例代码
import pandas as pd
import os
from datetime import datetime

DEBUG = True

def filter_excel(file_path, sheet_name,output_dir=None):
    # 读取Excel文件
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # 识别"模块名称"和"备注"列
    columns = df.columns.tolist()
    module_name_index = columns.index("模块名称") if "模块名称" in columns else -1
    remark_index = columns.index("备注") if "备注" in columns else -1
    
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
    
    # 处理筛选结果，生成calc表格
    process_filtered_data(output_file, output_dir, name_without_ext)
    
    return output_file

def process_filtered_data(input_file, output_dir, base_name):
    # 读取筛选后的文件
    df = pd.read_excel(input_file)
    
    # 检查是否有数据
    if df.empty:
        return
    
    # 检查数据行数是否足够
    if len(df) < 4:
        return
    
    # 建立map
    column_map = {}
    
    # 第一个数据key和值都是第一行读取第一列的名称
    first_column_name = df.columns[0] if len(df.columns) > 0 else ""
    column_map[first_column_name] = first_column_name
    
    # 从第二列第一行第二列开始，读取第一行的列名作为key，这一列的第一个数据作为value，插入到map中
    for i, column in enumerate(df.columns[1:], start=1):
        # 读取这一列的第一个数据（第二行）
        if len(df) > 0:
            first_value = df.iloc[0, i]
            column_map[column] = first_value
    
    # 从第四行开始，读取这个map key所对应行的数据
    # 第四行对应索引为3（因为Python从0开始）
    data_from_fourth_row = df.iloc[3:].copy()
    
    # 只保留map中的key对应的列
    columns_to_keep = list(column_map.keys())
    data_from_fourth_row = data_from_fourth_row[columns_to_keep]
    
    # 将得到的表格的列名，替换成map对应的value
    data_from_fourth_row.columns = [column_map[col] for col in data_from_fourth_row.columns]
    
    # 生成calc文件名
    if DEBUG:
        calc_file = os.path.join(output_dir, datetime.now().strftime("%Y%m%d%H%M%S") + f"{base_name}calc.xlsx")
    else:
        calc_file = os.path.join(output_dir, f"{base_name}calc.xlsx")
    
    # 保存到calc表格
    data_from_fourth_row.to_excel(calc_file, index=False)
    
    print(f"处理后的数据已保存到: {calc_file}")

# 使用示例
file_path = "E:\\test\\openclaw\\附件五：中国铁塔云南公司2024-2025年土建杆塔施工集中采购项目施工模块安全生产费明细表.xlsx"
sheet_name = "土建+杆塔"
# 假设"模块名称"和"备注"之后的列有相关筛选字段
output_file = filter_excel(file_path, sheet_name)
print(f"筛选结果已保存到: {output_file}")