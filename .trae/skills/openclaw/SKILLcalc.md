---
name: "openclaw"
description: "进行Excel表格数据筛选，并将筛选结果复制到另一张表格，保存名字为原来表格名字+selectdata。当用户需要对Excel数据进行筛选并保存结果时调用。"
---

# OpenClaw Excel筛选工具

## 功能描述

该技能用于对Excel表格数据进行筛选，筛选字段为"模块名称"和"备注"之后的列内容，筛选条件后，只保存"模块名称"以及"备注"以后的列（不包含"备注"列）到新的表格中，新表格的名称为原表格名称加上"selectdata"后缀。然后打开筛选出来的表格，读取第一列的名称，从第二行开始读取有内容的列，将这些数据组成列表，存储到表格名+calc的表格数据中。

## 使用方法

1. **输入参数**：
   - `file_path`：Excel文件路径
   - `sheet_name`：要筛选的工作表名称
   - `filter_criteria`：筛选条件，格式为字典，键为列名，值为筛选值
   - `output_dir`：输出目录（可选，默认为原文件所在目录）

2. **执行流程**：
   - 打开指定的Excel文件
   - 选择指定的工作表
   - 识别"模块名称"和"备注"列
   - 只保留"模块名称"以及"备注"以后的列（不包含"备注"列）
   - 应用筛选条件到保留的列
   - 创建新的工作表，名称为原表格名称+selectdata
   - 将筛选结果复制到新工作表
   - 保存文件
   - 打开筛选出来的表格
   - 建立一个map
   - map的第一个数据key和值都是第一行读取第一列的名称
   - 从第二列第一行第二列开始，读取第一行的列名作为key，这一列的第一个数据作为value，插入到map中
   - 从第四行开始，读取这个map key所对应行的数据
   - 存储到表格名+calc的表格数据中
   - 将得到的表格的列名，替换成map对应的value
   - 保存calc表格文件

3. **示例**：

```python
# 示例代码
import pandas as pd
import os

def filter_excel(file_path, sheet_name, filter_criteria, output_dir=None):
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
    
    # 应用筛选条件，只对保留的列进行筛选
    if isinstance(filter_criteria, dict) and filter_criteria:
        # 确保筛选的列存在于当前的df中
        for column, value in filter_criteria.items():
            if column in df.columns:
                # 只在列存在时进行筛选
                df = df[df[column] == value]
    
    # 生成输出文件名
    if output_dir is None:
        output_dir = os.path.dirname(file_path)
    base_name = os.path.basename(file_path)
    name_without_ext = os.path.splitext(base_name)[0]
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
    calc_file = os.path.join(output_dir, f"{base_name}calc.xlsx")
    
    # 保存到calc表格
    data_from_fourth_row.to_excel(calc_file, index=False)
    
    print(f"处理后的数据已保存到: {calc_file}")

# 使用示例
file_path = "data.xlsx"
sheet_name = "Sheet1"
# 假设"模块名称"和"备注"之后的列有"状态"和"优先级"
filter_criteria = {"状态": "已完成", "优先级": "高"}
output_file = filter_excel(file_path, sheet_name, filter_criteria)
print(f"筛选结果已保存到: {output_file}")
```

## 注意事项

- 确保Excel文件存在且可访问
- 筛选条件中的列名必须与Excel表格中的列名完全匹配，且只能是"模块名称"和"备注"之后的列
- 输出结果只包含"模块名称"以及"备注"以后的列（不包含"备注"列）
- 筛选完成后，会自动生成一个表格名+calc的表格，包含第一列的名称和所有有内容的列数据
- 支持的筛选条件为精确匹配，如需其他筛选方式（如包含、范围等），需修改代码
- 大文件可能会导致处理时间较长，请耐心等待

## 依赖项

- pandas
- openpyxl（用于读写Excel文件）

## 安装依赖

```bash
pip install pandas openpyxl
```