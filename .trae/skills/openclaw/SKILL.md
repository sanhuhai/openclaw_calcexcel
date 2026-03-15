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
   - 处理calc表格，按照列名称不为空进行筛选
   - 在新的表中新增指定的列名
   - 存储到表格名称+final新的表中
   - 保存final表格文件
   - 处理calc表格，按照列名称不为空进行筛选
   - 在新的表中新增指定的列名
   - 存储到表格名称+final新的表中
   - 保存final表格文件
   - 处理calc表格，按照列名称不为空进行筛选
   - 在新的表中新增指定的列名
   - 存储到表格名称+final新的表中
   - 保存final表格文件

3. **示例**：

```python
# 示例代码
import pandas as pd
import os
import datetime

DEBUG = False
import datetime

DEBUG = False
import datetime

DEBUG = False

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
    
    # 生成calc文件路径
    if DEBUG:
        calc_file = os.path.join(output_dir, datetime.now().strftime("%Y%m%d%H%M%S") + f"{name_without_ext}calc.xlsx")
    else:
        calc_file = os.path.join(output_dir, f"{name_without_ext}calc.xlsx")
    
    # 删除calc数据表和selectdata数据表
    try:
        if os.path.exists(output_file):
            os.remove(output_file)
            print(f"已删除selectdata文件: {output_file}")
        if os.path.exists(calc_file):
            os.remove(calc_file)
            print(f"已删除calc文件: {calc_file}")
    except Exception as e:
        print(f"删除文件时出错: {e}")
    
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
    
    # 处理calc表格，生成final表格
    process_calc_data(calc_file, output_dir, base_name)

def process_calc_data(calc_file, output_dir, base_name):
    # 读取calc文件
    df = pd.read_excel(calc_file)
    
    # 检查是否有数据
    if df.empty:
        return
    
    # 按照列名称不为空进行筛选
    # 只保留列名不为空且不包含"Unnamed"的列
    columns_with_name = [col for col in df.columns if pd.notna(col) and col.strip() != "" and "Unnamed" not in str(col)]
    df = df[columns_with_name]
    
    # 移除指定的列
    columns_to_remove = [
        "施工量",
        "单项结算金额单价",
        "单项结算金额",
        "工程结算金额31%",
        "外包结算比例",
        "外包结算单价",
        "外包结算金额",
        "外包结算金额31%",
        "本次外包结算比例",
        "本次外包结算金额",
        "外包已结算金额",
        "外包结算剩余金额",
        "利润率比例",
        "外包请款日期",
        "备注"
    ]
    
    # 移除这些列
    df = df.drop(columns=columns_to_remove, errors='ignore')
    
    # 重新整理数据格式，确保包含"城市名称"、"模块名称"和"不含增值税基准价（不含安全生产费）"三列
    # 检查是否有足够的列
    if len(df.columns) >= 1:
        # 存储所有处理后的数据
        all_data = []
        
        # 遍历所有列（从第二列开始）
        for i in range(1, len(df.columns)):
            # 获取城市名称（当前列的列名）
            city_name = df.columns[i]
            
            # 创建临时DataFrame
            temp_df = pd.DataFrame()
            
            # 填充模块名称列，使用原数据的第一列
            temp_df["模块名称"] = df.iloc[:, 0]
            
            # 填充城市名称列，使用当前列的列名，应用到所有行
            temp_df["城市名称"] = city_name
            
            # 填充不含增值税基准价（不含安全生产费）列，使用当前列对应的值
            temp_df["不含增值税基准价（不含安全生产费）"] = df.iloc[:, i]
            
            # 添加到所有数据中
            all_data.append(temp_df)
        
        # 合并所有数据
        if all_data:
            df = pd.concat(all_data, ignore_index=True)
        else:
            # 如果没有数据，创建空的DataFrame
            df = pd.DataFrame({
                "城市名称": [],
                "模块名称": [],
                "不含增值税基准价（不含安全生产费）": []
            })
    else:
        # 如果没有列，创建空的DataFrame
        df = pd.DataFrame({
            "城市名称": [],
            "模块名称": [],
            "不含增值税基准价（不含安全生产费）": []
        })
    
    # 重新整理数据格式，确保包含"城市名称"、"模块名称"和"不含增值税基准价（不含安全生产费）"三列
    # 检查是否有足够的列
    if len(df.columns) >= 3:
        # 重命名列
        df.columns = ["城市名称", "模块名称", "不含增值税基准价（不含安全生产费）"]
    elif len(df.columns) == 2:
        # 如果只有两列，添加第三列
        df.columns = ["城市名称", "模块名称"]
        df["不含增值税基准价（不含安全生产费）"] = None
    elif len(df.columns) == 1:
        # 如果只有一列，添加两列
        df.columns = ["城市名称"]
        df["模块名称"] = None
        df["不含增值税基准价（不含安全生产费）"] = None
    
    # 新增指定的列名
    new_columns = [
        "施工量",
        "单项结算金额单价",
        "单项结算金额",
        "工程结算金额31%",
        "外包结算比例",
        "外包结算单价",
        "外包结算金额",
        "外包结算金额31%",
        "本次外包结算比例",
        "本次外包结算金额",
        "外包已结算金额",
        "外包结算剩余金额",
        "利润率比例",
        "外包请款日期",
        "备注"
    ]
    
    # 为每个新列添加空值
    for column in new_columns:
        if column not in df.columns:
            df[column] = None
    
    # 将不含增值税基准价（不含安全生产费）列的数据复制到单项结算金额单价列
    if "不含增值税基准价（不含安全生产费）" in df.columns and "单项结算金额单价" in df.columns:
        df["单项结算金额单价"] = df["不含增值税基准价（不含安全生产费）"]
    
    # 按照规则计算各列的值
    # 施工量默认为0（如果未设置）
    if "施工量" in df.columns:
        df["施工量"] = df["施工量"].fillna(0)
    else:
        df["施工量"] = 0

    
    # 计算单项结算金额 = 施工量 * 单项结算金额单价
    if "单项结算金额" in df.columns and "施工量" in df.columns and "单项结算金额单价" in df.columns:
        df["单项结算金额"] = df["施工量"] * df["单项结算金额单价"]
    
    # 计算工程结算金额31% = 单项结算金额 * 0.31
    if "工程结算金额31%" in df.columns and "单项结算金额" in df.columns:
        df["工程结算金额31%"] = df["单项结算金额"] * 0.31
    
    # 外包结算比例 = 70%
    if "外包结算比例" in df.columns:
        df["外包结算比例"] = 0.7
    
    # 计算外包结算单价 = 单项结算单价 * 外包结算比例
    if "外包结算单价" in df.columns and "单项结算金额单价" in df.columns and "外包结算比例" in df.columns:
        df["外包结算单价"] = df["单项结算金额单价"] * df["外包结算比例"]
    
    # 计算外包结算金额 = 施工量 * 外包结算单价
    if "外包结算金额" in df.columns and "施工量" in df.columns and "外包结算单价" in df.columns:
        df["外包结算金额"] = df["施工量"] * df["外包结算单价"]
    
    # 计算外包结算金额31% = 外包结算金额 * 0.31
    if "外包结算金额31%" in df.columns and "外包结算金额" in df.columns:
        df["外包结算金额31%"] = df["外包结算金额"] * 0.31
    
    # 本次外包结算比例 = 90.00%
    if "本次外包结算比例" in df.columns:
        df["本次外包结算比例"] = 0.9
    
    # 计算本次外包结算金额 = 外包结算金额31% * 本次外包结算比例
    if "本次外包结算金额" in df.columns and "外包结算金额31%" in df.columns and "本次外包结算比例" in df.columns:
        df["本次外包结算金额"] = df["外包结算金额31%"] * df["本次外包结算比例"]
    
    # 外包已结算金额 = 0.00
    if "外包已结算金额" in df.columns:
        df["外包已结算金额"] = 0.0
    
    # 计算外包结算剩余金额 = 外包结算金额31% - 本次外包结算金额 - 外包已结算金额
    if "外包结算剩余金额" in df.columns and "外包结算金额31%" in df.columns and "本次外包结算金额" in df.columns and "外包已结算金额" in df.columns:
        df["外包结算剩余金额"] = df["外包结算金额31%"] - df["本次外包结算金额"] - df["外包已结算金额"]
    
    # 计算利润率比例 = (工程结算金额31% - 外包结算金额31%) / 工程结算金额31%
    if "利润率比例" in df.columns and "工程结算金额31%" in df.columns and "外包结算金额31%" in df.columns:
        # 避免除以零
        mask = df["工程结算金额31%"] != 0
        df.loc[mask, "利润率比例"] = (df.loc[mask, "工程结算金额31%"] - df.loc[mask, "外包结算金额31%"]) / df.loc[mask, "工程结算金额31%"]
        df.loc[~mask, "利润率比例"] = 0
    
    # 生成final文件名
    if DEBUG:
        final_file = os.path.join(output_dir, datetime.now().strftime("%Y%m%d%H%M%S") + f"{base_name}final.xlsx")
    else:
        final_file = os.path.join(output_dir, f"{base_name}final.xlsx")
    
    # 保存到final表格
    df.to_excel(final_file, index=False)
    
    print(f"最终数据已保存到: {final_file}")

# 使用示例
import glob

def process_all_excel_files(folder_path, sheet_name):
    """处理文件夹下所有的.xlsx文件"""
    # 获取文件夹下所有的.xlsx文件
    excel_files = glob.glob(os.path.join(folder_path, "*.xlsx"))
    
    if not excel_files:
        print(f"文件夹 {folder_path} 中没有找到.xlsx文件")
        return
    
    print(f"找到 {len(excel_files)} 个.xlsx文件，开始处理...")
    
    for file_path in excel_files:
        print(f"处理文件: {file_path}")
        try:
            output_file = filter_excel(file_path, sheet_name)
            print(f"文件处理完成，结果保存到: {output_file}")
        except Exception as e:
            print(f"处理文件 {file_path} 时出错: {e}")
    
    print("所有文件处理完成！")

# 处理文件夹下所有的.xlsx文件
if __name__ == "__main__":
    folder_path = "E:\test\openclaw"  # 替换为实际的计算文件夹路径
    sheet_name = "土建+杆塔"  # 替换为实际的工作表名称
    process_all_excel_files(folder_path, sheet_name)
```

## 注意事项

- 确保Excel文件存在且可访问
- 筛选条件中的列名必须与Excel表格中的列名完全匹配，且只能是"模块名称"和"备注"之后的列
- 输出结果只包含"模块名称"以及"备注"以后的列（不包含"备注"列）
- 筛选完成后，会自动生成一个表格名+calc的表格，包含第一列的名称和所有有内容的列数据
- 最后会生成一个表格名称+final的表格，按照列名称不为空且不包含"Unnamed"进行筛选，移除指定的列名，并按照"城市名称"、"模块名称"和"不含增值税基准价（不含安全生产费）"的格式重新整理数据
- 对于final表格的第二列以后的所有列，都会按照相同的规则处理，并将结果合并到一个表格中
- final表格会新增指定的列：施工量、单项结算金额单价、单项结算金额、工程结算金额31%、外包结算比例、外包结算单价、外包结算金额、外包结算金额31%、本次外包结算比例、本次外包结算金额、外包已结算金额、外包结算剩余金额、利润率比例、外包请款日期、备注
- final表格会按照以下规则计算各列的值：
  - 施工量默认为0
  - 单项结算金额 = 施工量 * 单项结算金额单价
  - 工程结算金额31% = 单项结算金额 * 0.31
  - 外包结算比例 = 70%
  - 外包结算单价 = 单项结算单价 * 外包结算比例
  - 外包结算金额 = 施工量 * 外包结算单价
  - 外包结算金额31% = 外包结算金额 * 0.31
  - 本次外包结算比例 = 90.00%
  - 本次外包结算金额 = 外包结算金额31% * 本次外包结算比例
  - 外包已结算金额 = 0.00
  - 外包结算剩余金额 = 外包结算金额31% - 本次外包结算金额 - 外包已结算金额
  - 利润率比例 = (工程结算金额31% - 外包结算金额31%) / 工程结算金额31%
- final表格会计算以下列的和，并在表格末尾添加合计行：
  - 单项结算金额
  - 工程结算金额31%
  - 外包结算金额
  - 外包结算金额31%
  - 本次外包结算金额
  - 外包已结算金额
  - 外包结算剩余金额
- 处理完成后，会自动删除calc数据表和selectdata数据表，只保留final数据表
- final表格会计算以下列的和，并在表格末尾添加合计行：
  - 单项结算金额
  - 工程结算金额31%
  - 外包结算金额
  - 外包结算金额31%
  - 本次外包结算金额
  - 外包已结算金额
  - 外包结算剩余金额
- 处理完成后，会自动删除calc数据表和selectdata数据表，只保留final数据表
- final表格会计算以下列的和，并在表格末尾添加合计行：
  - 单项结算金额
  - 工程结算金额31%
  - 外包结算金额
  - 外包结算金额31%
  - 本次外包结算金额
  - 外包已结算金额
  - 外包结算剩余金额
- 处理完成后，会自动删除calc数据表和selectdata数据表，只保留final数据表
- final表格会计算以下列的和，并在表格末尾添加合计行：
  - 单项结算金额
  - 工程结算金额31%
  - 外包结算金额
  - 外包结算金额31%
  - 本次外包结算金额
  - 外包已结算金额
  - 外包结算剩余金额
- 处理完成后，会自动删除calc数据表和selectdata数据表，只保留final数据表
- final表格会按照以下规则计算各列的值：
  - 施工量默认为0
  - 单项结算金额 = 施工量 * 单项结算金额单价
  - 工程结算金额31% = 单项结算金额 * 0.31
  - 外包结算比例 = 70%
  - 外包结算单价 = 单项结算单价 * 外包结算比例
  - 外包结算金额 = 施工量 * 外包结算单价
  - 外包结算金额31% = 外包结算金额 * 0.31
  - 本次外包结算比例 = 90.00%
  - 本次外包结算金额 = 外包结算金额31% * 本次外包结算比例
  - 外包已结算金额 = 0.00
  - 外包结算剩余金额 = 外包结算金额31% - 本次外包结算金额 - 外包已结算金额
  - 利润率比例 = (工程结算金额31% - 外包结算金额31%) / 工程结算金额31%
- final表格会按照以下规则计算各列的值：
  - 施工量默认为0
  - 单项结算金额 = 施工量 * 单项结算金额单价
  - 工程结算金额31% = 单项结算金额 * 0.31
  - 外包结算比例 = 70%
  - 外包结算单价 = 单项结算单价 * 外包结算比例
  - 外包结算金额 = 施工量 * 外包结算单价
  - 外包结算金额31% = 外包结算金额 * 0.31
  - 本次外包结算比例 = 90.00%
  - 本次外包结算金额 = 外包结算金额31% * 本次外包结算比例
  - 外包已结算金额 = 0.00
  - 外包结算剩余金额 = 外包结算金额31% - 本次外包结算金额 - 外包已结算金额
  - 利润率比例 = (工程结算金额31% - 外包结算金额31%) / 工程结算金额31%
- 最后会生成一个表格名称+final的表格，按照列名称不为空进行筛选，并新增指定的列名
- 最后会生成一个表格名称+final的表格，按照列名称不为空进行筛选，并新增指定的列名
- 支持的筛选条件为精确匹配，如需其他筛选方式（如包含、范围等），需修改代码
- 大文件可能会导致处理时间较长，请耐心等待

## 依赖项

- pandas
- openpyxl（用于读写Excel文件）

## 安装依赖

```bash
pip install pandas openpyxl
```