---
name: "openclaw-settlement-reader"
description: "读取外包结算单测试模版.xlsx文件的工程结算金额表格，保存单元格内容为钉钉编号的单元格及后面的数据。同时读取附件五表格数据，生成模板文件并自动计算相关金额列。当用户需要处理外包结算单数据时调用。"
---

# Openclaw 外包结算单读取技能

## 功能
- 读取 `外包结算单测试模版.xlsx` 文件中的工程结算金额和搬运结算金额表格
- 识别并提取单元格内容为 `钉钉编号` 的单元格及其后面的数据
- 删除提取数据中内容为 `合计` 的单元格及后面的数据
- 删除生成表格的第一行（包含Unnamed的行，不是真正的列名）
- 读取附件五文件中的模块名称、红河/文山单价和普洱单价数据
- 比较附件五数据和提取数据，根据比较结果添加颜色标记
- 生成模板文件（只保留标题行）
- 将附件五的红河/文山和普洱数据依次填入模板
- 在区域列填入具体的县名（region_mapping的key）
- 通过{区域、施工内容、单项结算金额单价}三列匹配来填充施工量
- 根据10个计算公式自动计算所有相关金额列
- 删除模板表格中的{比较结果，附件五单价}列
- 保存提取的数据和生成的模板到新的Excel文件中

## 计算公式
1. 单项结算金额 = 施工量 × 单项结算金额单价
2. 工程结算金额31% = 单项结算金额 × 0.31
3. 外包结算比例 = 70%（固定值）
4. 外包结算单价 = 单项结算单价 × 外包结算比例
5. 外包结算金额 = 施工量 × 外包结算单价
6. 外包结算金额31% = 外包结算金额 × 0.31
7. 本次外包结算比例 = 90.00%（固定值）
8. 本次外包结算金额 = 外包结算金额31% × 本次外包结算比例
9. 外包结算剩余金额 = 外包结算金额31% - 本次外包结算金额 - 外包已结算金额
10. 利润率比例 = (工程结算金额31% - 外包结算金额31%) / 工程结算金额31%

## 使用场景
- 当用户需要从外包结算单中提取工程结算金额数据时
- 当用户需要处理包含钉钉编号的结算表格数据时
- 当用户需要自动化处理外包结算单数据时
- 当用户需要对比附件五数据和提取数据时
- 当用户需要生成包含计算金额的模板文件时

## 实现原理
1. 打开 `外包结算单测试模版.xlsx` 文件
2. 遍历所有工作表
3. 在每个工作表中查找内容为 `钉钉编号` 的单元格
4. 提取该单元格及其后续的所有数据
5. 读取附件五文件中的模块名称和单价数据
6. 比较附件五数据和提取数据，根据比较结果添加颜色标记
7. 生成模板文件（只保留标题行）
8. 将附件五的红河/文山和普洱数据依次填入模板
9. 在区域列填入具体的县名
10. 通过三列匹配（区域、施工内容、单价）来填充施工量
11. 根据10个计算公式自动计算所有相关金额列
12. 删除不需要的列（比较结果、附件五单价）
13. 将提取的数据和生成的模板保存到新的Excel文件中

## 注意事项
- 确保 `外包结算单测试模版.xlsx` 文件存在于当前工作目录
- 确保附件五文件存在于当前工作目录
- 确保文件格式正确，包含工程结算金额和搬运结算金额表格
- 确保表格中包含 `钉钉编号` 标识的单元格

## 代码实现

```python
#!/usr/bin/env python3
"""
Read 外包结算单测试模版.xlsx file, find the cell with content "钉钉编号",
and save that cell and the data after it.
Also compare with 附件五 file and add color markings.
"""

import pandas as pd
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill

# Define the file paths
file_path = "外包结算单测试模版.xlsx"
attachment_file = "附件五：中国铁塔云南公司2024-2025年土建杆塔施工集中采购项目施工模块安全生产费明细表.xlsx"

def read_attachment_five():
    """读取附件五Excel文件，提取模块名称和红河/文山、普洱列数据"""
    print(f"\n读取附件五文件: {attachment_file}")
    
    if not os.path.exists(attachment_file):
        print(f"警告: 文件 {attachment_file} 不存在，跳过比较功能")
        return None
    
    # 读取Excel文件
    xl = pd.ExcelFile(attachment_file)
    print(f"工作表: {xl.sheet_names}")
    
    # 读取"土建+杆塔"工作表
    if '土建+杆塔' in xl.sheet_names:
        df = pd.read_excel(attachment_file, sheet_name='土建+杆塔')
        print(f"读取工作表: 土建+杆塔")
    else:
        df = pd.read_excel(attachment_file, sheet_name=xl.sheet_names[0])
        print(f"读取工作表: {xl.sheet_names[0]}")
    
    print(f"数据形状: {df.shape}")
    print(f"列名: {df.columns.tolist()}")
    
    # 查找模块名称、红河/文山和普洱列
    module_col = None
    price_col_honghe = None
    price_col_puer = None
    
    for col in df.columns:
        col_str = str(col).strip()
        if '模块名称' in col_str:
            module_col = col
            print(f"找到模块名称列: {col}")
        if '红河' in col_str or '文山' in col_str:
            price_col_honghe = col
            print(f"找到红河/文山列: {col}")
        if '普洱' in col_str:
            price_col_puer = col
            print(f"找到普洱列: {col}")
    
    if module_col is None or price_col_honghe is None or price_col_puer is None:
        print("警告: 未找到模块名称、红河/文山或普洱列，跳过比较功能")
        return None
    
    # 提取数据
    data = df[[module_col, price_col_honghe, price_col_puer]].copy()
    data.columns = ['施工内容', '红河文山单价', '普洱单价']
    
    # 删除空值行和标题行
    data = data.dropna(subset=['施工内容'])
    
    # 删除包含"模块名称"的行（标题行）
    data = data[data['施工内容'] != '模块名称']
    
    print(f"\n附件五提取的数据（前10行）:")
    print(data.head(10))
    print(f"\n总共 {len(data)} 条数据")
    
    return data

def compare_data(extracted_data, attachment_data):
    """比较提取的数据和附件五数据，返回比较结果"""
    if attachment_data is None:
        print("\n跳过数据比较（附件五数据不可用）")
        return extracted_data
    
    print("\n开始比较数据...")
    
    # 只比较"工程结算金额"工作表
    if '工程结算金额' not in extracted_data:
        print("未找到工程结算金额工作表，跳过比较")
        return extracted_data
    
    # 获取工程结算金额数据
    comparison_data = extracted_data['工程结算金额'].copy()
    
    # 添加比较结果列
    comparison_data['比较结果'] = ''
    comparison_data['附件五单价'] = None
    
    # 将附件五数据转换为字典，方便查找
    attachment_dict_honghe = {}
    attachment_dict_puer = {}
    for _, row in attachment_data.iterrows():
        content = str(row['施工内容']).strip()
        price_honghe = row['红河文山单价']
        price_puer = row['普洱单价']
        attachment_dict_honghe[content] = price_honghe
        attachment_dict_puer[content] = price_puer
    
    print(f"附件五红河/文山数据字典（前5条）: {dict(list(attachment_dict_honghe.items())[:5])}")
    print(f"附件五普洱数据字典（前5条）: {dict(list(attachment_dict_puer.items())[:5])}")
    
    # 比较每一行
    for idx, row in comparison_data.iterrows():
        content = str(row['施工内容']).strip()
        price = row['单项结算金额单价']
        region = str(row['区域']).strip() if '区域' in row and pd.notna(row['区域']) else ''
        
        # 根据区域选择对应的字典
        # 定义区域映射
        region_mapping = {
            '红河': '红河文山',
            '文山': '红河文山',
            '河口': '红河文山',
            '屏边': '红河文山',
            '金平': '红河文山',
            '绿春': '红河文山',
            '元阳': '红河文山',
            '红河': '红河文山',
            '文山': '红河文山',
            '西畴': '红河文山',
            '麻栗坡': '红河文山',
            '马关': '红河文山',
            '丘北': '红河文山',
            '广南': '红河文山',
            '富宁': '红河文山',
            '普洱': '普洱',
            '思茅': '普洱',
            '宁洱': '普洱',
            '墨江': '普洱',
            '景东': '普洱',
            '景谷': '普洱',
            '镇沅': '普洱',
            '江城': '普洱',
            '孟连': '普洱',
            '澜沧': '普洱',
            '西盟': '普洱'
        }
        
        # 查找区域类型
        region_type = None
        for key, value in region_mapping.items():
            if key in region:
                region_type = value
                break
        
        if region_type == '红河文山':
            attachment_dict = attachment_dict_honghe
            print(f"使用红河/文山价格进行比较: {content}")
        elif region_type == '普洱':
            attachment_dict = attachment_dict_puer
            print(f"使用普洱价格进行比较: {content}")
        else:
            attachment_dict = attachment_dict_honghe
            print(f"区域'{region}'未识别，使用红河/文山价格进行比较: {content}")
        
        if content in attachment_dict:
            attachment_price = attachment_dict[content]
            comparison_data.at[idx, '附件五单价'] = attachment_price
            
            # 比较价格是否相同
            try:
                # 转换为浮点数进行比较
                gen_price = float(price) if pd.notna(price) else 0
                att_price = float(attachment_price) if pd.notna(attachment_price) else 0
                
                # 允许一定的误差范围（0.01）
                if abs(gen_price - att_price) < 0.01:
                    comparison_data.at[idx, '比较结果'] = '相同'
                else:
                    comparison_data.at[idx, '比较结果'] = '不一致'
            except (ValueError, TypeError):
                # 如果无法转换为数字，直接比较字符串
                if str(price) == str(attachment_price):
                    comparison_data.at[idx, '比较结果'] = '相同'
                else:
                    comparison_data.at[idx, '比较结果'] = '不一致'
        else:
            comparison_data.at[idx, '比较结果'] = '不存在'
            comparison_data.at[idx, '附件五单价'] = None
    
    print(f"\n比较结果统计:")
    print(comparison_data['比较结果'].value_counts())
    
    # 更新提取的数据
    extracted_data['工程结算金额'] = comparison_data
    
    return extracted_data

def save_template_only(all_extracted_data, output_file, attachment_data):
    """只保存标题行到Excel文件，并填充附件五数据"""
    print(f"\n保存模板文件（只包含标题行）: {output_file}")
    
    # 创建新的工作簿
    wb = Workbook()
    
    # 删除默认的工作表
    if 'Sheet' in wb.sheet_names:
        wb.remove(wb['Sheet'])
    
    # 处理每个工作表的数据
    for sheet_name, extracted_data in all_extracted_data.items():
        # 创建工作表
        ws = wb.create_sheet(title=sheet_name)
        
        # 只写入标题行
        headers = list(extracted_data.columns)
        
        for col_idx, header in enumerate(headers, 1):
            ws.cell(row=1, column=col_idx, value=header)
        
        # 如果有附件五数据，填充到模板中
        if attachment_data is not None and sheet_name == '工程结算金额':
            print(f"\n填充附件五数据到模板的{sheet_name}工作表...")
            
            # 查找施工内容、单项结算金额单价、施工量和区域列的索引
            content_col_idx = None
            price_col_idx = None
            quantity_col_idx = None
            region_col_idx = None
            for col_idx, header in enumerate(headers, 1):
                if header == '施工内容':
                    content_col_idx = col_idx
                elif header == '单项结算金额单价':
                    price_col_idx = col_idx
                elif header == '施工量':
                    quantity_col_idx = col_idx
                elif header == '区域':
                    region_col_idx = col_idx
            
            if content_col_idx is not None and price_col_idx is not None:
                # 定义红河/文山县名列表和普洱县名列表
                honghe_wenshan_counties = ['红河', '文山', '河口', '屏边', '金平', '绿春', '元阳', '西畴', '麻栗坡', '马关', '丘北', '广南', '富宁']
                puer_counties = ['普洱', '思茅', '宁洱', '墨江', '景东', '景谷', '镇沅', '江城', '孟连', '澜沧', '西盟']
                
                # 填充附件五数据（先填充红河/文山价格）
                print(f"\n第一步：填充红河/文山数据...")
                for row_idx, (_, row) in enumerate(attachment_data.iterrows(), 2):
                    content = row['施工内容']
                    price_honghe = row['红河文山单价']
                    
                    # 计算应该使用哪个县名（循环使用县名）
                    county_index = (row_idx - 2) % len(honghe_wenshan_counties)
                    county_name = honghe_wenshan_counties[county_index]
                    
                    # 写入施工内容
                    ws.cell(row=row_idx, column=content_col_idx, value=content)
                    # 写入单项结算金额单价（使用红河/文山价格）
                    if pd.notna(price_honghe):
                        if isinstance(price_honghe, (int, float)):
                            ws.cell(row=row_idx, column=price_col_idx, value=price_honghe)
                        else:
                            ws.cell(row=row_idx, column=price_col_idx, value=str(price_honghe))
                    else:
                        ws.cell(row=row_idx, column=price_col_idx, value="")
                    # 写入区域（使用具体的县名）
                    if region_col_idx is not None:
                        ws.cell(row=row_idx, column=region_col_idx, value=county_name)
                
                # 填充附件五数据（再填充普洱价格，放在红河/文山数据下面）
                print(f"\n第二步：填充普洱数据（放在红河/文山数据下面）...")
                start_row = len(attachment_data) + 2  # 从红河/文山数据之后开始
                for row_idx, (_, row) in enumerate(attachment_data.iterrows(), start_row):
                    content = row['施工内容']
                    price_puer = row['普洱单价']
                    
                    # 计算应该使用哪个县名（循环使用县名）
                    county_index = (row_idx - start_row) % len(puer_counties)
                    county_name = puer_counties[county_index]
                    
                    # 写入施工内容
                    ws.cell(row=row_idx, column=content_col_idx, value=content)
                    # 写入单项结算金额单价（使用普洱价格）
                    if pd.notna(price_puer):
                        if isinstance(price_puer, (int, float)):
                            ws.cell(row=row_idx, column=price_col_idx, value=price_puer)
                        else:
                            ws.cell(row=row_idx, column=price_col_idx, value=str(price_puer))
                    else:
                        ws.cell(row=row_idx, column=price_col_idx, value="")
                    # 写入区域（使用具体的县名）
                    if region_col_idx is not None:
                        ws.cell(row=row_idx, column=region_col_idx, value=county_name)
                
                print(f"成功填充 {len(attachment_data)} 条红河/文山数据到模板")
                print(f"成功填充 {len(attachment_data)} 条普洱数据到模板（放在红河/文山数据下面）")
                print(f"总共 {len(attachment_data) * 2} 条数据")
                
                # 填充extracted表格中的施工量数据
                if quantity_col_idx is not None:
                    print(f"\n填充extracted表格中的施工量数据到模板...")
                    
                    # 获取extracted表格中的数据
                    extracted_data = all_extracted_data[sheet_name]
                    
                    # 定义区域映射，用于将extracted中的具体县名映射到区域类型
                    region_mapping = {
                        '红河': '红河文山',
                        '文山': '红河文山',
                        '河口': '红河文山',
                        '屏边': '红河文山',
                        '金平': '红河文山',
                        '绿春': '红河文山',
                        '元阳': '红河文山',
                        '西畴': '红河文山',
                        '麻栗坡': '红河文山',
                        '马关': '红河文山',
                        '丘北': '红河文山',
                        '广南': '红河文山',
                        '富宁': '红河文山',
                        '普洱': '普洱',
                        '思茅': '普洱',
                        '宁洱': '普洱',
                        '墨江': '普洱',
                        '景东': '普洱',
                        '景谷': '普洱',
                        '镇沅': '普洱',
                        '江城': '普洱',
                        '孟连': '普洱',
                        '澜沧': '普洱',
                        '西盟': '普洱'
                    }
                    
                    # 遍历template表格中的每一行（从第2行开始，第1行是标题）
                    # 现在总共有 len(attachment_data) * 2 行数据
                    total_rows = len(attachment_data) * 2
                    for template_row_idx in range(2, total_rows + 2):
                        # 获取template表格中当前行的区域、施工内容和价格
                        template_region = ws.cell(row=template_row_idx, column=region_col_idx).value if region_col_idx is not None else None
                        template_content = ws.cell(row=template_row_idx, column=content_col_idx).value
                        template_price = ws.cell(row=template_row_idx, column=price_col_idx).value
                        
                        # 在extracted表格中查找匹配的行（根据区域、施工内容和价格）
                        for _, extracted_row in extracted_data.iterrows():
                            extracted_region = str(extracted_row['区域']).strip() if '区域' in extracted_row and pd.notna(extracted_row['区域']) else ''
                            extracted_content = extracted_row['施工内容']
                            extracted_price = extracted_row['单项结算金额单价']
                            extracted_quantity = extracted_row['施工量']
                            
                            # 比较区域是否匹配（template中的县名属于extracted中的区域类型）
                            region_match = False
                            if template_region and extracted_region:
                                # 查找template中的县名对应的区域类型
                                template_region_type = None
                                for key, value in region_mapping.items():
                                    if key in str(template_region):
                                        template_region_type = value
                                        break
                                # 查找extracted中的县名对应的区域类型
                                extracted_region_type = None
                                for key, value in region_mapping.items():
                                    if key in extracted_region:
                                        extracted_region_type = value
                                        break
                                # 比较区域类型是否相同
                                region_match = (template_region_type == extracted_region_type) if template_region_type and extracted_region_type else False
                            
                            # 比较施工内容是否相等
                            content_match = str(template_content) == str(extracted_content) if pd.notna(template_content) and pd.notna(extracted_content) else False
                            # 比较价格是否相等（允许0.01的误差）
                            price_match = False
                            if pd.notna(template_price) and pd.notna(extracted_price):
                                try:
                                    price_match = abs(float(template_price) - float(extracted_price)) < 0.01
                                except:
                                    price_match = str(template_price) == str(extracted_price)
                            
                            if region_match and content_match and price_match:
                                # 填充施工量
                                if pd.notna(extracted_quantity):
                                    if isinstance(extracted_quantity, (int, float)):
                                        ws.cell(row=template_row_idx, column=quantity_col_idx, value=extracted_quantity)
                                    else:
                                        ws.cell(row=template_row_idx, column=quantity_col_idx, value=str(extracted_quantity))
                                else:
                                    ws.cell(row=template_row_idx, column=quantity_col_idx, value="")
                                print(f"匹配成功: 区域={template_region}, 施工内容={template_content}, 价格={template_price}, 施工量={extracted_quantity}")
                                break
                    
                    # 计算所有需要的列
                    print(f"\n计算其他列数据...")
                    
                    # 查找所有需要计算的列的索引
                    total_amount_col_idx = None
                    project_amount_31_col_idx = None
                    outsourcing_ratio_col_idx = None
                    outsourcing_price_col_idx = None
                    outsourcing_amount_col_idx = None
                    outsourcing_amount_31_col_idx = None
                    current_outsourcing_ratio_col_idx = None
                    current_outsourcing_amount_col_idx = None
                    outsourcing_remaining_col_idx = None
                    profit_ratio_col_idx = None
                    outsourced_already_col_idx = None
                    
                    for col_idx, header in enumerate(headers, 1):
                        if header == '单项结算金额':
                            total_amount_col_idx = col_idx
                        elif header == '工程结算金额31%':
                            project_amount_31_col_idx = col_idx
                        elif header == '外包结算比例':
                            outsourcing_ratio_col_idx = col_idx
                        elif header == '外包结算单价':
                            outsourcing_price_col_idx = col_idx
                        elif header == '外包结算金额':
                            outsourcing_amount_col_idx = col_idx
                        elif header == '外包结算金额31%':
                            outsourcing_amount_31_col_idx = col_idx
                        elif header == '本次外包结算比例':
                            current_outsourcing_ratio_col_idx = col_idx
                        elif header == '本次外包结算金额':
                            current_outsourcing_amount_col_idx = col_idx
                        elif header == '外包结算剩余金额':
                            outsourcing_remaining_col_idx = col_idx
                        elif header == '利润率比例':
                            profit_ratio_col_idx = col_idx
                        elif header == '外包已结算金额':
                            outsourced_already_col_idx = col_idx
                    
                    # 遍历所有行进行计算
                    for template_row_idx in range(2, total_rows + 2):
                        # 获取需要的值
                        quantity = ws.cell(row=template_row_idx, column=quantity_col_idx).value if quantity_col_idx is not None else None
                        price = ws.cell(row=template_row_idx, column=price_col_idx).value if price_col_idx is not None else None
                        outsourced_already = ws.cell(row=template_row_idx, column=outsourced_already_col_idx).value if outsourced_already_col_idx is not None else 0
                        
                        # 只有当有施工量时才进行计算
                        if pd.notna(quantity) and pd.notna(price) and quantity != '' and price != '':
                            try:
                                quantity_num = float(quantity)
                                price_num = float(price)
                                outsourced_already_num = float(outsourced_already) if pd.notna(outsourced_already) else 0
                                
                                # 1. 单项结算金额 = 施工量 * 单项结算金额单价
                                total_amount = quantity_num * price_num
                                if total_amount_col_idx is not None:
                                    ws.cell(row=template_row_idx, column=total_amount_col_idx, value=total_amount)
                                
                                # 2. 工程结算金额31% = 单项结算金额 * 0.31
                                project_amount_31 = total_amount * 0.31
                                if project_amount_31_col_idx is not None:
                                    ws.cell(row=template_row_idx, column=project_amount_31_col_idx, value=project_amount_31)
                                
                                # 3. 外包结算比例 = 70%
                                outsourcing_ratio = 0.7
                                if outsourcing_ratio_col_idx is not None:
                                    ws.cell(row=template_row_idx, column=outsourcing_ratio_col_idx, value=outsourcing_ratio)
                                
                                # 4. 外包结算单价 = 单项结算单价 * 外包结算比例
                                outsourcing_price = price_num * outsourcing_ratio
                                if outsourcing_price_col_idx is not None:
                                    ws.cell(row=template_row_idx, column=outsourcing_price_col_idx, value=outsourcing_price)
                                
                                # 5. 外包结算金额 = 施工量 * 外包结算单价
                                outsourcing_amount = quantity_num * outsourcing_price
                                if outsourcing_amount_col_idx is not None:
                                    ws.cell(row=template_row_idx, column=outsourcing_amount_col_idx, value=outsourcing_amount)
                                
                                # 6. 外包结算金额31% = 外包结算金额 * 0.31
                                outsourcing_amount_31 = outsourcing_amount * 0.31
                                if outsourcing_amount_31_col_idx is not None:
                                    ws.cell(row=template_row_idx, column=outsourcing_amount_31_col_idx, value=outsourcing_amount_31)
                                
                                # 7. 本次外包结算比例 = 90.00%
                                current_outsourcing_ratio = 0.9
                                if current_outsourcing_ratio_col_idx is not None:
                                    ws.cell(row=template_row_idx, column=current_outsourcing_ratio_col_idx, value=current_outsourcing_ratio)
                                
                                # 8. 本次外包结算金额 = 外包结算金额31% * 本次外包结算比例
                                current_outsourcing_amount = outsourcing_amount_31 * current_outsourcing_ratio
                                if current_outsourcing_amount_col_idx is not None:
                                    ws.cell(row=template_row_idx, column=current_outsourcing_amount_col_idx, value=current_outsourcing_amount)
                                
                                # 9. 外包结算剩余金额 = 外包结算金额31% - 本次外包结算金额 - 外包已结算金额
                                outsourcing_remaining = outsourcing_amount_31 - current_outsourcing_amount - outsourced_already_num
                                if outsourcing_remaining_col_idx is not None:
                                    ws.cell(row=template_row_idx, column=outsourcing_remaining_col_idx, value=outsourcing_remaining)
                                
                                # 10. 利润率比例 = (工程结算金额31% - 外包结算金额31%) / 工程结算金额31%
                                profit_ratio = 0
                                if project_amount_31 != 0:
                                    profit_ratio = (project_amount_31 - outsourcing_amount_31) / project_amount_31
                                if profit_ratio_col_idx is not None:
                                    ws.cell(row=template_row_idx, column=profit_ratio_col_idx, value=profit_ratio)
                                
                            except Exception as e:
                                print(f"第 {template_row_idx} 行计算出错: {str(e)}")
                    print(f"计算完成！")
                    
                    # 删除不需要的列：比较结果、附件五单价
                    print(f"\n删除不需要的列...")
                    # 从后往前删除，避免索引变化
                    cols_to_delete = ['比较结果', '附件五单价']
                    for col_name in reversed(cols_to_delete):
                        for col_idx, header in enumerate(headers, 1):
                            if header == col_name:
                                ws.delete_cols(col_idx)
                                print(f"已删除列: {col_name}")
                                break
        
        # 调整列宽
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if cell.value and len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width
    
    # 保存文件
    try:
        wb.save(output_file)
        print(f"成功保存模板文件: {output_file}")
        # 验证文件是否可以打开
        from openpyxl import load_workbook
        test_wb = load_workbook(output_file)
        print(f"模板文件验证成功，包含 {len(test_wb.sheetnames)} 个工作表: {test_wb.sheetnames}")
        return True
    except Exception as e:
        print(f"保存模板文件时出错: {str(e)}")
        return False

def save_to_excel_with_openpyxl(all_extracted_data, output_file):
    """使用openpyxl保存数据到Excel文件，并添加颜色标记"""
    print(f"\n保存结果到文件: {output_file}")
    
    # 创建新的工作簿
    wb = Workbook()
    
    # 删除默认的工作表
    if 'Sheet' in wb.sheet_names:
        wb.remove(wb['Sheet'])
    
    # 定义颜色
    green_fill = PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')  # 浅绿色
    red_fill = PatternFill(start_color='FFB6C1', end_color='FFB6C1', fill_type='solid')    # 浅红色
    yellow_fill = PatternFill(start_color='FFFFE0', end_color='FFFFE0', fill_type='solid')  # 浅黄色
    
    # 处理每个工作表的数据
    for sheet_name, extracted_data in all_extracted_data.items():
        # 创建工作表
        ws = wb.create_sheet(title=sheet_name)
        
        # 写入标题行
        headers = list(extracted_data.columns)
        for col_idx, header in enumerate(headers, 1):
            ws.cell(row=1, column=col_idx, value=header)
        
        # 写入数据并添加颜色
        for row_idx, (_, row) in enumerate(extracted_data.iterrows(), 2):
            for col_idx, (col_name, value) in enumerate(row.items(), 1):
                # 写入数据
                if pd.notna(value):
                    if isinstance(value, (int, float)):
                        ws.cell(row=row_idx, column=col_idx, value=value)
                    else:
                        ws.cell(row=row_idx, column=col_idx, value=str(value))
                else:
                    ws.cell(row=row_idx, column=col_idx, value="")
                
                # 添加颜色标记（只对工程结算金额工作表的单项结算金额单价列）
                if sheet_name == '工程结算金额' and col_name == '单项结算金额单价':
                    if '比较结果' in row:
                        result = row['比较结果']
                        if result == '相同':
                            ws.cell(row=row_idx, column=col_idx).fill = green_fill
                        elif result == '不一致':
                            ws.cell(row=row_idx, column=col_idx).fill = red_fill
                        elif result == '不存在':
                            ws.cell(row=row_idx, column=col_idx).fill = yellow_fill
        
        # 调整列宽
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if cell.value and len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width
    
    # 保存文件
    try:
        wb.save(output_file)
        print(f"成功保存文件: {output_file}")
        # 验证文件是否可以打开
        from openpyxl import load_workbook
        test_wb = load_workbook(output_file)
        print(f"文件验证成功，包含 {len(test_wb.sheetnames)} 个工作表: {test_wb.sheetnames}")
        return True
    except Exception as e:
        print(f"保存文件时出错: {str(e)}")
        return False

def extract_data_from_sheet(df, sheet_name):
    """从工作表中提取数据"""
    print(f"\nProcessing sheet: {sheet_name}")
    print(f"Sheet shape: {df.shape}")
    
    # 检查是否已经有"钉钉编号"列（已经是提取好的格式）
    if '钉钉编号' in df.columns:
        print(f"'钉钉编号' column found directly in sheet {sheet_name}")
        print(f"原始数据行数: {len(df)}")
        
        # 只保留有钉钉编号的行
        extracted_data = df[pd.notna(df['钉钉编号']) & (df['钉钉编号'].astype(str).str.strip() != '')].copy()
        extracted_data = extracted_data.reset_index(drop=True)
        
        print(f"过滤后数据行数（只保留有钉钉编号的行）: {len(extracted_data)}")
        print(f"删除的行数（没有钉钉编号）: {len(df) - len(extracted_data)}")
        print(f"Extracted data shape: {extracted_data.shape}")
        print("Extracted data:")
        print(extracted_data)
        return extracted_data
    
    # 搜索包含"钉钉编号"的单元格
    found = False
    for row_idx, row in df.iterrows():
        for col_idx, value in enumerate(row):
            if str(value) == "钉钉编号":
                print(f"Found '{value}' at row {row_idx}, column {col_idx}")
                
                # Extract data from this cell onwards
                # Get the remaining columns from this cell
                remaining_columns = df.columns[col_idx:]
                # Get the remaining rows including current row
                extracted_data = df.iloc[row_idx:, col_idx:]
                
                # Remove rows from "合计" onwards
                total_row_idx = None
                for i, row in enumerate(extracted_data.iterrows()):
                    _, row_data = row
                    for value in row_data:
                        value_str = str(value).strip()
                        # Check if the value contains "合计" (with or without spaces)
                        if "合" in value_str and "计" in value_str:
                            total_row_idx = i
                            break
                    if total_row_idx is not None:
                        break
                
                if total_row_idx is not None:
                    print(f"Found '合计' at row {total_row_idx} in extracted data")
                    extracted_data = extracted_data.iloc[:total_row_idx]
                    print(f"Data shape after removing '合计' and beyond: {extracted_data.shape}")
                
                # Set the first row as column names and remove it from data
                if not extracted_data.empty:
                    # Get the first row as column names
                    new_columns = extracted_data.iloc[0]
                    # Remove the first row from data
                    extracted_data = extracted_data.iloc[1:]
                    # Set the new column names
                    extracted_data.columns = new_columns
                    print("Set first row as column names and removed it from data")
                
                # Process merged cells: fill empty cells with previous values
                if not extracted_data.empty:
                    # Create a copy to avoid modifying the original
                    processed_data = extracted_data.copy()
                    # Initialize variables to store previous values
                    previous_values = {}
                    
                    # Iterate through each row
                    for i, row in processed_data.iterrows():
                        # Check if this row has a钉钉编号
                        has_dingding = pd.notna(row.get('钉钉编号')) and str(row.get('钉钉编号')).strip() != ''
                        
                        if has_dingding:
                            # Save all non-empty values for future rows
                            for col in processed_data.columns:
                                if pd.notna(row[col]) and str(row[col]).strip() != '':
                                    previous_values[col] = row[col]
                        else:
                            # Fill empty cells with previous values
                            for col in processed_data.columns:
                                if not (pd.notna(row[col]) and str(row[col]).strip() != ''):
                                    if col in previous_values:
                                        processed_data.at[i, col] = previous_values[col]
                    
                    # 只保留有钉钉编号的行
                    processed_data = processed_data[pd.notna(processed_data['钉钉编号']) & (processed_data['钉钉编号'].astype(str).str.strip() != '')].copy()
                    processed_data = processed_data.reset_index(drop=True)
                    
                    # Replace with processed data
                    extracted_data = processed_data
                    print("Processed merged cells and kept only rows with 钉钉编号")
                
                # Reset index to ensure all rows are included
                extracted_data = extracted_data.reset_index(drop=True)
                
                print(f"Extracted data shape: {extracted_data.shape}")
                print("Extracted data:")
                print(extracted_data)
                
                return extracted_data
    
    print(f"'钉钉编号' not found in sheet {sheet_name}")
    return None

def main():
    """主函数"""
    print("=" * 60)
    print("开始处理结算数据")
    print("=" * 60)
    print(f"File: {file_path}")
    
    if not os.path.exists(file_path):
        print(f"错误: 文件 {file_path} 不存在")
        return
    
    try:
        # 读取附件五数据
        attachment_data = read_attachment_five()
        
        # Read all sheets from the Excel file
        xl = pd.ExcelFile(file_path)
        print(f"\n工作表: {xl.sheet_names}")
        
        # 存储所有提取的数据
        all_extracted_data = {}
        
        # Process each sheet
        for sheet_name in xl.sheet_names:
            # Read the sheet
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            # 提取数据
            extracted_data = extract_data_from_sheet(df, sheet_name)
            
            if extracted_data is not None:
                # 存储提取的数据
                all_extracted_data[sheet_name] = extracted_data
                print(f"\nSuccessfully extracted data from sheet '{sheet_name}'")
        
        # 比较数据
        all_extracted_data = compare_data(all_extracted_data, attachment_data)
        
        # 保存所有提取的数据到Excel文件
        if all_extracted_data:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = f"外包结算单测试模版_extracted_{timestamp}.xlsx"
            if save_to_excel_with_openpyxl(all_extracted_data, output_file):
                print(f"\nAll sheets processed. Final output file: {output_file}")
            else:
                print("\n文件保存失败")
            
            # 保存模板文件（只包含标题行）
            template_file = f"外包结算单测试模版_template_{timestamp}.xlsx"
            if save_template_only(all_extracted_data, template_file, attachment_data):
                print(f"\nTemplate file created: {template_file}")
            else:
                print("\n模板文件保存失败")
        else:
            print("\n没有提取到任何数据")
                
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        import traceback
        traceback.print_exc()
    
    print("=" * 60)
    print("处理完成!")
    print("=" * 60)

if __name__ == "__main__":
    main()
```