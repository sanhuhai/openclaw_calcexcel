#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OpenClaw 结算计算器
计算.xlsx文件中的各项结算金额
"""

import os
import pandas as pd

# 计算规则配置
OUTSOURCING_RATIO = 0.7  # 外包结算比例
ENGINEERING_RATIO = 0.31  # 工程结算比例
CURRENT_OUTSOURCING_RATIO = 0.9  # 本次外包结算比例


def calculate_settlement(file_path):
    """
    计算单个Excel文件的结算金额
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)
        
        # 计算各项指标
        df['单项结算金额'] = df['施工量'] * df['单项结算金额单价']
        df['工程结算金额31%'] = df['单项结算金额'] * ENGINEERING_RATIO
        df['外包结算比例'] = OUTSOURCING_RATIO
        df['外包结算单价'] = df['单项结算金额单价'] * OUTSOURCING_RATIO
        df['外包结算金额'] = df['施工量'] * df['外包结算单价']
        df['外包结算金额31%'] = df['外包结算金额'] * ENGINEERING_RATIO
        df['本次外包结算比例'] = CURRENT_OUTSOURCING_RATIO
        df['本次外包结算金额'] = df['外包结算金额'] * 0.3 * CURRENT_OUTSOURCING_RATIO
        df['外包已结算金额'] = 0.00
        df['外包结算剩余金额'] = df['外包结算金额31%'] - df['本次外包结算金额'] - df['外包已结算金额']
        
        # 计算利润率比例，避免除零错误
        def calculate_profit_margin(row):
            if row['工程结算金额31%'] == 0:
                return 0
            return (row['工程结算金额31%'] - row['外包结算金额31%']) / row['工程结算金额31%']
        
        df['利润率比例'] = df.apply(calculate_profit_margin, axis=1)
        
        # 计算指定列的总和
        sum_columns = ['单项结算金额', '工程结算金额31%', '外包结算金额', '外包结算金额31%', '本次外包结算金额', '外包已结算金额', '外包结算剩余金额']
        
        # 创建总和行
        sum_row = pd.Series(index=df.columns, name='总计')
        
        # 计算各列的总和
        for col in sum_columns:
            if col in df.columns:
                sum_row[col] = df[col].sum()
        
        # 将总和行添加到DataFrame末尾
        # df = df.append(sum_row, ignore_index=True)
        sum_row_df = pd.DataFrame([sum_row])
        df = pd.concat([df, sum_row_df], ignore_index=True)
        
        # 保存计算结果
        output_file = file_path.replace('.xlsx', '_calculated.xlsx')
        df.to_excel(output_file, index=False)
        
        return f"成功处理文件: {file_path}\n结果保存到: {output_file}"
        
    except Exception as e:
        return f"处理文件 {file_path} 时出错: {str(e)}"


def process_all_xlsx():
    """
    处理当前目录下的所有.xlsx文件
    """
    current_dir = os.getcwd()
    xlsx_files = [f for f in os.listdir(current_dir) if f.endswith('final.xlsx')]
    
    if not xlsx_files:
        return "当前目录下没有找到.xlsx文件"
    
    results = []
    for file in xlsx_files:
        file_path = os.path.join(current_dir, file)
        result = calculate_settlement(file_path)
        results.append(result)
    
    return "\n".join(results)


if __name__ == "__main__":
    print("OpenClaw 结算计算器开始运行...")
    print("=" * 50)
    result = process_all_xlsx()
    print(result)
    print("=" * 50)
    print("计算完成！")