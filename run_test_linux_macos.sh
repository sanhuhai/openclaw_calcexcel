#!/bin/bash

# 运行test.py的Linux/MacOS脚本

echo "开始运行Excel筛选工具..."

# 检查Python是否安装
if ! command -v python3 &> /dev/null; then
    echo "错误: Python 3 未安装"
    exit 1
fi

# 检查依赖是否安装
echo "检查依赖..."
python3 -c "import pandas; import openpyxl" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "安装依赖..."
    pip3 install pandas openpyxl
    if [ $? -ne 0 ]; then
        echo "错误: 依赖安装失败"
        exit 1
    fi
fi

# 运行test.py
echo "运行test.py..."
python3 test.py

echo "运行完成！"
