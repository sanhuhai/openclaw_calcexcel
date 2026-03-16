@echo off

REM 运行test.py的Windows脚本

echo 开始运行Excel筛选工具...

REM 检查Python是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: Python 未安装
    pause
    exit /b 1
)

REM 检查依赖是否安装
echo 检查依赖...
python -c "import pandas; import openpyxl" >nul 2>&1
if %errorlevel% neq 0 (
    echo 安装依赖...
    pip install pandas openpyxl
    if %errorlevel% neq 0 (
        echo 错误: 依赖安装失败
        pause
        exit /b 1
    )
)

REM 运行test.py
echo 运行test.py...
python test.py

echo 运行完成！
pause