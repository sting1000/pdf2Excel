@echo off
echo PDF表格转Excel工具启动脚本

:: 检查Python环境
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未找到Python, 请先安装Python
    pause
    exit /b
)

:: 检查虚拟环境
if not exist venv\ (
    echo 创建Python虚拟环境...
    python -m venv venv
)

:: 激活虚拟环境
echo 激活虚拟环境...
call venv\Scripts\activate.bat

:: 安装依赖
echo 安装依赖...
pip install -r requirements.txt

:: 运行程序
echo 启动程序...
python pdf_table_converter.py

:: 退出虚拟环境
deactivate

pause 