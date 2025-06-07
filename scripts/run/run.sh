#!/bin/bash

# 检查Python环境
if ! command -v python3 &> /dev/null; then
    echo "错误: 未找到Python3, 请先安装Python3"
    exit 1
fi

# 检查虚拟环境
if [ ! -d "venv" ]; then
    echo "创建Python虚拟环境..."
    python3 -m venv venv
fi

# 激活虚拟环境
echo "激活虚拟环境..."
source venv/bin/activate

# 安装依赖
echo "安装依赖..."
pip install -r requirements.txt

# 运行程序
echo "启动程序..."
python pdf_table_converter.py

# 退出虚拟环境
deactivate