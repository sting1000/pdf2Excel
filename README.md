# PDF表格转Excel工具

一款简单易用的工具，用于从PDF文件中提取表格并转换为Excel格式。

## 功能特点

- 从PDF文件中自动提取所有表格
- 每个表格保存为Excel文件中的单独工作表
- 实时显示处理进度和预计剩余时间
- 支持处理大型PDF文件
- 用户友好的界面
- 支持中断处理过程

## 运行环境要求

- Python 3.8或更高版本
- Java运行环境(JRE) - 用于tabula-py库

## 安装步骤

1. 安装Java运行环境(如果尚未安装)
   - 从[Java官网](https://www.java.com/)下载并安装

2. 安装Python依赖包
   ```bash
   pip install -r requirements.txt
   ```

3. 运行程序
   ```bash
   python pdf_table_converter.py
   ```

## 使用方法

1. 点击"浏览"按钮选择PDF文件
2. 点击"选择位置"按钮设置输出Excel文件的位置和名称
3. 点击"开始转换"按钮开始处理
4. 等待处理完成，可以通过进度条和状态信息查看处理进度
5. 处理完成后，可以选择直接打开输出文件所在目录

## 打包为可执行文件

### Windows
```bash
pyinstaller --onefile --windowed --name "PDF表格转Excel工具" --add-data "path\to\java;java" pdf_table_converter.py
```

### MacOS
```bash
pyinstaller --onefile --windowed --name "PDF表格转Excel工具" --add-data "path/to/java:java" pdf_table_converter.py
```

## 常见问题

1. **问题**: 提示"未检测到Java环境"  
   **解决方案**: 安装Java运行环境(JRE)，从[Java官网](https://www.java.com/)下载

2. **问题**: 处理大型PDF文件时内存不足  
   **解决方案**: 程序会自动分批处理PDF，但如果仍然出现问题，请确保系统有足够的可用内存

3. **问题**: 某些表格未被正确提取  
   **解决方案**: 表格提取基于tabula-py库，该库可能无法识别一些特殊格式的表格，特别是扫描或图片格式的表格

## 许可证

本项目使用MIT许可证 - 详情请参见LICENSE文件 