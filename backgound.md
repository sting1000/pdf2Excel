# PDF表格转Excel工具 - 完整实现方案

下面是使用PyInstaller和PySimpleGUI创建PDF表格转Excel工具的完整方案。我会提供完整代码和详细的实现步骤。

## 1. 完整源代码

创建一个名为`pdf_table_converter.py`的文件，复制以下代码:

```python
import PySimpleGUI as sg
import os
import sys
import threading
import tabula
import pandas as pd
import PyPDF2
import time
import math
from pathlib import Path

# 确保Java路径问题不会影响程序运行
os.environ["PATH"] = os.environ["PATH"] + ";" + os.path.join(os.path.dirname(sys.executable), "java")

def suppress_stdout_stderr(func):
    """装饰器：用于完全抑制函数执行过程中的stdout和stderr输出"""
    def wrapper(*args, **kwargs):
        # 保存原始的文件描述符
        try:
            # Windows和Unix兼容的标准输出/错误重定向
            original_stdout = sys.stdout
            original_stderr = sys.stderr
            
            # 重定向到空设备
            with open(os.devnull, 'w') as devnull:
                sys.stdout = devnull
                sys.stderr = devnull
                
                # 执行原函数
                result = func(*args, **kwargs)
                return result
        finally:
            # 恢复原始的stdout和stderr
            sys.stdout = original_stdout
            sys.stderr = original_stderr
    return wrapper

@suppress_stdout_stderr
def extract_tables_silent(pdf_path, page_range):
    """抑制所有输出的表格提取函数"""
    return tabula.read_pdf(pdf_path, pages=page_range, multiple_tables=True)

def convert_pdf_to_excel(pdf_path, output_path, progress_callback, cancel_flag):
    """
    将PDF中的表格转换为Excel
    
    参数:
    - pdf_path: PDF文件路径
    - output_path: 输出Excel文件路径
    - progress_callback: 进度回调函数, 接收 (percent, status_text, tables_found)
    - cancel_flag: 取消标志字典 {"cancel": False}
    """
    try:
        # 初始化进度
        progress_callback(0, "正在分析PDF文件...", 0)
        
        # 获取PDF总页数
        with open(pdf_path, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            total_pages = len(pdf_reader.pages)
        
        progress_callback(1, f"PDF共有 {total_pages} 页，开始提取表格...", 0)
        
        # 对于大型PDF，使用批处理方式
        batch_size = 10  # 每批处理10页
        total_batches = math.ceil(total_pages / batch_size)
        
        all_tables = []
        start_time = time.time()
        total_tables_found = 0
        
        # 处理每一批次
        for batch in range(total_batches):
            # 检查取消标志
            if cancel_flag.get("cancel", False):
                progress_callback(0, "操作已取消", 0)
                return False
                
            start_page = batch * batch_size + 1
            end_page = min((batch + 1) * batch_size, total_pages)
            
            # 构建页范围字符串
            page_range = f"{start_page}-{end_page}"
            
            try:
                # 更新状态
                progress_callback(
                    int(batch * 100 / total_batches), 
                    f"正在处理页面 {start_page}-{end_page} (共{total_pages}页)...", 
                    total_tables_found
                )
                
                # 提取表格
                tables = extract_tables_silent(pdf_path, page_range)
                
                if tables:
                    all_tables.extend(tables)
                    total_tables_found += len(tables)
                    
                # 计算进度百分比 (总体完成的80%用于提取，20%用于保存)
                percent = int(batch * 80 / total_batches)
                
                # 计算已用时间和预计剩余时间
                elapsed = time.time() - start_time
                if batch > 0:
                    avg_time = elapsed / batch
                    est_remaining = avg_time * (total_batches - batch)
                    est_remaining_min = est_remaining / 60
                    
                    status = f"已处理: {start_page}-{end_page}/{total_pages}页 | 找到: {total_tables_found}表格 | 剩余: {est_remaining_min:.1f}分钟"
                else:
                    status = f"已处理: {start_page}-{end_page}/{total_pages}页 | 找到: {total_tables_found}表格"
                
                progress_callback(percent, status, total_tables_found)
                
            except Exception as e:
                error_msg = f"处理页 {page_range} 时出错: {str(e)}"
                progress_callback(percent, error_msg, total_tables_found)
                # 继续处理下一批次
        
        # 保存到Excel
        if all_tables and not cancel_flag.get("cancel", False):
            progress_callback(80, f"正在保存 {total_tables_found} 个表格到Excel...", total_tables_found)
            
            with pd.ExcelWriter(output_path) as writer:
                for i, df in enumerate(all_tables):
                    # 检查取消标志
                    if cancel_flag.get("cancel", False):
                        progress_callback(0, "操作已取消", 0)
                        return False
                        
                    # 计算保存进度 (从80%到100%)
                    save_percent = 80 + int((i + 1) * 20 / len(all_tables))
                    
                    sheet_name = f"Table_{i+1}"
                    # 表格名称长度限制
                    if len(sheet_name) > 31:  # Excel工作表名称最大31字符
                        sheet_name = f"T{i+1}"
                    
                    # 检查空表格
                    if df.empty:
                        continue
                        
                    try:
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        progress_callback(
                            save_percent, 
                            f"保存表格: {i+1}/{len(all_tables)}", 
                            total_tables_found
                        )
                    except Exception as e:
                        progress_callback(
                            save_percent, 
                            f"保存表格 {i+1} 时出错: {str(e)}", 
                            total_tables_found
                        )
            
            total_time = time.time() - start_time
            progress_callback(
                100, 
                f"✅ 完成! 已保存 {total_tables_found} 个表格，用时: {total_time:.1f}秒", 
                total_tables_found
            )
            return True
        elif cancel_flag.get("cancel", False):
            progress_callback(0, "操作已取消", 0)
            return False
        else:
            progress_callback(100, "⚠️ 未找到任何表格", 0)
            return False
            
    except Exception as e:
        progress_callback(0, f"转换过程中出错: {str(e)}", 0)
        return False

def check_java_installation():
    """检查Java是否已安装"""
    try:
        import subprocess
        result = subprocess.run(['java', '-version'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return result.returncode == 0
    except:
        return False

def main():
    # 设置主题
    sg.theme('LightBlue2')
    
    # 检查Java
    if not check_java_installation():
        sg.popup_error('错误: 未检测到Java环境!\n\n此应用需要Java才能运行。\n请安装Java后再运行本程序。\n\n可以从 https://www.java.com 下载安装Java。', title='Java未安装')
    
    # 界面布局
    layout = [
        [sg.Text('PDF表格转Excel工具', font=('Any', 16), justification='center', expand_x=True)],
        [sg.Text('本工具可将PDF文件中的表格提取并保存为Excel文件', justification='center', expand_x=True)],
        [sg.HorizontalSeparator()],
        [sg.Text('选择PDF文件:')],
        [sg.Input(key='-FILE-', enable_events=True), sg.FileBrowse('浏览', file_types=(("PDF文件", "*.pdf"),))],
        [sg.Text('输出Excel文件:')],
        [sg.Input(key='-OUTPUT-', enable_events=True), sg.SaveAs('选择位置', file_types=(("Excel文件", "*.xlsx"),))],
        [sg.Text('状态:')],
        [sg.Multiline(key='-STATUS-', disabled=True, size=(65, 5), autoscroll=True)],
        [sg.Text('进度:')],
        [sg.ProgressBar(100, orientation='h', size=(40, 20), key='-PROGRESS-', expand_x=True)],
        [sg.Text('0', key='-PERCENT-'), sg.Push(), sg.Text('找到表格: 0', key='-TABLES-')],
        [sg.Button('开始转换', key='-CONVERT-', disabled=True), sg.Button('取消', key='-CANCEL-', disabled=True), sg.Push(), sg.Button('退出')]
    ]
    
    window = sg.Window('PDF表格转Excel工具', layout, finalize=True, resizable=True)
    
    # 状态变量
    conversion_thread = None
    cancel_flag = {"cancel": False}
    
    # 事件循环
    while True:
        event, values = window.read(timeout=100)
        
        # 窗口关闭
        if event == sg.WIN_CLOSED or event == '退出':
            break
            
        # 文件选择变化
        if event == '-FILE-' or event == '-OUTPUT-':
            # 只有当输入和输出都有值时启用转换按钮
            window['-CONVERT-'].update(disabled=not (values['-FILE-'] and values['-OUTPUT-']))
        
        # 开始转换
        if event == '-CONVERT-':
            pdf_path = values['-FILE-']
            output_path = values['-OUTPUT-']
            
            # 检查文件路径
            if not os.path.exists(pdf_path):
                sg.popup_error(f'PDF文件不存在: {pdf_path}')
                continue
                
            # 确保输出路径有.xlsx扩展名
            if not output_path.lower().endswith('.xlsx'):
                output_path += '.xlsx'
                window['-OUTPUT-'].update(output_path)
            
            # 检查输出目录是否存在
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                try:
                    os.makedirs(output_dir)
                except Exception as e:
                    sg.popup_error(f'无法创建输出目录: {str(e)}')
                    continue
            
            # 重置状态
            window['-STATUS-'].update('')
            window['-PROGRESS-'].update(0)
            window['-PERCENT-'].update('0%')
            window['-TABLES-'].update('找到表格: 0')
            
            # 重置取消标志
            cancel_flag["cancel"] = False
            
            # 进度回调函数
            def update_progress(percent, status_text, tables_found):
                window['-PROGRESS-'].update(percent)
                window['-PERCENT-'].update(f'{percent}%')
                window['-STATUS-'].update(status_text + '\n', append=True)
                window['-TABLES-'].update(f'找到表格: {tables_found}')
            
            # 禁用转换按钮，启用取消按钮
            window['-CONVERT-'].update(disabled=True)
            window['-CANCEL-'].update(disabled=False)
            
            # 在后台线程中处理转换
            conversion_thread = threading.Thread(
                target=convert_pdf_to_excel, 
                args=(pdf_path, output_path, update_progress, cancel_flag),
                daemon=True
            )
            conversion_thread.start()
        
        # 取消操作
        if event == '-CANCEL-':
            if conversion_thread and conversion_thread.is_alive():
                window['-STATUS-'].update('正在取消操作，请稍候...\n', append=True)
                cancel_flag["cancel"] = True
        
        # 检查转换线程是否完成
        if conversion_thread and not conversion_thread.is_alive():
            window['-CONVERT-'].update(disabled=not (values['-FILE-'] and values['-OUTPUT-']))
            window['-CANCEL-'].update(disabled=True)
            conversion_thread = None
    
    window.close()

if __name__ == '__main__':
    main()
```

## 2. 实现步骤

### 第1步：设置开发环境

1. **安装Python**:
   - 确保您安装了Python 3.8或更高版本
   - 从[python.org](https://www.python.org/downloads/)下载并安装

2. **创建并激活虚拟环境**:

   **Windows**:
   ```bash
   # 创建虚拟环境
   python -m venv pdfconverter-env
   
   # 激活虚拟环境
   pdfconverter-env\Scripts\activate
   ```

   **Mac/Linux**:
   ```bash
   # 创建虚拟环境
   python -m venv pdfconverter-env
   
   # 激活虚拟环境
   source pdfconverter-env/bin/activate
   ```

3. **安装必要的库**:
   ```bash
   pip install PySimpleGUI tabula-py pandas openpyxl PyPDF2 pyinstaller
   ```

### 第2步：创建应用程序代码

1. 创建`pdf_table_converter.py`文件，复制上面提供的完整代码
2. 保存文件

### 第3步：测试应用程序

在打包前测试应用程序：
```bash
python pdf_table_converter.py
```

确保应用程序能正常运行且所有功能正常。

### 第4步：使用PyInstaller打包应用程序

#### Windows打包命令:

```bash
# 基本打包命令
pyinstaller --onefile --windowed --name "PDF表格转Excel工具" --icon=NONE pdf_table_converter.py

# 如果遇到问题，可以尝试添加更多选项
pyinstaller --onefile --windowed --name "PDF表格转Excel工具" --add-data "venv/Lib/site-packages/tabula;tabula" --hidden-import pkg_resources.py2_warn --hidden-import jpype --icon=NONE pdf_table_converter.py
```

#### Mac打包命令:

```bash
# 基本打包命令
pyinstaller --onefile --windowed --name "PDF表格转Excel工具" pdf_table_converter.py

# 如果遇到问题，可以尝试添加更多选项
pyinstaller --onefile --windowed --name "PDF表格转Excel工具" --add-data "pdfconverter-env/lib/python3.x/site-packages/tabula:tabula" --hidden-import jpype pdf_table_converter.py
```

### 第5步：创建分发包

1. **找到生成的可执行文件**:
   - Windows: 在`dist`目录下找到`PDF表格转Excel工具.exe`
   - Mac: 在`dist`目录下找到`PDF表格转Excel工具.app`

2. **创建发布包**:
   - 创建一个包含以下内容的ZIP文件:
     - 可执行文件
     - README.txt (使用说明)
     - LICENSE.txt (如适用)

3. **编写README.txt**:
```
PDF表格转Excel工具

功能:
- 从PDF文件中提取表格并转换为Excel格式
- 支持多页PDF文件
- 显示详细进度和状态信息

使用方法:
1. 运行程序
2. 点击"浏览"选择PDF文件
3. 点击"选择位置"设置输出Excel文件位置
4. 点击"开始转换"
5. 等待处理完成

要求:
- Java运行环境 (JRE) - 如未安装，请从 https://www.java.com 下载

问题排查:
- 如果程序无法启动，请确保已安装Java
- 对于大型PDF文件，处理可能需要较长时间
```

## 3. 常见问题及解决方案

### Java依赖问题

**问题**: 用户未安装Java
**解决方案**: 程序启动时会检查Java安装，如果未安装会显示提示信息

### 内存问题

**问题**: 处理大型PDF时内存不足
**解决方案**: 程序使用批处理方式逐页处理，减少内存占用

### 打包问题

**问题**: PyInstaller无法正确包含所有依赖
**解决方案**: 使用`--add-data`和`--hidden-import`选项指定额外依赖

## 4. 测试和验证

1. 在目标平台(Windows/Mac)上安装生成的应用程序
2. 测试各种PDF文件:
   - 小型PDF (几页)
   - 中型PDF (几十页)
   - 大型PDF (上百页)
3. 验证提取的表格是否正确

## 5. 发布

1. 在GitHub创建Release
2. 上传Windows和Mac版本的安装包
3. 提供详细的安装和使用说明

## 总结

这个完整的解决方案提供了:
- 用户友好的界面
- 详细的进度显示
- 取消功能
- 错误处理
- 跨平台兼容性
- 打包和部署指南

按照上述步骤操作，您应该能够成功创建一个可分发的PDF表格转Excel工具，方便Windows和Mac用户使用。
