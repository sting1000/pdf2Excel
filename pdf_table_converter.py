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
    
    # 优化：添加应用图标（如果有）
    app_icon = sg.DEFAULT_BASE64_ICON if not os.path.exists("icon.ico") else "icon.ico"
    
    # 优化：界面布局，调整为更现代的布局
    layout = [
        [sg.Text('PDF表格转Excel工具', font=('Any', 20, 'bold'), justification='center', expand_x=True, pad=(0, 10))],
        [sg.Text('将PDF文件中的表格快速提取并转换为Excel格式', justification='center', expand_x=True, pad=(0, 15))],
        [sg.HorizontalSeparator()],
        [sg.Frame('文件选择', [
            [sg.Text('选择PDF文件:', size=(15, 1)), 
             sg.Input(key='-FILE-', enable_events=True, size=(45, 1)), 
             sg.FileBrowse('浏览', file_types=(("PDF文件", "*.pdf"),), button_color=('white', '#1E88E5'))],
            [sg.Text('输出Excel文件:', size=(15, 1)), 
             sg.Input(key='-OUTPUT-', enable_events=True, size=(45, 1)), 
             sg.SaveAs('选择位置', file_types=(("Excel文件", "*.xlsx"),), button_color=('white', '#1E88E5'))]
        ], pad=(0, 15), relief=sg.RELIEF_GROOVE)],
        [sg.Frame('处理状态', [
            [sg.Text('进度:')],
            [sg.ProgressBar(100, orientation='h', size=(50, 20), key='-PROGRESS-', expand_x=True)],
            [sg.Text('0%', key='-PERCENT-', size=(5, 1)), sg.Push(), sg.Text('找到表格: 0', key='-TABLES-')],
            [sg.Multiline(key='-STATUS-', disabled=True, size=(65, 5), autoscroll=True, background_color='#F0F0F0')]
        ], pad=(0, 15), relief=sg.RELIEF_GROOVE)],
        [sg.Push(), 
         sg.Button('开始转换', key='-CONVERT-', disabled=True, size=(15, 1), button_color=('white', '#4CAF50'), font=('Any', 10, 'bold')),
         sg.Button('取消', key='-CANCEL-', disabled=True, size=(15, 1), button_color=('white', '#F44336')),
         sg.Button('退出', size=(15, 1), button_color=('white', '#9E9E9E')),
         sg.Push()]
    ]
    
    # 创建窗口，添加图标和调整大小
    window = sg.Window('PDF表格转Excel工具', layout, finalize=True, resizable=True, 
                      element_justification='center', icon=app_icon, size=(700, 550))
    
    # 优化：记住上次使用的目录
    last_dir = os.path.expanduser("~")
    
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
        if event == '-FILE-':
            # 记住文件目录
            if values['-FILE-']:
                last_dir = os.path.dirname(values['-FILE-'])
                
                # 自动提供默认输出文件名
                if not values['-OUTPUT-']:
                    base_name = os.path.basename(values['-FILE-'])
                    output_name = os.path.splitext(base_name)[0] + '.xlsx'
                    output_path = os.path.join(last_dir, output_name)
                    window['-OUTPUT-'].update(output_path)
            
            # 更新按钮状态
            window['-CONVERT-'].update(disabled=not (values['-FILE-'] and values['-OUTPUT-']))
            
        # 输出文件变化
        if event == '-OUTPUT-':
            # 更新按钮状态
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
            
            # 检查输出文件是否已存在
            if os.path.exists(output_path):
                if not sg.popup_yes_no(f'文件 {os.path.basename(output_path)} 已存在，是否覆盖?', 
                                      title='文件已存在') == 'Yes':
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
            
            # 添加完成后操作：询问是否打开输出目录
            if os.path.exists(values['-OUTPUT-']) and not cancel_flag["cancel"]:
                if sg.popup_yes_no('转换完成！是否打开输出文件所在目录?', title='转换完成') == 'Yes':
                    # 打开文件所在目录
                    output_dir = os.path.dirname(values['-OUTPUT-'])
                    try:
                        if sys.platform == 'win32':
                            os.startfile(output_dir)
                        elif sys.platform == 'darwin':  # macOS
                            import subprocess
                            subprocess.Popen(['open', output_dir])
                        else:  # Linux
                            import subprocess
                            subprocess.Popen(['xdg-open', output_dir])
                    except Exception as e:
                        sg.popup_error(f'无法打开目录: {str(e)}')
            
            conversion_thread = None
    
    window.close()

if __name__ == '__main__':
    main() 