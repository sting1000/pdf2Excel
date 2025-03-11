import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
import threading
import tabula
import pandas as pd
import PyPDF2
import time
import math
import io
import multiprocessing
import platform
import gc
import psutil
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor, ProcessPoolExecutor
from pathlib import Path
import ctypes

# 内存管理器
class MemoryManager:
    """内存使用监控和管理"""
    
    @staticmethod
    def get_memory_usage():
        """获取当前进程内存使用量（MB）"""
        process = psutil.Process(os.getpid())
        memory_info = process.memory_info()
        return memory_info.rss / 1024 / 1024  # 转换为MB
    
    @staticmethod
    def free_memory():
        """强制垃圾回收，释放内存"""
        # 调用多次垃圾回收
        gc.collect(0)  # 收集第0代（最年轻的对象）
        gc.collect(1)  # 收集第1代
        gc.collect(2)  # 收集第2代（最老的对象）
        
        # 尝试释放未使用的内存返回给OS
        if hasattr(os, 'malloc_trim'):  # Linux特有
            os.malloc_trim(0)
        elif sys.platform == 'darwin':  # macOS
            libc = ctypes.CDLL('libc.dylib')
            if hasattr(libc, 'malloc_zone_pressure_relief'):
                # 释放100MB内存
                libc.malloc_zone_pressure_relief(None, 100)
        
        # 强制Python释放未使用的内存池
        import multiprocessing
        p = multiprocessing.Process(target=lambda: None)
        p.start()
        p.join()
        
        return gc.get_count()[0]  # 返回回收的对象数量
    
    @staticmethod
    def print_memory_status():
        """输出当前内存状态"""
        mem_usage = MemoryManager.get_memory_usage()
        print(f"当前内存使用: {mem_usage:.2f} MB")
        
    @staticmethod
    def check_and_free_memory(threshold=1000):
        """
        检查内存使用，如果超过阈值则尝试释放
        threshold: 内存使用阈值，单位MB
        """
        mem_usage = MemoryManager.get_memory_usage()
        if mem_usage > threshold:
            print(f"内存使用超过阈值 ({mem_usage:.2f} MB > {threshold} MB)，尝试释放内存...")
            collected = MemoryManager.free_memory()
            new_usage = MemoryManager.get_memory_usage()
            print(f"已释放对象数: {collected}, 当前内存使用: {new_usage:.2f} MB")
            return True
        return False

# 确保Java路径问题不会影响程序运行
if sys.platform == 'win32':
    os.environ["PATH"] = os.environ["PATH"] + ";" + os.path.join(os.path.dirname(sys.executable), "java")
else:
    os.environ["PATH"] = os.environ["PATH"] + ":" + os.path.join(os.path.dirname(sys.executable), "java")

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
    """静默提取表格，不输出任何信息"""
    try:
        # 在这里添加文件句柄管理
        with open(pdf_path, 'rb') as pdf_file:
            # 使用打开的文件对象而不是路径
            tables = tabula.read_pdf(pdf_file, pages=page_range, multiple_tables=True, silent=True)
            # 立即释放文件
            pdf_file.close()
        return tables
    except Exception as e:
        print(f"表格提取错误: {str(e)}")
        # 确保返回空列表而不是None
        return []

def process_batch(args):
    """处理单个PDF批次的函数，用于并行处理"""
    pdf_path, start_page, end_page = args
    page_range = f"{start_page}-{end_page}"
    try:
        tables = extract_tables_silent(pdf_path, page_range)
        result = (tables, len(tables) if tables else 0)
        # 释放内存
        MemoryManager.check_and_free_memory(threshold=500)
        return result
    except Exception as e:
        print(f"处理页 {page_range} 出错: {str(e)}")
        return [], 0

def optimize_dataframe(df):
    """优化DataFrame内存使用"""
    # 对象类型列转换为类别类型
    for col in df.select_dtypes(include=['object']).columns:
        if df[col].nunique() < len(df[col]) * 0.5:  # 如果唯一值少于50%
            df[col] = df[col].astype('category')
    
    # 将浮点数列转换为最合适的数值类型
    for col in df.select_dtypes(include=['float']).columns:
        df[col] = pd.to_numeric(df[col], downcast='float')
    
    # 将整数列转换为最合适的整数类型
    for col in df.select_dtypes(include=['int']).columns:
        df[col] = pd.to_numeric(df[col], downcast='integer')
    
    return df

def save_tables_chunk(args):
    """并行保存表格分块到Excel"""
    tables_chunk, start_idx, output_file = args
    try:
        # 内存使用监控
        start_mem = MemoryManager.get_memory_usage()
        
        # 使用内存中的Excel writer，移除options参数
        with pd.ExcelWriter(output_file, engine='openpyxl', mode='a' if os.path.exists(output_file) else 'w') as writer:
            for i, df in enumerate(tables_chunk):
                idx = start_idx + i
                sheet_name = f"Table_{idx+1}"
                # 表格名称长度限制
                if len(sheet_name) > 31:  # Excel工作表名称最大31字符
                    sheet_name = f"T{idx+1}"
                
                # 检查空表格
                if df.empty:
                    continue
                
                # 内存优化
                try:
                    # 优化DataFrame内存
                    df = optimize_dataframe(df)
                    
                    # 设置Excel选项，减少内存使用
                    df.to_excel(writer, sheet_name=sheet_name, index=False, engine='openpyxl')
                    
                    # 显式删除DataFrame以释放内存
                    del df
                    
                    # 每5个表格检查一次内存
                    if (i + 1) % 5 == 0:
                        MemoryManager.check_and_free_memory(threshold=500)
                        
                except Exception as e:
                    print(f"保存表格 {idx+1} 时出错: {str(e)}")
        
        # 显式垃圾回收
        end_mem = MemoryManager.get_memory_usage()
        print(f"保存批次内存使用: {start_mem:.2f} MB -> {end_mem:.2f} MB, 差异: {end_mem-start_mem:.2f} MB")
        MemoryManager.free_memory()
        
        return True, len(tables_chunk)
    except Exception as e:
        print(f"保存批次出错: {str(e)}")
        return False, 0

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
        
        # 优化：根据系统可用核心数和PDF大小确定批处理大小和并行程度
        cpu_count = multiprocessing.cpu_count()
        workers = max(1, min(cpu_count - 1, 4))  # 保留至少一个核心给系统
        
        # 批处理大小，根据PDF大小动态调整
        if total_pages > 10000:
            batch_size = 500  # 超大PDF
        elif total_pages > 1000:
            batch_size = 100  # 大型PDF
        elif total_pages > 100:
            batch_size = 50   # 中型PDF
        else:
            batch_size = 20   # 小型PDF
        
        total_batches = math.ceil(total_pages / batch_size)
        
        # 创建批处理任务列表
        batches = []
        for batch in range(total_batches):
            if cancel_flag.get("cancel", False):
                progress_callback(0, "操作已取消", 0)
                return False
                
            start_page = batch * batch_size + 1
            end_page = min((batch + 1) * batch_size, total_pages)
            batches.append((pdf_path, start_page, end_page))
        
        all_tables = []
        start_time = time.time()
        total_tables_found = 0
        completed_batches = 0
        
        # 使用线程池同时处理多个批次
        with ThreadPoolExecutor(max_workers=workers) as executor:
            # 提交所有批次任务
            future_to_batch = {executor.submit(process_batch, batch): i for i, batch in enumerate(batches)}
            
            # 处理完成的任务结果
            for future in concurrent.futures.as_completed(future_to_batch):
                if cancel_flag.get("cancel", False):
                    executor.shutdown(wait=False)
                    progress_callback(0, "操作已取消", 0)
                    return False
                
                batch_index = future_to_batch[future]
                start_page = batches[batch_index][1]
                end_page = batches[batch_index][2]
                
                try:
                    tables, tables_count = future.result()
                    if tables:
                        all_tables.extend(tables)
                        total_tables_found += tables_count
                except Exception as e:
                    progress_callback(
                        int(completed_batches * 80 / total_batches),
                        f"处理页 {start_page}-{end_page} 时出错: {str(e)}",
                        total_tables_found
                    )
                
                completed_batches += 1
                
                # 计算进度百分比 (总体完成的80%用于提取，20%用于保存)
                percent = int(completed_batches * 80 / total_batches)
                
                # 计算已用时间和预计剩余时间
                elapsed = time.time() - start_time
                if completed_batches > 0:
                    avg_time = elapsed / completed_batches
                    est_remaining = avg_time * (total_batches - completed_batches)
                    est_remaining_min = est_remaining / 60
                    
                    # 更精确的时间估计
                    if est_remaining_min > 60:
                        time_str = f"{est_remaining_min/60:.1f}小时"
                    else:
                        time_str = f"{est_remaining_min:.1f}分钟"
                    
                    status = f"已处理: {completed_batches}/{total_batches}批次 ({start_page}-{end_page}/{total_pages}页) | 找到: {total_tables_found}表格 | 剩余: {time_str}"
                else:
                    status = f"已处理: {completed_batches}/{total_batches}批次 ({start_page}-{end_page}/{total_pages}页) | 找到: {total_tables_found}表格"
                
                progress_callback(percent, status, total_tables_found)
                
                # 每完成5个批次检查一次内存
                if completed_batches % 5 == 0:
                    MemoryManager.check_and_free_memory(threshold=800)
        
        # 保存到Excel - 采用并行分块保存方式
        if all_tables and not cancel_flag.get("cancel", False):
            progress_callback(80, f"正在保存 {total_tables_found} 个表格到Excel...", total_tables_found)
            
            # 创建空的输出文件
            if os.path.exists(output_path):
                os.remove(output_path)
            
            # 分块保存，每块最多100个表格（对于大量表格的情况，减小块大小以避免内存问题）
            chunk_size = min(100, max(50, 10000 // total_tables_found + 1)) if total_tables_found > 0 else 100
            num_chunks = math.ceil(len(all_tables) / chunk_size)
            save_start_time = time.time()
            saved_tables = 0
            
            # 准备保存任务
            save_tasks = []
            for i in range(num_chunks):
                start_idx = i * chunk_size
                end_idx = min((i + 1) * chunk_size, len(all_tables))
                chunk = all_tables[start_idx:end_idx]
                save_tasks.append((chunk, start_idx, output_path))
            
            # 清空all_tables并触发垃圾回收以释放内存
            all_tables = None
            MemoryManager.free_memory()
            
            # 使用线程池并行保存，但限制并行度以控制内存使用
            max_save_workers = min(2, num_chunks)
            with ThreadPoolExecutor(max_workers=max_save_workers) as save_executor:
                # 一次提交少量任务，避免内存溢出
                batch_size = 5
                for batch_start in range(0, len(save_tasks), batch_size):
                    batch_end = min(batch_start + batch_size, len(save_tasks))
                    batch_tasks = save_tasks[batch_start:batch_end]
                    
                    future_to_save = {save_executor.submit(save_tables_chunk, task): i+batch_start for i, task in enumerate(batch_tasks)}
                    
                    for future in concurrent.futures.as_completed(future_to_save):
                        if cancel_flag.get("cancel", False):
                            save_executor.shutdown(wait=False)
                            progress_callback(0, "操作已取消", 0)
                            return False
                        
                        task_index = future_to_save[future]
                        try:
                            success, num_saved = future.result()
                            if success:
                                saved_tables += num_saved
                                
                                # 计算保存进度 (从80%到100%)
                                save_percent = 80 + int((task_index + 1) * 20 / len(save_tasks))
                                
                                # 预估剩余时间
                                save_elapsed = time.time() - save_start_time
                                if task_index > 0:
                                    avg_save_time = save_elapsed / (task_index + 1)
                                    save_remaining = avg_save_time * (len(save_tasks) - task_index - 1)
                                    save_remaining_min = save_remaining / 60
                                    
                                    if save_remaining_min > 60:
                                        save_time_str = f"{save_remaining_min/60:.1f}小时"
                                    else:
                                        save_time_str = f"{save_remaining_min:.1f}分钟"
                                    
                                    progress_callback(
                                        save_percent,
                                        f"保存进度: {task_index+1}/{len(save_tasks)}批次 | 已保存: {saved_tables}/{total_tables_found}表格 | 剩余: {save_time_str}",
                                        total_tables_found
                                    )
                                else:
                                    progress_callback(
                                        save_percent,
                                        f"保存进度: {task_index+1}/{len(save_tasks)}批次 | 已保存: {saved_tables}/{total_tables_found}表格",
                                        total_tables_found
                                    )
                        except Exception as e:
                            progress_callback(
                                80 + int(task_index * 20 / len(save_tasks)),
                                f"保存批次 {task_index+1} 时出错: {str(e)}",
                                total_tables_found
                            )
                    
                    # 每批次完成后触发内存清理
                    MemoryManager.free_memory()
            
            # 清理内存
            MemoryManager.free_memory()
            
            total_time = time.time() - start_time
            # 格式化总时间
            if total_time > 3600:
                time_str = f"{total_time/3600:.2f}小时"
            elif total_time > 60:
                time_str = f"{total_time/60:.2f}分钟"
            else:
                time_str = f"{total_time:.1f}秒"
                
            progress_callback(
                100, 
                f"✅ 完成! 已保存 {saved_tables} 个表格，用时: {time_str}", 
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

class PDFTableConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF表格转Excel工具")
        self.root.geometry("750x600")
        
        # 使用系统默认字体，解决文字显示问题
        system = platform.system()
        if system == "Darwin":  # macOS
            self.default_font = ('SF Pro', 12)
            self.header_font = ('SF Pro', 20, 'bold')
            self.subheader_font = ('SF Pro', 14)
        elif system == "Windows":
            self.default_font = ('Microsoft YaHei UI', 10)
            self.header_font = ('Microsoft YaHei UI', 18, 'bold')
            self.subheader_font = ('Microsoft YaHei UI', 12)
        else:  # Linux和其他系统
            self.default_font = ('Noto Sans CJK SC', 10)
            self.header_font = ('Noto Sans CJK SC', 18, 'bold')
            self.subheader_font = ('Noto Sans CJK SC', 12)
        
        # 明确设置背景颜色
        self.bg_color = "#F0F0F0"
        self.frame_bg = "#F8F8F8"
        self.text_color = "#333333"
        
        # 应用全局样式
        self.root.configure(bg=self.bg_color)
        
        # 创建主框架
        self.main_frame = tk.Frame(root, bg=self.bg_color, padx=20, pady=20)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        self.title_label = tk.Label(self.main_frame, text="PDF表格转Excel工具", 
                                   font=self.header_font, bg=self.bg_color, fg=self.text_color)
        self.title_label.pack(pady=(0, 10))
        
        self.subtitle_label = tk.Label(self.main_frame, text="将PDF文件中的表格快速提取并转换为Excel格式", 
                                      font=self.subheader_font, bg=self.bg_color, fg=self.text_color)
        self.subtitle_label.pack(pady=(0, 20))
        
        # 分隔线
        separator = tk.Frame(self.main_frame, height=2, bg="#CCCCCC")
        separator.pack(fill="x", pady=10)
        
        # 文件选择框架
        self.file_frame = tk.LabelFrame(self.main_frame, text="文件选择", font=self.default_font,
                                      bg=self.frame_bg, fg=self.text_color, padx=15, pady=15)
        self.file_frame.pack(fill="x", pady=10)
        
        # PDF文件选择
        self.pdf_frame = tk.Frame(self.file_frame, bg=self.frame_bg)
        self.pdf_frame.pack(fill="x", pady=5)
        
        self.pdf_label = tk.Label(self.pdf_frame, text="选择PDF文件:", font=self.default_font,
                                bg=self.frame_bg, fg=self.text_color, width=15)
        self.pdf_label.pack(side=tk.LEFT)
        
        self.pdf_path_var = tk.StringVar()
        self.pdf_entry = tk.Entry(self.pdf_frame, textvariable=self.pdf_path_var, 
                                font=self.default_font, width=50)
        self.pdf_entry.pack(side=tk.LEFT, fill="x", expand=True, padx=5)
        
        self.browse_button = tk.Button(self.pdf_frame, text="浏览", font=self.default_font,
                                     command=self.browse_pdf, bg="#4A86E8", fg="white",
                                     activebackground="#2A66C8", activeforeground="white",
                                     relief=tk.RAISED, bd=1)
        self.browse_button.pack(side=tk.LEFT)
        
        # 输出Excel文件选择
        self.excel_frame = tk.Frame(self.file_frame, bg=self.frame_bg)
        self.excel_frame.pack(fill="x", pady=5)
        
        self.excel_label = tk.Label(self.excel_frame, text="输出Excel文件:", font=self.default_font,
                                  bg=self.frame_bg, fg=self.text_color, width=15)
        self.excel_label.pack(side=tk.LEFT)
        
        self.excel_path_var = tk.StringVar()
        self.excel_entry = tk.Entry(self.excel_frame, textvariable=self.excel_path_var, 
                                  font=self.default_font, width=50)
        self.excel_entry.pack(side=tk.LEFT, fill="x", expand=True, padx=5)
        
        self.save_button = tk.Button(self.excel_frame, text="选择位置", font=self.default_font,
                                   command=self.save_excel, bg="#4A86E8", fg="white",
                                   activebackground="#2A66C8", activeforeground="white",
                                   relief=tk.RAISED, bd=1)
        self.save_button.pack(side=tk.LEFT)
        
        # 处理状态框架
        self.status_frame = tk.LabelFrame(self.main_frame, text="处理状态", font=self.default_font,
                                        bg=self.frame_bg, fg=self.text_color, padx=15, pady=15)
        self.status_frame.pack(fill="both", expand=True, pady=10)
        
        # 进度条
        self.progress_label = tk.Label(self.status_frame, text="进度:", font=self.default_font,
                                     bg=self.frame_bg, fg=self.text_color)
        self.progress_label.pack(anchor="w", pady=(0, 5))
        
        # 创建自定义样式的进度条
        self.progress_var = tk.DoubleVar()
        self.progress_frame = tk.Frame(self.status_frame, height=20, bg="#DDDDDD", bd=1, relief=tk.SUNKEN)
        self.progress_frame.pack(fill="x", pady=(0, 5))
        self.progress_frame.pack_propagate(False)
        
        self.progress_bar = tk.Frame(self.progress_frame, bg="#4CAF50", width=0)
        self.progress_bar.place(relx=0, rely=0, relheight=1, relwidth=0)
        
        # 进度信息框架
        self.info_frame = tk.Frame(self.status_frame, bg=self.frame_bg)
        self.info_frame.pack(fill="x")
        
        self.percent_var = tk.StringVar(value="0%")
        self.percent_label = tk.Label(self.info_frame, textvariable=self.percent_var, 
                                    font=self.default_font, bg=self.frame_bg, fg=self.text_color, width=5)
        self.percent_label.pack(side=tk.LEFT)
        
        self.tables_var = tk.StringVar(value="找到表格: 0")
        self.tables_label = tk.Label(self.info_frame, textvariable=self.tables_var, 
                                   font=self.default_font, bg=self.frame_bg, fg=self.text_color)
        self.tables_label.pack(side=tk.RIGHT)
        
        # 状态文本框
        self.status_text = tk.Text(self.status_frame, height=8, wrap=tk.WORD, 
                                 font=self.default_font, state="disabled", 
                                 bg="#FFFFFF", bd=1, relief=tk.SOLID)
        self.status_text.pack(fill="both", expand=True, pady=10)
        
        # 按钮框架
        self.button_frame = tk.Frame(self.main_frame, bg=self.bg_color)
        self.button_frame.pack(fill="x", pady=10)
        
        # 添加空白框架作为弹簧
        self.spacer1 = tk.Frame(self.button_frame, bg=self.bg_color)
        self.spacer1.pack(side=tk.LEFT, fill="x", expand=True)
        
        self.convert_button = tk.Button(self.button_frame, text="开始转换", font=(self.default_font[0], 10, 'bold'),
                                      command=self.start_conversion, state="disabled", width=15,
                                      bg="#4CAF50", fg="white", activebackground="#3D8B40", 
                                      activeforeground="white", relief=tk.RAISED, bd=2, highlightthickness=0,
                                      padx=10, pady=5)
        self.convert_button.pack(side=tk.LEFT, padx=5)
        
        self.cancel_button = tk.Button(self.button_frame, text="取消", font=self.default_font,
                                     command=self.cancel_conversion, state="disabled", width=15,
                                     bg="#F44336", fg="white", activebackground="#D32F2F", 
                                     activeforeground="white", relief=tk.RAISED, bd=1)
        self.cancel_button.pack(side=tk.LEFT, padx=5)
        
        self.exit_button = tk.Button(self.button_frame, text="退出", font=self.default_font,
                                   command=self.exit_app, width=15, bg="#9E9E9E", fg="white",
                                   activebackground="#757575", activeforeground="white",
                                   relief=tk.RAISED, bd=1)
        self.exit_button.pack(side=tk.LEFT, padx=5)
        
        # 添加空白框架作为弹簧
        self.spacer2 = tk.Frame(self.button_frame, bg=self.bg_color)
        self.spacer2.pack(side=tk.LEFT, fill="x", expand=True)
        
        # 绑定事件
        self.pdf_path_var.trace_add("write", self.update_button_states)
        self.excel_path_var.trace_add("write", self.update_button_states)
        
        # 状态变量
        self.conversion_thread = None
        self.cancel_flag = {"cancel": False}
        self.last_dir = os.path.expanduser("~")
        
        # 检查Java
        if not check_java_installation():
            messagebox.showerror("Java未安装", 
                               "错误: 未检测到Java环境!\n\n此应用需要Java才能运行。\n请安装Java后再运行本程序。\n\n可以从 https://www.java.com 下载安装Java。")
    
    def update_progress_bar(self, percent):
        """更新进度条显示"""
        self.progress_bar.place(relwidth=percent/100)
    
    def browse_pdf(self):
        file_path = filedialog.askopenfilename(
            title="选择PDF文件",
            initialdir=self.last_dir,
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        
        if file_path:
            self.pdf_path_var.set(file_path)
            self.last_dir = os.path.dirname(file_path)
            
            # 自动提供默认输出文件名
            if not self.excel_path_var.get():
                base_name = os.path.basename(file_path)
                output_name = os.path.splitext(base_name)[0] + '.xlsx'
                output_path = os.path.join(self.last_dir, output_name)
                self.excel_path_var.set(output_path)
    
    def save_excel(self):
        file_path = filedialog.asksaveasfilename(
            title="保存Excel文件",
            initialdir=self.last_dir,
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        
        if file_path:
            self.excel_path_var.set(file_path)
            self.last_dir = os.path.dirname(file_path)
    
    def update_button_states(self, *args):
        if self.pdf_path_var.get() and self.excel_path_var.get():
            self.convert_button["state"] = "normal"
        else:
            self.convert_button["state"] = "disabled"
    
    def update_status_text(self, text):
        self.status_text.configure(state="normal")
        self.status_text.insert(tk.END, text + "\n")
        self.status_text.see(tk.END)
        self.status_text.configure(state="disabled")
    
    def update_progress(self, percent, status_text, tables_found):
        # 使用after方法确保在主线程中更新UI
        self.root.after(0, lambda: self._update_progress_impl(percent, status_text, tables_found))
    
    def _update_progress_impl(self, percent, status_text, tables_found):
        self.update_progress_bar(percent)
        self.percent_var.set(f"{percent}%")
        self.update_status_text(status_text)
        self.tables_var.set(f"找到表格: {tables_found}")
    
    def start_conversion(self):
        pdf_path = self.pdf_path_var.get()
        output_path = self.excel_path_var.get()
        
        # 检查文件路径
        if not os.path.exists(pdf_path):
            messagebox.showerror("错误", f"PDF文件不存在: {pdf_path}")
            return
        
        # 确保输出路径有.xlsx扩展名
        if not output_path.lower().endswith('.xlsx'):
            output_path += '.xlsx'
            self.excel_path_var.set(output_path)
        
        # 检查输出目录是否存在
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except Exception as e:
                messagebox.showerror("错误", f"无法创建输出目录: {str(e)}")
                return
        
        # 检查输出文件是否已存在
        if os.path.exists(output_path):
            if not messagebox.askyesno("文件已存在", 
                                     f"文件 {os.path.basename(output_path)} 已存在，是否覆盖?"):
                return
        
        # 重置状态
        self.status_text.configure(state="normal")
        self.status_text.delete(1.0, tk.END)
        self.status_text.configure(state="disabled")
        self.update_progress_bar(0)
        self.percent_var.set("0%")
        self.tables_var.set("找到表格: 0")
        
        # 重置取消标志
        self.cancel_flag["cancel"] = False
        
        # 更新按钮状态
        self.convert_button["state"] = "disabled"
        self.cancel_button["state"] = "normal"
        
        # 在后台线程中处理转换
        self.conversion_thread = threading.Thread(
            target=convert_pdf_to_excel,
            args=(pdf_path, output_path, self.update_progress, self.cancel_flag),
            daemon=True
        )
        self.conversion_thread.start()
        
        # 定期检查线程是否完成
        self.root.after(100, self.check_conversion_thread)
    
    def check_conversion_thread(self):
        if self.conversion_thread and not self.conversion_thread.is_alive():
            # 恢复按钮状态
            self.convert_button["state"] = "normal" if (self.pdf_path_var.get() and self.excel_path_var.get()) else "disabled"
            self.cancel_button["state"] = "disabled"
            
            # 检查是否成功完成转换
            if os.path.exists(self.excel_path_var.get()) and not self.cancel_flag["cancel"]:
                if messagebox.askyesno("转换完成", "转换完成！是否打开输出文件所在目录?"):
                    # 打开文件所在目录
                    self.open_output_dir()
            
            self.conversion_thread = None
        elif self.conversion_thread:
            # 继续检查
            self.root.after(100, self.check_conversion_thread)
    
    def open_output_dir(self):
        output_dir = os.path.dirname(self.excel_path_var.get())
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
            messagebox.showerror("错误", f"无法打开目录: {str(e)}")
    
    def cancel_conversion(self):
        if self.conversion_thread and self.conversion_thread.is_alive():
            self.update_status_text("正在取消操作，请稍候...")
            self.cancel_flag["cancel"] = True
    
    def exit_app(self):
        # 如果有转换线程正在运行，询问是否确定退出
        if self.conversion_thread and self.conversion_thread.is_alive():
            if not messagebox.askyesno("确认退出", "转换正在进行中，确定要退出吗？"):
                return
            self.cancel_flag["cancel"] = True
        
        self.root.destroy()

def main():
    # 确保导入concurrent.futures
    global concurrent
    import concurrent.futures
    
    # 设置Tk应用
    root = tk.Tk()
    # 设置DPI感知
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
    
    app = PDFTableConverterApp(root)
    root.mainloop()

if __name__ == '__main__':
    main() 