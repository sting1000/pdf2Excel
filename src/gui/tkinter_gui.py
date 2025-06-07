#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Tkinter GUI模块
提供基于Tkinter的图形用户界面
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
import threading
import time
from pathlib import Path
from typing import Optional, Callable

from ..core.pdf_processor import pdf_processor
from ..core.excel_writer import excel_writer
from ..utils.logger import logger
from ..utils.config import config


class TkinterGUI:
    """Tkinter图形用户界面类"""
    
    def __init__(self):
        """初始化GUI"""
        self.root = None
        self.pdf_path = None
        self.output_path = None
        self.progress_var = None
        self.status_text = None
        self.tables_found_var = None
        
        # 转换相关变量
        self.conversion_thread = None
        self.is_converting = False
        
        logger.info("TkinterGUI 初始化完成")
    
    def create_main_window(self):
        """创建主窗口"""
        self.root = tk.Tk()
        self.root.title(f"{config.get('app.name', 'PDF表格转Excel工具')} v{config.get('app.version', '2.0.0')}")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # 创建tkinter变量（必须在root窗口创建后）
        self.pdf_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.progress_var = tk.IntVar()
        self.status_text = tk.StringVar()
        self.tables_found_var = tk.IntVar()
        
        # 设置窗口图标（如果有的话）
        try:
            icon_path = Path(__file__).parent.parent.parent / "docs" / "icon.ico"
            if icon_path.exists():
                self.root.iconbitmap(str(icon_path))
        except:
            pass
        
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 创建界面元素
        self.create_file_selection_section(main_frame)
        self.create_progress_section(main_frame)
        self.create_control_buttons(main_frame)
        self.create_status_section(main_frame)
        
        # 绑定事件
        self.bind_events()
        
        # 初始化状态
        self.update_button_states()
        self.status_text.set("就绪")
        
        logger.info("主窗口创建完成")
    
    def create_file_selection_section(self, parent):
        """创建文件选择区域"""
        # PDF文件选择
        pdf_frame = ttk.LabelFrame(parent, text="选择PDF文件", padding="10")
        pdf_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        pdf_frame.columnconfigure(1, weight=1)
        
        ttk.Label(pdf_frame, text="PDF文件:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        pdf_entry = ttk.Entry(pdf_frame, textvariable=self.pdf_path, width=50)
        pdf_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(pdf_frame, text="浏览", command=self.browse_pdf).grid(row=0, column=2, sticky=tk.W)
        
        # 输出文件选择
        output_frame = ttk.LabelFrame(parent, text="输出Excel文件", padding="10")
        output_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        output_frame.columnconfigure(1, weight=1)
        
        ttk.Label(output_frame, text="Excel文件:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        output_entry = ttk.Entry(output_frame, textvariable=self.output_path, width=50)
        output_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(output_frame, text="另存为", command=self.save_excel).grid(row=0, column=2, sticky=tk.W)
    
    def create_progress_section(self, parent):
        """创建进度显示区域"""
        progress_frame = ttk.LabelFrame(parent, text="转换进度", padding="10")
        progress_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)
        
        # 进度条
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            variable=self.progress_var, 
            maximum=100,
            length=400,
            mode='determinate'
        )
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # 进度信息框架
        info_frame = ttk.Frame(progress_frame)
        info_frame.grid(row=1, column=0, sticky=(tk.W, tk.E))
        info_frame.columnconfigure(1, weight=1)
        
        # 进度百分比
        ttk.Label(info_frame, text="进度:").grid(row=0, column=0, sticky=tk.W)
        self.progress_label = ttk.Label(info_frame, text="0%")
        self.progress_label.grid(row=0, column=1, sticky=tk.W, padx=(5, 20))
        
        # 找到的表格数量
        ttk.Label(info_frame, text="已发现表格:").grid(row=0, column=2, sticky=tk.W)
        self.tables_label = ttk.Label(info_frame, text="0")
        self.tables_label.grid(row=0, column=3, sticky=tk.W, padx=(5, 0))
    
    def create_control_buttons(self, parent):
        """创建控制按钮"""
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=3, column=0, columnspan=3, pady=(0, 10))
        
        # 开始转换按钮
        self.convert_button = ttk.Button(
            button_frame, 
            text="开始转换", 
            command=self.start_conversion,
            style='Accent.TButton'
        )
        self.convert_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # 取消转换按钮
        self.cancel_button = ttk.Button(
            button_frame, 
            text="取消转换", 
            command=self.cancel_conversion,
            state='disabled'
        )
        self.cancel_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # 打开输出目录按钮
        self.open_dir_button = ttk.Button(
            button_frame, 
            text="打开输出目录", 
            command=self.open_output_dir,
            state='disabled'
        )
        self.open_dir_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # 退出按钮
        ttk.Button(
            button_frame, 
            text="退出", 
            command=self.exit_app
        ).pack(side=tk.RIGHT)
    
    def create_status_section(self, parent):
        """创建状态显示区域"""
        status_frame = ttk.LabelFrame(parent, text="状态信息", padding="10")
        status_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(0, weight=1)
        
        # 状态文本框
        self.status_textbox = tk.Text(
            status_frame, 
            height=8, 
            wrap=tk.WORD,
            state='disabled',
            bg='#f0f0f0'
        )
        self.status_textbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 滚动条
        scrollbar = ttk.Scrollbar(status_frame, orient=tk.VERTICAL, command=self.status_textbox.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.status_textbox.config(yscrollcommand=scrollbar.set)
        
        # 确保状态区域可以扩展
        parent.rowconfigure(4, weight=1)
    
    def bind_events(self):
        """绑定事件"""
        # 文件路径变化时更新按钮状态
        self.pdf_path.trace('w', self.update_button_states)
        self.output_path.trace('w', self.update_button_states)
        
        # 窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.exit_app)
        
        # 进度变量变化时更新显示
        self.progress_var.trace('w', self.update_progress_display)
        self.tables_found_var.trace('w', self.update_tables_display)
    
    def browse_pdf(self):
        """浏览并选择PDF文件"""
        try:
            filename = filedialog.askopenfilename(
                title="选择PDF文件",
                filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")],
                initialdir=config.get('gui.default_input_dir', str(Path.home()))
            )
            
            if filename:
                self.pdf_path.set(filename)
                
                # 自动设置输出文件名
                if not self.output_path.get():
                    pdf_file = Path(filename)
                    output_file = pdf_file.with_suffix('.xlsx')
                    self.output_path.set(str(output_file))
                
                self.log_status(f"已选择PDF文件: {Path(filename).name}")
                logger.info(f"用户选择PDF文件: {filename}")
                
        except Exception as e:
            error_msg = f"选择PDF文件时出错: {str(e)}"
            self.log_status(error_msg)
            logger.error(error_msg)
            messagebox.showerror("错误", error_msg)
    
    def save_excel(self):
        """选择Excel输出文件"""
        try:
            filename = filedialog.asksaveasfilename(
                title="保存Excel文件",
                defaultextension=".xlsx",
                filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
                initialdir=config.get('gui.default_output_dir', str(Path.home()))
            )
            
            if filename:
                self.output_path.set(filename)
                self.log_status(f"输出文件设置为: {Path(filename).name}")
                logger.info(f"用户设置输出文件: {filename}")
                
        except Exception as e:
            error_msg = f"选择输出文件时出错: {str(e)}"
            self.log_status(error_msg)
            logger.error(error_msg)
            messagebox.showerror("错误", error_msg)
    
    def update_button_states(self, *args):
        """更新按钮状态"""
        has_pdf = bool(self.pdf_path.get().strip())
        has_output = bool(self.output_path.get().strip())
        can_convert = has_pdf and has_output and not self.is_converting
        
        # 更新转换按钮状态
        self.convert_button.config(state='normal' if can_convert else 'disabled')
        
        # 更新打开目录按钮状态
        output_dir_exists = False
        if has_output:
            try:
                output_dir_exists = Path(self.output_path.get()).parent.exists()
            except:
                pass
        
        self.open_dir_button.config(state='normal' if output_dir_exists else 'disabled')
    
    def update_progress_display(self, *args):
        """更新进度显示"""
        percent = self.progress_var.get()
        self.progress_label.config(text=f"{percent}%")
    
    def update_tables_display(self, *args):
        """更新表格数量显示"""
        count = self.tables_found_var.get()
        self.tables_label.config(text=str(count))
    
    def log_status(self, message: str):
        """记录状态信息到文本框"""
        timestamp = time.strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"
        
        self.status_textbox.config(state='normal')
        self.status_textbox.insert(tk.END, formatted_message)
        self.status_textbox.see(tk.END)
        self.status_textbox.config(state='disabled')
        
        # 更新状态栏
        self.status_text.set(message)
    
    def progress_callback(self, percent: int, status: str, tables_found: int):
        """进度回调函数"""
        # 使用after方法确保在主线程中更新UI
        self.root.after(0, self._update_progress_impl, percent, status, tables_found)
        
    def _update_progress_impl(self, percent: int, status: str, tables_found: int):
        """在主线程中更新进度"""
        self.progress_var.set(percent)
        self.tables_found_var.set(tables_found)
        self.log_status(status)
    
    def start_conversion(self):
        """开始转换"""
        if self.is_converting:
            return
        
        # 验证输入
        pdf_path = self.pdf_path.get().strip()
        output_path = self.output_path.get().strip()
        
        if not pdf_path:
            messagebox.showerror("错误", "请选择PDF文件")
            return
        
        if not output_path:
            messagebox.showerror("错误", "请设置输出文件路径")
            return
        
        if not os.path.exists(pdf_path):
            messagebox.showerror("错误", f"PDF文件不存在: {pdf_path}")
            return
        
        # 开始转换
        self.is_converting = True
        self.update_button_states()
        self.cancel_button.config(state='normal')
        
        # 清空状态信息
        self.status_textbox.config(state='normal')
        self.status_textbox.delete(1.0, tk.END)
        self.status_textbox.config(state='disabled')
        
        # 重置进度
        self.progress_var.set(0)
        self.tables_found_var.set(0)
        
        self.log_status("开始转换...")
        logger.info(f"开始转换: {pdf_path} -> {output_path}")
        
        # 在新线程中执行转换
        self.conversion_thread = threading.Thread(
            target=self._conversion_worker,
            args=(pdf_path, output_path),
            daemon=True
        )
        self.conversion_thread.start()
        
        # 定期检查转换状态
        self.check_conversion_thread()
    
    def _conversion_worker(self, pdf_path: str, output_path: str):
        """转换工作线程"""
        try:
            # 提取表格
            tables = pdf_processor.extract_tables(pdf_path, self.progress_callback)
            
            if not tables:
                self.root.after(0, self._conversion_finished, False, "未找到任何表格")
                return
            
            # 保存到Excel
            success = excel_writer.save_tables(tables, output_path, self.progress_callback)
            
            # 转换完成
            if success:
                self.root.after(0, self._conversion_finished, True, f"转换完成! 共处理 {len(tables)} 个表格")
            else:
                self.root.after(0, self._conversion_finished, False, "保存Excel文件失败")
                
        except Exception as e:
            error_msg = f"转换过程中出现错误: {str(e)}"
            logger.exception(error_msg)
            self.root.after(0, self._conversion_finished, False, error_msg)
    
    def _conversion_finished(self, success: bool, message: str):
        """转换完成处理"""
        self.is_converting = False
        self.update_button_states()
        self.cancel_button.config(state='disabled')
        
        if success:
            self.progress_var.set(100)
            self.log_status(message)
            messagebox.showinfo("完成", message)
            logger.info(f"转换成功: {message}")
        else:
            self.log_status(f"转换失败: {message}")
            messagebox.showerror("错误", f"转换失败: {message}")
            logger.error(f"转换失败: {message}")
    
    def check_conversion_thread(self):
        """检查转换线程状态"""
        if self.conversion_thread and self.conversion_thread.is_alive():
            # 线程仍在运行，继续检查
            self.root.after(100, self.check_conversion_thread)
        else:
            # 线程已结束
            if self.is_converting:
                # 如果标志仍为True，说明是异常结束
                self._conversion_finished(False, "转换过程异常终止")
    
    def cancel_conversion(self):
        """取消转换"""
        if not self.is_converting:
            return
        
        # 设置取消标志
        pdf_processor.cancel_processing()
        excel_writer.cancel_saving()
        
        self.log_status("正在取消转换...")
        logger.info("用户取消转换操作")
        
        # 等待一段时间让取消操作生效
        self.root.after(1000, self._check_cancellation)
    
    def _check_cancellation(self):
        """检查取消状态"""
        if self.is_converting:
            self.is_converting = False
            self.update_button_states()
            self.cancel_button.config(state='disabled')
            self.log_status("转换已取消")
            messagebox.showinfo("已取消", "转换操作已取消")
    
    def open_output_dir(self):
        """打开输出目录"""
        try:
            output_path = self.output_path.get().strip()
            if output_path:
                output_dir = Path(output_path).parent
                if output_dir.exists():
                    if sys.platform.startswith('darwin'):  # macOS
                        os.system(f'open "{output_dir}"')
                    elif sys.platform.startswith('win'):  # Windows
                        os.system(f'explorer "{output_dir}"')
                    else:  # Linux
                        os.system(f'xdg-open "{output_dir}"')
                    
                    self.log_status(f"已打开输出目录: {output_dir}")
                    logger.info(f"打开输出目录: {output_dir}")
                else:
                    messagebox.showwarning("警告", "输出目录不存在")
            
        except Exception as e:
            error_msg = f"打开输出目录失败: {str(e)}"
            self.log_status(error_msg)
            logger.error(error_msg)
            messagebox.showerror("错误", error_msg)
    
    def exit_app(self):
        """退出应用"""
        if self.is_converting:
            result = messagebox.askyesno(
                "确认",
                "转换正在进行中，确定要退出吗？\n退出后转换过程将被中断。"
            )
            if not result:
                return
            
            # 取消转换
            self.cancel_conversion()
        
        logger.info("用户退出应用")
        self.root.quit()
        self.root.destroy()
    
    def run(self):
        """运行GUI应用"""
        try:
            self.create_main_window()
            logger.info("启动Tkinter GUI界面")
            self.root.mainloop()
        except Exception as e:
            logger.exception(f"GUI运行错误: {str(e)}")
            raise


# 导出类，不创建实例 