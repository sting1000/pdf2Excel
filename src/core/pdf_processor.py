#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
PDF处理核心模块
负责PDF文件的读取、分析和表格提取
"""

import os
import sys
import time
import math
import threading
from typing import List, Tuple, Callable, Dict, Any, Optional
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, ProcessPoolExecutor

import PyPDF2
import tabula
import pandas as pd

from ..utils.logger import logger
from ..utils.config import config
from .memory_manager import memory_manager


def suppress_stdout_stderr(func):
    """装饰器：用于完全抑制函数执行过程中的stdout和stderr输出"""
    def wrapper(*args, **kwargs):
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
def extract_tables_silent(pdf_path: str, page_range: str) -> List:
    """
    静默提取表格，不输出任何信息
    
    Args:
        pdf_path: PDF文件路径
        page_range: 页面范围，如 "1-10"
        
    Returns:
        List: 提取的表格列表
    """
    try:
        logger.debug(f"开始提取PDF表格: {pdf_path}, 页面: {page_range}")
        
        # 使用文件对象而不是路径，减少文件句柄占用
        with open(pdf_path, 'rb') as pdf_file:
            tables = tabula.read_pdf(
                pdf_file, 
                pages=page_range, 
                multiple_tables=True, 
                silent=True
            )
        
        # 过滤空表格
        tables = [table for table in tables if not table.empty]
        
        logger.debug(f"页面 {page_range} 提取到 {len(tables)} 个有效表格")
        return tables
    except Exception as e:
        logger.error(f"表格提取失败 - 页面 {page_range}: {str(e)}")
        return []


class PDFProcessor:
    """PDF处理器类"""
    
    def __init__(self):
        """初始化PDF处理器"""
        self.cancel_flag = {"cancel": False}
        self.processing_thread = None
        
    def get_pdf_page_count(self, pdf_path: str) -> int:
        """
        获取PDF总页数
        
        Args:
            pdf_path: PDF文件路径
            
        Returns:
            int: PDF总页数
        """
        try:
            with open(pdf_path, 'rb') as pdf_file:
                pdf_reader = PyPDF2.PdfReader(pdf_file)
                page_count = len(pdf_reader.pages)
                logger.info(f"PDF文件 {Path(pdf_path).name} 共有 {page_count} 页")
                return page_count
        except Exception as e:
            logger.error(f"读取PDF页数失败: {str(e)}")
            raise
    
    def extract_tables(
        self, 
        pdf_path: str, 
        progress_callback: Callable[[int, str, int], None]
    ) -> List[pd.DataFrame]:
        """
        从PDF中提取所有表格
        
        Args:
            pdf_path: PDF文件路径
            progress_callback: 进度回调函数 (percent, status_text, tables_found)
            
        Returns:
            List[pd.DataFrame]: 提取的表格列表
        """
        try:
            # 重置取消标志
            self.cancel_flag["cancel"] = False
            
            # 初始化进度
            progress_callback(0, "正在分析PDF文件...", 0)
            
            # 获取PDF总页数
            total_pages = self.get_pdf_page_count(pdf_path)
            progress_callback(1, f"PDF共有 {total_pages} 页，开始提取表格...", 0)
            
            # 获取批处理大小
            base_batch_size = config.get('processing.batch_size', 10)
            batch_size = memory_manager.get_optimal_batch_size(total_pages, base_batch_size)
            total_batches = math.ceil(total_pages / batch_size)
            
            logger.info(f"使用批处理模式: 批大小={batch_size}, 总批次={total_batches}")
            
            all_tables = []
            start_time = time.time()
            total_tables_found = 0
            
            # 处理每一批次
            for batch in range(total_batches):
                # 检查取消标志
                if self.cancel_flag["cancel"]:
                    progress_callback(0, "操作已取消", 0)
                    return []
                    
                start_page = batch * batch_size + 1
                end_page = min((batch + 1) * batch_size, total_pages)
                
                # 构建页范围字符串
                page_range = f"{start_page}-{end_page}"
                
                try:
                    # 更新状态
                    progress_callback(
                        int(batch * 80 / total_batches), 
                        f"正在处理页面 {start_page}-{end_page} (共{total_pages}页)...", 
                        total_tables_found
                    )
                    
                    # 提取表格
                    tables = extract_tables_silent(pdf_path, page_range)
                    
                    if tables:
                        all_tables.extend(tables)
                        total_tables_found += len(tables)
                    
                    # 计算进度百分比 (总体完成的80%用于提取)
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
                    
                    # 定期释放内存
                    memory_manager.check_and_free_memory(
                        threshold=config.get('processing.memory_threshold', 1000)
                    )
                    
                except Exception as e:
                    error_msg = f"处理页 {page_range} 时出错: {str(e)}"
                    logger.error(error_msg)
                    progress_callback(percent, error_msg, total_tables_found)
                    # 继续处理下一批次
            
            # 返回结果
            total_time = time.time() - start_time
            
            if self.cancel_flag["cancel"]:
                progress_callback(0, "操作已取消", 0)
                return []
            elif total_tables_found > 0:
                progress_callback(
                    80, 
                    f"✅ 表格提取完成! 共找到 {total_tables_found} 个表格，用时: {total_time:.1f}秒", 
                    total_tables_found
                )
                return all_tables
            else:
                progress_callback(80, "⚠️ 未找到任何表格", 0)
                return []
                
        except Exception as e:
            logger.exception(f"PDF表格提取过程中出错: {str(e)}")
            progress_callback(0, f"提取过程中出错: {str(e)}", 0)
            return []
    
    def cancel_processing(self):
        """取消处理过程"""
        logger.info("用户请求取消PDF处理")
        self.cancel_flag["cancel"] = True
    
    def is_processing(self) -> bool:
        """检查是否正在处理"""
        return (self.processing_thread is not None and 
                self.processing_thread.is_alive())


# 创建全局PDF处理器实例
pdf_processor = PDFProcessor() 