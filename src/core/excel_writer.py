#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Excel写入模块
负责将表格数据写入Excel文件
"""

import os
import time
from typing import List, Callable, Optional
from pathlib import Path

import pandas as pd

from ..utils.logger import logger
from ..utils.config import config
from .memory_manager import memory_manager


class ExcelWriter:
    """Excel写入器类"""
    
    def __init__(self):
        """初始化Excel写入器"""
        self.cancel_flag = {"cancel": False}
    
    def save_tables(
        self,
        tables: List[pd.DataFrame],
        output_path: str,
        progress_callback: Callable[[int, str, int], None]
    ) -> bool:
        """
        保存表格到Excel文件
        
        Args:
            tables: 表格列表
            output_path: 输出文件路径
            progress_callback: 进度回调函数 (percent, status_text, tables_found)
            
        Returns:
            bool: 是否保存成功
        """
        if not tables:
            logger.warning("没有表格需要保存")
            progress_callback(100, "⚠️ 没有表格需要保存", 0)
            return False
        
        try:
            # 重置取消标志
            self.cancel_flag["cancel"] = False
            
            # 开始保存进度提示
            total_tables = len(tables)
            progress_callback(80, f"正在保存 {total_tables} 个表格到Excel...", total_tables)
            
            start_time = time.time()
            
            # 获取配置
            excel_engine = config.get('output.excel_engine', 'openpyxl')
            sheet_prefix = config.get('output.sheet_name_prefix', 'Table_')
            max_sheet_name_length = config.get('output.max_sheet_name_length', 31)
            
            # 确保输出目录存在
            output_dir = Path(output_path).parent
            output_dir.mkdir(parents=True, exist_ok=True)
            
            logger.info(f"开始保存 {total_tables} 个表格到 {output_path}")
            
            # 使用 ExcelWriter 一次性写入所有表格
            with pd.ExcelWriter(output_path, engine=excel_engine) as writer:
                for i, df in enumerate(tables):
                    # 检查取消标志
                    if self.cancel_flag["cancel"]:
                        progress_callback(0, "操作已取消", 0)
                        return False
                    
                    # 计算保存进度 (从80%到100%)
                    save_percent = 80 + int((i + 1) * 20 / total_tables)
                    
                    # 生成工作表名称
                    sheet_name = f"{sheet_prefix}{i+1}"
                    
                    # 表格名称长度限制
                    if len(sheet_name) > max_sheet_name_length:
                        sheet_name = f"T{i+1}"
                    
                    # 检查空表格
                    if df.empty:
                        logger.warning(f"跳过空表格 {i+1}")
                        continue
                    
                    try:
                        # 优化DataFrame内存使用
                        df = memory_manager.optimize_dataframe(df)
                        
                        # 保存表格
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        # 更新进度
                        progress_callback(
                            save_percent, 
                            f"保存表格: {i+1}/{total_tables} ({sheet_name})", 
                            total_tables
                        )
                        
                        logger.debug(f"成功保存表格 {i+1} 到工作表 {sheet_name}")
                        
                        # 定期检查内存使用并释放
                        if (i + 1) % 5 == 0:
                            memory_manager.check_and_free_memory(
                                threshold=config.get('processing.memory_threshold', 1000)
                            )
                        
                    except Exception as e:
                        error_msg = f"保存表格 {i+1} 时出错: {str(e)}"
                        logger.error(error_msg)
                        progress_callback(
                            save_percent, 
                            error_msg, 
                            total_tables
                        )
                        # 继续处理下一个表格
                        continue
            
            # 保存完成
            total_time = time.time() - start_time
            file_size = self._get_file_size(output_path)
            
            if self.cancel_flag["cancel"]:
                progress_callback(0, "操作已取消", 0)
                return False
            else:
                success_msg = (f"✅ 保存完成! 已保存 {total_tables} 个表格到 {Path(output_path).name}\n"
                              f"文件大小: {file_size}, 用时: {total_time:.1f}秒")
                progress_callback(100, success_msg, total_tables)
                logger.info(f"Excel文件保存成功: {output_path}")
                return True
                
        except Exception as e:
            error_msg = f"保存Excel文件时出错: {str(e)}"
            logger.exception(error_msg)
            progress_callback(0, error_msg, 0)
            return False
    
    def save_tables_chunked(
        self,
        tables: List[pd.DataFrame],
        output_path: str,
        progress_callback: Callable[[int, str, int], None],
        chunk_size: int = 50
    ) -> bool:
        """
        分块保存表格到Excel文件（适用于大量表格）
        
        Args:
            tables: 表格列表
            output_path: 输出文件路径
            progress_callback: 进度回调函数
            chunk_size: 每块的表格数量
            
        Returns:
            bool: 是否保存成功
        """
        if not tables:
            logger.warning("没有表格需要保存")
            progress_callback(100, "⚠️ 没有表格需要保存", 0)
            return False
        
        try:
            total_tables = len(tables)
            
            # 如果表格数量不多，使用常规保存方式
            if total_tables <= chunk_size:
                return self.save_tables(tables, output_path, progress_callback)
            
            logger.info(f"使用分块保存模式: {total_tables} 个表格，每块 {chunk_size} 个")
            
            # 重置取消标志
            self.cancel_flag["cancel"] = False
            
            start_time = time.time()
            
            # 获取配置
            excel_engine = config.get('output.excel_engine', 'openpyxl')
            sheet_prefix = config.get('output.sheet_name_prefix', 'Table_')
            max_sheet_name_length = config.get('output.max_sheet_name_length', 31)
            
            # 确保输出目录存在
            output_dir = Path(output_path).parent
            output_dir.mkdir(parents=True, exist_ok=True)
            
            # 分块保存
            with pd.ExcelWriter(output_path, engine=excel_engine) as writer:
                for i, df in enumerate(tables):
                    # 检查取消标志
                    if self.cancel_flag["cancel"]:
                        progress_callback(0, "操作已取消", 0)
                        return False
                    
                    # 计算保存进度
                    save_percent = 80 + int((i + 1) * 20 / total_tables)
                    
                    # 生成工作表名称
                    sheet_name = f"{sheet_prefix}{i+1}"
                    if len(sheet_name) > max_sheet_name_length:
                        sheet_name = f"T{i+1}"
                    
                    if df.empty:
                        continue
                    
                    try:
                        # 优化DataFrame
                        df = memory_manager.optimize_dataframe(df)
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        # 更新进度
                        progress_callback(
                            save_percent, 
                            f"分块保存: {i+1}/{total_tables}", 
                            total_tables
                        )
                        
                        # 每处理一定数量的表格后释放内存
                        if (i + 1) % chunk_size == 0:
                            memory_manager.free_memory()
                            progress_callback(
                                save_percent, 
                                f"已保存 {i+1}/{total_tables} 个表格，正在释放内存...", 
                                total_tables
                            )
                        
                    except Exception as e:
                        logger.error(f"分块保存表格 {i+1} 时出错: {str(e)}")
                        continue
            
            # 保存完成
            total_time = time.time() - start_time
            file_size = self._get_file_size(output_path)
            
            success_msg = (f"✅ 分块保存完成! 已保存 {total_tables} 个表格\n"
                          f"文件: {Path(output_path).name}, 大小: {file_size}, 用时: {total_time:.1f}秒")
            progress_callback(100, success_msg, total_tables)
            logger.info(f"Excel文件分块保存成功: {output_path}")
            return True
            
        except Exception as e:
            error_msg = f"分块保存Excel文件时出错: {str(e)}"
            logger.exception(error_msg)
            progress_callback(0, error_msg, 0)
            return False
    
    def _get_file_size(self, file_path: str) -> str:
        """
        获取文件大小的可读格式
        
        Args:
            file_path: 文件路径
            
        Returns:
            str: 文件大小字符串
        """
        try:
            size = os.path.getsize(file_path)
            if size < 1024:
                return f"{size} B"
            elif size < 1024 * 1024:
                return f"{size / 1024:.1f} KB"
            else:
                return f"{size / (1024 * 1024):.1f} MB"
        except Exception as e:
            logger.warning(f"获取文件大小失败: {str(e)}")
            return "未知"
    
    def cancel_saving(self):
        """取消保存过程"""
        logger.info("用户请求取消Excel保存")
        self.cancel_flag["cancel"] = True


# 创建全局Excel写入器实例
excel_writer = ExcelWriter() 