#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
内存管理模块
监控和优化内存使用
"""

import gc
import os
import sys
import ctypes
import platform
from typing import Optional

from ..utils.logger import logger


class MemoryManager:
    """内存使用监控和管理"""
    
    @staticmethod
    def get_memory_usage() -> float:
        """
        获取当前进程内存使用量（MB）
        
        Returns:
            float: 内存使用量（MB）
        """
        try:
            import psutil
            process = psutil.Process(os.getpid())
            memory_info = process.memory_info()
            usage_mb = memory_info.rss / 1024 / 1024
            return usage_mb
        except ImportError:
            logger.warning("psutil未安装，无法获取精确内存使用量")
            return 0.0
        except Exception as e:
            logger.error(f"获取内存使用量失败: {str(e)}")
            return 0.0
    
    @staticmethod
    def get_system_memory() -> tuple[float, float, float]:
        """
        获取系统内存信息
        
        Returns:
            tuple[float, float, float]: (总内存MB, 可用内存MB, 使用百分比)
        """
        try:
            import psutil
            memory = psutil.virtual_memory()
            total_mb = memory.total / 1024 / 1024
            available_mb = memory.available / 1024 / 1024
            percent = memory.percent
            return total_mb, available_mb, percent
        except ImportError:
            logger.warning("psutil未安装，无法获取系统内存信息")
            return 0.0, 0.0, 0.0
        except Exception as e:
            logger.error(f"获取系统内存信息失败: {str(e)}")
            return 0.0, 0.0, 0.0
    
    @staticmethod
    def free_memory() -> int:
        """
        强制垃圾回收，释放内存
        
        Returns:
            int: 回收的对象数量
        """
        try:
            # 记录释放前的内存使用
            before_mb = MemoryManager.get_memory_usage()
            
            # 调用多次垃圾回收
            collected_0 = gc.collect(0)  # 收集第0代（最年轻的对象）
            collected_1 = gc.collect(1)  # 收集第1代
            collected_2 = gc.collect(2)  # 收集第2代（最老的对象）
            
            total_collected = collected_0 + collected_1 + collected_2
            
            # 尝试释放未使用的内存返回给OS
            if platform.system() == 'Linux' and hasattr(os, 'malloc_trim'):
                os.malloc_trim(0)
            elif platform.system() == 'Darwin':  # macOS
                try:
                    libc = ctypes.CDLL('libc.dylib')
                    if hasattr(libc, 'malloc_zone_pressure_relief'):
                        libc.malloc_zone_pressure_relief(None, 100)
                except Exception:
                    pass  # 忽略释放失败
            
            # 记录释放后的内存使用
            after_mb = MemoryManager.get_memory_usage()
            freed_mb = before_mb - after_mb
            
            if freed_mb > 0:
                logger.debug(f"内存释放成功: 释放 {freed_mb:.2f}MB, 回收对象 {total_collected} 个")
            
            return total_collected
            
        except Exception as e:
            logger.error(f"内存释放失败: {str(e)}")
            return 0
    
    @staticmethod
    def print_memory_status():
        """输出当前内存状态"""
        try:
            process_mb = MemoryManager.get_memory_usage()
            total_mb, available_mb, percent = MemoryManager.get_system_memory()
            
            logger.info(f"内存状态 - 进程使用: {process_mb:.2f}MB, "
                       f"系统: {available_mb:.2f}/{total_mb:.2f}MB ({100-percent:.1f}%可用)")
        except Exception as e:
            logger.error(f"获取内存状态失败: {str(e)}")
    
    @staticmethod
    def check_and_free_memory(threshold: float = 1000) -> bool:
        """
        检查内存使用，如果超过阈值则尝试释放
        
        Args:
            threshold: 内存使用阈值，单位MB
            
        Returns:
            bool: 是否进行了内存释放
        """
        try:
            mem_usage = MemoryManager.get_memory_usage()
            if mem_usage > threshold:
                logger.info(f"内存使用超过阈值 ({mem_usage:.2f}MB > {threshold}MB)，尝试释放内存...")
                
                collected = MemoryManager.free_memory()
                new_usage = MemoryManager.get_memory_usage()
                
                logger.info(f"内存释放完成: 使用量 {mem_usage:.2f}MB -> {new_usage:.2f}MB, "
                           f"释放 {mem_usage - new_usage:.2f}MB")
                return True
            return False
        except Exception as e:
            logger.error(f"内存检查和释放失败: {str(e)}")
            return False
    
    @staticmethod
    def optimize_dataframe(df) -> object:
        """
        优化DataFrame内存使用
        
        Args:
            df: pandas DataFrame
            
        Returns:
            object: 优化后的DataFrame
        """
        try:
            import pandas as pd
            
            # 记录优化前的内存使用
            before_memory = df.memory_usage(deep=True).sum() / 1024 / 1024
            
            # 对象类型列转换为类别类型
            for col in df.select_dtypes(include=['object']).columns:
                unique_ratio = df[col].nunique() / len(df[col])
                if unique_ratio < 0.5:  # 如果唯一值少于50%
                    df[col] = df[col].astype('category')
            
            # 将浮点数列转换为最合适的数值类型
            for col in df.select_dtypes(include=['float']).columns:
                df[col] = pd.to_numeric(df[col], downcast='float')
            
            # 将整数列转换为最合适的整数类型
            for col in df.select_dtypes(include=['int']).columns:
                df[col] = pd.to_numeric(df[col], downcast='integer')
            
            # 记录优化后的内存使用
            after_memory = df.memory_usage(deep=True).sum() / 1024 / 1024
            saved_memory = before_memory - after_memory
            
            if saved_memory > 0:
                logger.debug(f"DataFrame内存优化: 节省 {saved_memory:.2f}MB "
                           f"({before_memory:.2f}MB -> {after_memory:.2f}MB)")
            
            return df
            
        except Exception as e:
            logger.warning(f"DataFrame内存优化失败: {str(e)}")
            return df
    
    @classmethod
    def get_optimal_batch_size(cls, total_items: int, base_batch_size: int = 10) -> int:
        """
        根据系统内存情况计算最优批处理大小
        
        Args:
            total_items: 总处理项目数
            base_batch_size: 基础批处理大小
            
        Returns:
            int: 优化后的批处理大小
        """
        try:
            total_mb, available_mb, percent = cls.get_system_memory()
            
            # 如果可用内存充足（>2GB），可以增加批处理大小
            if available_mb > 2048:
                multiplier = min(4, int(available_mb / 512))  # 最多4倍
                optimal_size = base_batch_size * multiplier
            # 如果内存紧张（<512MB），减少批处理大小
            elif available_mb < 512:
                optimal_size = max(2, base_batch_size // 2)
            else:
                optimal_size = base_batch_size
            
            # 确保批处理大小不超过总项目数
            optimal_size = min(optimal_size, total_items)
            
            logger.debug(f"批处理大小优化: {base_batch_size} -> {optimal_size} "
                        f"(可用内存: {available_mb:.0f}MB)")
            
            return optimal_size
            
        except Exception as e:
            logger.warning(f"批处理大小优化失败，使用默认值: {str(e)}")
            return base_batch_size


# 创建全局内存管理器实例
memory_manager = MemoryManager() 