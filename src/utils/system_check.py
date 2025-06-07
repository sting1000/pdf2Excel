#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
系统检查模块
检查运行环境和依赖
"""

import os
import sys
import subprocess
import platform
from typing import Tuple, Optional

from .logger import logger


class SystemChecker:
    """系统环境检查器"""
    
    @staticmethod
    def check_java() -> Tuple[bool, str]:
        """
        检查Java是否已安装
        
        Returns:
            Tuple[bool, str]: (是否安装, 版本信息或错误信息)
        """
        try:
            # 尝试运行java -version命令
            result = subprocess.run(
                ['java', '-version'], 
                stdout=subprocess.PIPE, 
                stderr=subprocess.PIPE,
                text=True,
                timeout=10
            )
            
            if result.returncode == 0:
                # Java版本信息通常在stderr中
                version_info = result.stderr.strip()
                logger.info(f"Java检查成功: {version_info.split()[0] if version_info else 'Unknown'}")
                return True, version_info
            else:
                error_msg = f"Java运行失败，返回码: {result.returncode}"
                logger.warning(error_msg)
                return False, error_msg
                
        except subprocess.TimeoutExpired:
            error_msg = "Java检查超时"
            logger.warning(error_msg)
            return False, error_msg
        except FileNotFoundError:
            error_msg = "未找到Java命令，可能未安装Java"
            logger.warning(error_msg)
            return False, error_msg
        except Exception as e:
            error_msg = f"Java检查时出现异常: {str(e)}"
            logger.error(error_msg)
            return False, error_msg
    
    @staticmethod
    def get_system_info() -> dict:
        """
        获取系统信息
        
        Returns:
            dict: 系统信息字典
        """
        return {
            'platform': platform.platform(),
            'system': platform.system(),
            'machine': platform.machine(),
            'processor': platform.processor(),
            'python_version': platform.python_version(),
            'python_executable': sys.executable,
        }
    
    @staticmethod
    def check_memory() -> Tuple[float, float]:
        """
        检查系统内存
        
        Returns:
            Tuple[float, float]: (总内存GB, 可用内存GB)
        """
        try:
            import psutil
            memory = psutil.virtual_memory()
            total_gb = memory.total / (1024**3)
            available_gb = memory.available / (1024**3)
            logger.info(f"系统内存: 总计 {total_gb:.2f}GB, 可用 {available_gb:.2f}GB")
            return total_gb, available_gb
        except ImportError:
            logger.warning("psutil未安装，无法检查内存信息")
            return 0.0, 0.0
        except Exception as e:
            logger.error(f"内存检查失败: {str(e)}")
            return 0.0, 0.0
    
    @staticmethod
    def check_disk_space(path: str = ".") -> float:
        """
        检查磁盘空间
        
        Args:
            path: 检查路径，默认当前目录
            
        Returns:
            float: 可用空间GB
        """
        try:
            if platform.system() == 'Windows':
                import ctypes
                free_bytes = ctypes.c_ulonglong(0)
                ctypes.windll.kernel32.GetDiskFreeSpaceExW(
                    ctypes.c_wchar_p(path), 
                    ctypes.pointer(free_bytes), 
                    None, 
                    None
                )
                free_gb = free_bytes.value / (1024**3)
            else:
                statvfs = os.statvfs(path)
                free_gb = (statvfs.f_frsize * statvfs.f_bavail) / (1024**3)
            
            logger.info(f"磁盘可用空间: {free_gb:.2f}GB")
            return free_gb
        except Exception as e:
            logger.error(f"磁盘空间检查失败: {str(e)}")
            return 0.0
    
    @staticmethod
    def check_dependencies() -> dict:
        """
        检查Python依赖包
        
        Returns:
            dict: 依赖包检查结果
        """
        required_packages = [
            'pandas',
            'tabula',
            'PyPDF2', 
            'openpyxl',
            'psutil'
        ]
        
        results = {}
        for package in required_packages:
            try:
                if package == 'tabula':
                    import tabula
                    results[package] = {'installed': True, 'version': getattr(tabula, '__version__', 'Unknown')}
                else:
                    __import__(package)
                    module = sys.modules[package]
                    version = getattr(module, '__version__', 'Unknown')
                    results[package] = {'installed': True, 'version': version}
                logger.info(f"依赖检查成功: {package} v{results[package]['version']}")
            except ImportError:
                results[package] = {'installed': False, 'version': None}
                logger.warning(f"依赖缺失: {package}")
            except Exception as e:
                results[package] = {'installed': False, 'version': None, 'error': str(e)}
                logger.error(f"依赖检查失败: {package} - {str(e)}")
        
        return results
    
    @staticmethod
    def setup_java_path():
        """设置Java路径到环境变量"""
        try:
            # 根据不同操作系统设置路径分隔符
            separator = ";" if platform.system() == 'Windows' else ":"
            java_path = os.path.join(os.path.dirname(sys.executable), "java")
            
            if os.path.exists(java_path):
                current_path = os.environ.get("PATH", "")
                if java_path not in current_path:
                    os.environ["PATH"] = current_path + separator + java_path
                    logger.info(f"已添加Java路径到环境变量: {java_path}")
            else:
                logger.debug(f"Java路径不存在: {java_path}")
        except Exception as e:
            logger.error(f"设置Java路径失败: {str(e)}")
    
    @classmethod
    def full_system_check(cls) -> dict:
        """
        执行完整的系统检查
        
        Returns:
            dict: 完整的系统检查结果
        """
        logger.info("开始系统环境检查...")
        
        # 设置Java路径
        cls.setup_java_path()
        
        results = {
            'system_info': cls.get_system_info(),
            'java_check': cls.check_java(),
            'memory_info': cls.check_memory(),
            'disk_space': cls.check_disk_space(),
            'dependencies': cls.check_dependencies(),
        }
        
        # 总结检查结果
        java_ok = results['java_check'][0]
        deps_ok = all(dep['installed'] for dep in results['dependencies'].values())
        memory_sufficient = results['memory_info'][1] > 0.5  # 至少500MB可用内存
        disk_sufficient = results['disk_space'] > 1.0  # 至少1GB可用空间
        
        results['summary'] = {
            'all_checks_passed': java_ok and deps_ok and memory_sufficient and disk_sufficient,
            'java_available': java_ok,
            'dependencies_satisfied': deps_ok,
            'sufficient_memory': memory_sufficient,
            'sufficient_disk': disk_sufficient,
        }
        
        if results['summary']['all_checks_passed']:
            logger.info("✅ 所有系统检查通过")
        else:
            logger.warning("⚠️ 系统检查发现问题，请查看详细结果")
        
        return results


# 创建全局系统检查器实例
system_checker = SystemChecker() 