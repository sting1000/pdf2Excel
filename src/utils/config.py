#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
配置管理模块
处理应用程序的配置信息
"""

import os
import yaml
from pathlib import Path
from typing import Dict, Any


class Config:
    """配置管理类"""
    
    def __init__(self, config_file: str = None):
        """
        初始化配置管理器
        
        Args:
            config_file: 配置文件路径，如果为None则使用默认配置
        """
        self._config_data = self._load_default_config()
        
        if config_file and os.path.exists(config_file):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    file_config = yaml.safe_load(f)
                    if file_config:
                        self._merge_config(file_config)
            except Exception as e:
                print(f"配置文件加载失败: {e}")
    
    def _load_default_config(self) -> Dict[str, Any]:
        """加载默认配置"""
        return {
            'app': {
                'name': 'PDF表格转Excel工具',
                'version': '2.0.0',
                'window_size': (800, 600),
                'theme': 'default'
            },
            'processing': {
                'batch_size': 10,  # 每批处理的页数
                'max_workers': 4,  # 并行处理的最大工作线程数
                'memory_threshold': 1000,  # 内存使用阈值(MB)
                'timeout': 300,  # 处理超时时间(秒)
            },
            'output': {
                'excel_engine': 'openpyxl',
                'sheet_name_prefix': 'Table_',
                'max_sheet_name_length': 31,
            },
            'ui': {
                'update_interval': 100,  # UI更新间隔(毫秒)
                'progress_precision': 1,  # 进度显示精度
            }
        }
    
    def _merge_config(self, file_config: Dict[str, Any]):
        """合并配置"""
        def merge_dict(base: dict, override: dict):
            for key, value in override.items():
                if key in base and isinstance(base[key], dict) and isinstance(value, dict):
                    merge_dict(base[key], value)
                else:
                    base[key] = value
        
        merge_dict(self._config_data, file_config)
    
    def get(self, key: str, default=None):
        """
        获取配置值，支持点号分隔的路径
        
        Args:
            key: 配置键，支持 'app.name' 格式
            default: 默认值
        """
        keys = key.split('.')
        value = self._config_data
        
        for k in keys:
            if isinstance(value, dict) and k in value:
                value = value[k]
            else:
                return default
        
        return value
    
    def set(self, key: str, value: Any):
        """
        设置配置值
        
        Args:
            key: 配置键，支持 'app.name' 格式
            value: 配置值
        """
        keys = key.split('.')
        target = self._config_data
        
        for k in keys[:-1]:
            if k not in target:
                target[k] = {}
            target = target[k]
        
        target[keys[-1]] = value
    
    def save(self, config_file: str):
        """
        保存配置到文件
        
        Args:
            config_file: 配置文件路径
        """
        os.makedirs(os.path.dirname(config_file), exist_ok=True)
        with open(config_file, 'w', encoding='utf-8') as f:
            yaml.dump(self._config_data, f, default_flow_style=False, 
                     allow_unicode=True, encoding='utf-8')
    
    @property
    def config_data(self) -> Dict[str, Any]:
        """获取完整配置数据"""
        return self._config_data.copy()


# 全局配置实例
config = Config()

# 尝试加载项目配置文件
project_config_file = Path(__file__).parent.parent.parent / "config" / "app_config.yaml"
if project_config_file.exists():
    config = Config(str(project_config_file)) 