#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
PDF表格转Excel工具 - 优化版快速启动脚本
自动安装依赖并启动应用程序
"""

import os
import sys
import subprocess
import importlib
import platform

def check_and_install_dependencies():
    """检查并安装所需依赖"""
    dependencies = [
        ("tkinter", "python-tk" if platform.system() == "Linux" else "tkinter"),
        ("pandas", "pandas"),
        ("tabula", "tabula-py"),
        ("PyPDF2", "PyPDF2"),
        ("openpyxl", "openpyxl"),
        ("psutil", "psutil")
    ]
    
    missing_deps = []
    
    print("正在检查依赖...")
    for module_name, package_name in dependencies:
        try:
            if module_name == "tkinter":
                # 特殊处理tkinter
                import tkinter
            else:
                importlib.import_module(module_name)
            print(f"✓ {module_name} 已安装")
        except ImportError:
            print(f"✗ {module_name} 未安装")
            missing_deps.append(package_name)
    
    if missing_deps:
        print("\n需要安装以下依赖:")
        for dep in missing_deps:
            print(f"  - {dep}")
        
        user_input = input("\n是否自动安装这些依赖? (y/n): ")
        if user_input.lower() == 'y':
            print("\n正在安装依赖...")
            for dep in missing_deps:
                print(f"安装 {dep}...")
                try:
                    subprocess.check_call([sys.executable, "-m", "pip", "install", dep])
                except subprocess.CalledProcessError as e:
                    print(f"安装 {dep} 失败: {e}")
                    if dep == "python-tk" and platform.system() == "Linux":
                        print("请使用系统包管理器安装tkinter，例如:")
                        print("  Ubuntu/Debian: sudo apt-get install python3-tk")
                        print("  Fedora: sudo dnf install python3-tkinter")
                        print("  Arch Linux: sudo pacman -S tk")
                    return False
            print("\n所有依赖已安装完成!")
        else:
            print("\n请手动安装缺少的依赖后再运行此脚本。")
            return False
    
    return True

def check_java():
    """检查是否安装了Java"""
    try:
        java_version = subprocess.check_output(["java", "-version"], stderr=subprocess.STDOUT, text=True)
        print("\n✓ Java已安装:")
        for line in java_version.split('\n')[:3]:
            if line.strip():
                print(f"  {line.strip()}")
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("\n✗ 未检测到Java，这个应用需要Java来提取PDF表格。")
        print("请从 https://www.java.com 下载安装Java。")
        
        user_input = input("\n是否仍然继续启动程序? (y/n): ")
        return user_input.lower() == 'y'

def run_application():
    """运行优化版应用程序"""
    script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pdf_table_converter_tkinter.py")
    
    if not os.path.exists(script_path):
        print(f"错误: 未找到应用程序脚本: {script_path}")
        return False
    
    print("\n正在启动PDF表格转Excel工具 (优化版)...")
    try:
        subprocess.check_call([sys.executable, script_path])
        return True
    except subprocess.CalledProcessError as e:
        print(f"启动应用程序失败: {e}")
        return False

def main():
    """主函数"""
    # 显示欢迎信息
    print("=" * 60)
    print("  PDF表格转Excel工具 - 优化版快速启动")
    print("=" * 60)
    print("\n这个脚本将检查所需依赖并启动优化版应用程序。\n")
    
    # 检查依赖
    if not check_and_install_dependencies():
        print("\n启动失败: 依赖问题未解决。")
        input("按Enter键退出...")
        return
    
    # 检查Java
    if not check_java():
        print("\n启动取消。")
        input("按Enter键退出...")
        return
    
    # 运行应用程序
    if not run_application():
        print("\n应用程序启动失败。")
        input("按Enter键退出...")
        return
    
    print("\n应用程序已关闭。")

if __name__ == "__main__":
    main() 