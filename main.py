#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
PDF表格转Excel工具 - 主入口文件
"""

import sys
import os
import argparse
from pathlib import Path

# 添加src目录到Python路径
sys.path.insert(0, str(Path(__file__).parent / "src"))

from src.utils.logger import logger
from src.utils.system_check import system_checker
from src.utils.config import config


def check_environment():
    """检查运行环境"""
    logger.info("开始环境检查...")
    
    # 执行完整的系统检查
    check_results = system_checker.full_system_check()
    
    if not check_results['summary']['all_checks_passed']:
        logger.warning("环境检查发现问题:")
        
        if not check_results['summary']['java_available']:
            logger.error("Java环境未安装或不可用")
            print("\n❌ 错误: 未检测到Java环境!")
            print("此应用需要Java才能运行。")
            print("请从 https://www.java.com 下载安装Java。\n")
            return False
        
        if not check_results['summary']['dependencies_satisfied']:
            logger.error("Python依赖包不完整")
            print("\n❌ 错误: Python依赖包不完整!")
            print("请运行以下命令安装依赖:")
            print("pip install -r requirements.txt\n")
            return False
        
        if not check_results['summary']['sufficient_memory']:
            logger.warning("系统可用内存不足")
            print("\n⚠️ 警告: 系统可用内存不足，可能影响处理大型PDF文件")
        
        if not check_results['summary']['sufficient_disk']:
            logger.warning("磁盘可用空间不足")
            print("\n⚠️ 警告: 磁盘可用空间不足，可能影响输出文件保存")
    
    logger.info("环境检查完成")
    return True


def run_gui(gui_type: str = "tkinter"):
    """
    运行GUI界面
    
    Args:
        gui_type: GUI类型，支持 "tkinter" 或 "pysimplegui"
    """
    try:
        if gui_type.lower() == "tkinter":
            from src.gui.tkinter_gui import TkinterGUI
            app = TkinterGUI()
            app.run()
        elif gui_type.lower() == "pysimplegui":
            from src.gui.pysimplegui_gui import PySimpleGUI_App
            app = PySimpleGUI_App()
            app.run()
        else:
            logger.error(f"不支持的GUI类型: {gui_type}")
            print(f"❌ 错误: 不支持的GUI类型 '{gui_type}'")
            print("支持的类型: tkinter, pysimplegui")
            return False
        
        return True
        
    except ImportError as e:
        logger.error(f"GUI模块导入失败: {str(e)}")
        print(f"❌ 错误: GUI模块导入失败")
        print(f"详细信息: {str(e)}")
        print("\n请确保已安装所需的GUI依赖包")
        return False
    except Exception as e:
        logger.exception(f"GUI运行失败: {str(e)}")
        print(f"❌ 错误: GUI运行失败")
        print(f"详细信息: {str(e)}")
        return False


def run_cli(pdf_path: str, output_path: str):
    """
    运行命令行模式
    
    Args:
        pdf_path: PDF文件路径
        output_path: 输出Excel文件路径
    """
    try:
        from src.core.pdf_processor import pdf_processor
        from src.core.excel_writer import excel_writer
        
        # 验证输入文件
        if not os.path.exists(pdf_path):
            logger.error(f"PDF文件不存在: {pdf_path}")
            print(f"❌ 错误: PDF文件不存在 '{pdf_path}'")
            return False
        
        # 创建输出目录
        output_dir = Path(output_path).parent
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # 进度回调函数
        def progress_callback(percent, status, tables_found):
            print(f"\r进度: {percent}% - {status}", end="", flush=True)
        
        print(f"开始处理PDF文件: {pdf_path}")
        
        # 提取表格
        tables = pdf_processor.extract_tables(pdf_path, progress_callback)
        
        if not tables:
            print("\n⚠️ 未找到任何表格")
            return False
        
        print(f"\n找到 {len(tables)} 个表格，开始保存...")
        
        # 保存到Excel
        success = excel_writer.save_tables(tables, output_path, progress_callback)
        
        if success:
            print(f"\n✅ 转换完成! 输出文件: {output_path}")
            return True
        else:
            print("\n❌ 保存失败")
            return False
        
    except Exception as e:
        logger.exception(f"命令行模式运行失败: {str(e)}")
        print(f"\n❌ 错误: {str(e)}")
        return False


def main():
    """主函数"""
    # 解析命令行参数
    parser = argparse.ArgumentParser(
        description="PDF表格转Excel工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  # 启动GUI界面 (默认tkinter)
  python main.py
  
  # 启动PySimpleGUI界面
  python main.py --gui pysimplegui
  
  # 命令行模式
  python main.py --cli input.pdf output.xlsx
        """
    )
    
    parser.add_argument(
        "--gui", 
        choices=["tkinter", "pysimplegui"], 
        default="tkinter",
        help="GUI类型 (默认: tkinter)"
    )
    
    parser.add_argument(
        "--cli", 
        nargs=2, 
        metavar=("PDF_FILE", "OUTPUT_FILE"),
        help="命令行模式: 指定输入PDF文件和输出Excel文件路径"
    )
    
    parser.add_argument(
        "--no-check", 
        action="store_true",
        help="跳过环境检查"
    )
    
    parser.add_argument(
        "--version", 
        action="version",
        version=f"PDF表格转Excel工具 v{config.get('app.version', '2.0.0')}"
    )
    
    args = parser.parse_args()
    
    # 显示欢迎信息
    print("=" * 60)
    print(f"  {config.get('app.name', 'PDF表格转Excel工具')} v{config.get('app.version', '2.0.0')}")
    print("=" * 60)
    
    # 环境检查
    if not args.no_check:
        if not check_environment():
            sys.exit(1)
    
    # 根据参数选择运行模式
    if args.cli:
        # 命令行模式
        pdf_path, output_path = args.cli
        success = run_cli(pdf_path, output_path)
        sys.exit(0 if success else 1)
    else:
        # GUI模式
        success = run_gui(args.gui)
        sys.exit(0 if success else 1)


if __name__ == "__main__":
    main() 