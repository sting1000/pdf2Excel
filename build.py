#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
PDF表格转Excel工具打包脚本
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

def check_requirements():
    """检查所需依赖是否安装"""
    try:
        import PyInstaller
    except ImportError:
        print("正在安装PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    # 检查其他依赖
    requirements = ["tabula-py", "pandas", "openpyxl", "PyPDF2"]
    for req in requirements:
        try:
            __import__(req.replace("-", "_"))
        except ImportError:
            print(f"正在安装{req}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", req])

def clean_build_dirs():
    """清理构建目录"""
    dirs_to_clean = ["build", "dist"]
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            print(f"清理目录: {dir_name}...")
            shutil.rmtree(dir_name)

def build_app():
    """打包应用程序"""
    print("=" * 60)
    print("开始打包PDF表格转Excel工具...")
    print("=" * 60)
    
    # 使用优化版本的tkinter
    source_file = "pdf_table_converter_tkinter.py"
    
    if not os.path.exists(source_file):
        print(f"错误: 未找到源文件 {source_file}")
        return False
    
    # 确定平台
    platform = sys.platform
    if platform == "win32":
        platform_name = "Windows"
    elif platform == "darwin":
        platform_name = "Mac"
    else:
        platform_name = "Linux"
    
    print(f"检测到平台: {platform_name}")
    print(f"使用源文件: {source_file}")
    
    # 构建PyInstaller命令
    cmd = [
        "pyinstaller",
        "--onefile",
        "--windowed",
        "--name", f"PDF表格转Excel工具",
        "--add-data", f"requirements.txt{';' if platform == 'win32' else ':'}.",
        "--add-data", f"README.md{';' if platform == 'win32' else ':'}.",
        "--hidden-import", "concurrent.futures",
        source_file
    ]
    
    # 添加图标（如果存在）
    icon_file = "icon.ico" if platform == "win32" else "icon.icns"
    if os.path.exists(icon_file):
        cmd.extend(["--icon", icon_file])
    
    # 执行构建
    print("\n开始构建...")
    try:
        subprocess.check_call(cmd)
        print("\n构建成功!")
        
        # 复制其他文件到dist目录
        if os.path.exists("dist"):
            # 创建输出zip文件
            output_zip = f"PDF表格转Excel工具_{platform_name}.zip"
            print(f"\n创建发布包: {output_zip}")
            
            # 进入dist目录并创建zip
            os.chdir("dist")
            try:
                if platform == "win32":
                    zip_cmd = ["powershell", "Compress-Archive", "-Path", "PDF表格转Excel工具.exe", "-DestinationPath", output_zip]
                    subprocess.check_call(zip_cmd)
                else:
                    if platform == "darwin":
                        target = "PDF表格转Excel工具.app"
                    else:
                        target = "PDF表格转Excel工具"
                    zip_cmd = ["zip", "-r", output_zip, target]
                    subprocess.check_call(zip_cmd)
                
                os.chdir("..")
                
                print(f"\n发布包创建成功: dist/{output_zip}")
                exe_extension = ".exe" if platform == "win32" else ""
                print(f"可执行文件路径: dist/PDF表格转Excel工具{exe_extension}")
                
                return True
            except subprocess.CalledProcessError as e:
                os.chdir("..")
                print(f"\n创建zip文件时出错: {e}")
                return False
            except Exception as e:
                os.chdir("..")
                print(f"\n创建zip文件时出错: {e}")
                return False
    except subprocess.CalledProcessError as e:
        print(f"\n构建过程中出错: {e}")
        return False

def create_release_notes():
    """创建发布说明"""
    release_notes = """# PDF表格转Excel工具 - 发布说明

## 版本 1.1.0

### 新功能和改进
- 全面优化界面显示，解决文字显示问题
- 增加并行处理能力，显著提高大型PDF的处理速度
- 根据PDF大小动态调整批处理策略
- 改进内存管理，优化大型PDF文件的处理
- 提供更精确的进度和时间估计
- 更友好的用户界面和交互体验

### 系统要求
- Windows 10/11 或 macOS 10.13+
- Java运行环境(JRE) - 用于表格提取功能

### 安装说明
1. 下载适合您系统的安装包
2. 解压缩文件
3. 运行应用程序
4. 如果提示缺少Java，请从[Java官网](https://www.java.com/)下载安装
"""
    
    with open("RELEASE_NOTES.md", "w", encoding="utf-8") as f:
        f.write(release_notes)
    
    print("已创建发布说明文件: RELEASE_NOTES.md")
    return True

def main():
    # 检查依赖
    print("检查依赖...")
    check_requirements()
    
    # 清理旧的构建目录
    clean_build_dirs()
    
    # 创建发布说明
    create_release_notes()
    
    # 构建应用
    if build_app():
        print("\n构建完成。你可以在dist目录找到可执行文件和发布包。")
        print("\n发布步骤:")
        print("1. 登录GitHub仓库")
        print("2. 点击'Releases'标签")
        print("3. 点击'Draft a new release'")
        print("4. 创建一个新标签(例如v1.1.0)")
        print("5. 上传生成的zip文件")
        print("6. 复制RELEASE_NOTES.md的内容到发布说明中")
        print("7. 点击'Publish release'")
    else:
        print("\n构建失败，请检查错误信息并修复问题。")

if __name__ == "__main__":
    main() 