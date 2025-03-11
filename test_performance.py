#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
PDF表格转Excel工具 - 性能测试脚本
用于比较优化前后处理速度的变化
"""

import os
import sys
import time
import PyPDF2
import pandas as pd
from pathlib import Path
import concurrent.futures
import multiprocessing

# 导入原始和优化后的处理函数
from pdf_table_converter_tkinter import extract_tables_silent, process_batch

def test_performance(pdf_path, method="original", batch_size=10, workers=1):
    """
    测试PDF处理性能
    
    参数:
    - pdf_path: PDF文件路径
    - method: 处理方法 ("original" 或 "optimized")
    - batch_size: 批处理大小
    - workers: 并行处理的工作线程数
    
    返回:
    - 处理时间（秒）
    - 找到的表格数量
    """
    print(f"正在测试 {method} 方法, 批处理大小: {batch_size}, 工作线程: {workers}")
    
    # 获取PDF总页数
    with open(pdf_path, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        total_pages = len(pdf_reader.pages)
    
    print(f"PDF共有 {total_pages} 页")
    start_time = time.time()
    tables_found = 0
    
    if method == "original":
        # 原始方法：顺序处理每个批次
        for batch in range(0, total_pages, batch_size):
            start_page = batch + 1
            end_page = min(batch + batch_size, total_pages)
            page_range = f"{start_page}-{end_page}"
            
            try:
                tables = extract_tables_silent(pdf_path, page_range)
                if tables:
                    tables_found += len(tables)
                print(f"处理页 {page_range} 完成，当前共找到 {tables_found} 个表格")
            except Exception as e:
                print(f"处理页 {page_range} 出错: {str(e)}")
    
    else:  # optimized
        # 优化方法：并行处理批次
        batches = []
        for batch in range(0, total_pages, batch_size):
            start_page = batch + 1
            end_page = min(batch + batch_size, total_pages)
            batches.append((pdf_path, start_page, end_page))
        
        with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as executor:
            futures = [executor.submit(process_batch, batch) for batch in batches]
            
            for future in concurrent.futures.as_completed(futures):
                try:
                    tables, count = future.result()
                    tables_found += count
                    batch_index = futures.index(future)
                    start_page = batches[batch_index][1]
                    end_page = batches[batch_index][2]
                    print(f"处理页 {start_page}-{end_page} 完成，当前共找到 {tables_found} 个表格")
                except Exception as e:
                    print(f"处理批次时出错: {str(e)}")
    
    end_time = time.time()
    processing_time = end_time - start_time
    
    print(f"处理完成！共找到 {tables_found} 个表格，用时: {processing_time:.2f} 秒")
    return processing_time, tables_found

def run_performance_tests(pdf_path):
    """运行不同配置的性能测试"""
    if not os.path.exists(pdf_path):
        print(f"错误: PDF文件不存在: {pdf_path}")
        return
    
    results = []
    
    # 测试原始方法
    time_original, tables_original = test_performance(
        pdf_path, 
        method="original", 
        batch_size=10, 
        workers=1
    )
    results.append(("原始方法", 10, 1, time_original, tables_original))
    
    # 测试不同批处理大小的优化方法
    batch_sizes = [20, 50, 100]
    cpu_count = multiprocessing.cpu_count()
    workers = max(1, min(cpu_count - 1, 4))
    
    for batch_size in batch_sizes:
        time_opt, tables_opt = test_performance(
            pdf_path, 
            method="optimized", 
            batch_size=batch_size, 
            workers=workers
        )
        results.append((f"优化方法 (批次={batch_size})", batch_size, workers, time_opt, tables_opt))
    
    # 显示性能对比
    print("\n=== 性能测试结果 ===")
    print(f"{'方法':<25} {'批次大小':<10} {'线程数':<10} {'处理时间(秒)':<15} {'加速比':<10} {'表格数':<10}")
    print("-" * 80)
    
    baseline_time = results[0][3]
    for method, batch_size, workers, proc_time, tables in results:
        speedup = baseline_time / proc_time if proc_time > 0 else 0
        print(f"{method:<25} {batch_size:<10} {workers:<10} {proc_time:<15.2f} {speedup:<10.2f} {tables:<10}")

if __name__ == "__main__":
    # 获取PDF文件路径
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    else:
        print("请提供PDF文件路径作为参数")
        pdf_path = input("PDF文件路径: ").strip()
    
    run_performance_tests(pdf_path) 