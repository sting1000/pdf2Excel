# PDF表格转Excel工具 v2.0

一款现代化、模块化的PDF表格提取转换工具，支持从PDF文件中自动提取表格并转换为Excel格式。

## ✨ 功能特点

- 🔍 **智能表格识别** - 自动识别和提取PDF中的所有表格
- 📊 **Excel输出** - 每个表格保存为Excel文件中的单独工作表
- ⚡ **高性能处理** - 支持批处理和并行处理大型PDF文件
- 💾 **内存优化** - 智能内存管理，支持处理大文件
- 🎯 **实时进度** - 显示处理进度和预计剩余时间
- 🖥️ **多界面支持** - 支持Tkinter和PySimpleGUI两种界面
- 📱 **命令行模式** - 支持批量处理和自动化
- 🛡️ **错误处理** - 完善的错误处理和日志记录
- ⚙️ **配置管理** - 灵活的配置系统

## 🏗️ 项目结构

```
pdf2Excel/
├── src/                        # 源代码
│   ├── core/                   # 核心业务逻辑
│   │   ├── pdf_processor.py    # PDF处理模块
│   │   ├── excel_writer.py     # Excel写入模块
│   │   └── memory_manager.py   # 内存管理模块
│   ├── gui/                    # 用户界面
│   │   ├── base_gui.py         # GUI基类
│   │   ├── tkinter_gui.py      # Tkinter界面
│   │   └── pysimplegui_gui.py  # PySimpleGUI界面
│   └── utils/                  # 工具模块
│       ├── config.py           # 配置管理
│       ├── logger.py           # 日志管理
│       └── system_check.py     # 系统检查
├── config/                     # 配置文件
│   └── app_config.yaml         # 应用配置
├── scripts/                    # 脚本目录
│   ├── build/                  # 构建脚本
│   └── run/                    # 运行脚本
├── tests/                      # 测试文件
├── docs/                       # 文档
├── example/                    # 示例文件
├── main.py                     # 主入口文件
├── requirements.txt            # 生产依赖
├── requirements-dev.txt        # 开发依赖
└── README.md                   # 项目说明
```

## 🚀 快速开始

### 环境要求

- **Python**: 3.8或更高版本
- **Java**: JRE 8或更高版本 (用于tabula-py库)
- **操作系统**: Windows, macOS, Linux

### 安装步骤

1. **克隆项目**
   ```bash
   git clone <repository-url>
   cd pdf2Excel
   ```

2. **安装Java运行环境** (如果尚未安装)
   - 从 [Java官网](https://www.java.com/) 下载并安装

3. **创建虚拟环境** (推荐)
   ```bash
   python -m venv venv
   
   # Windows
   venv\Scripts\activate
   
   # macOS/Linux
   source venv/bin/activate
   ```

4. **安装依赖**
   ```bash
   # 生产环境
   pip install -r requirements.txt
   
   # 开发环境
   pip install -r requirements-dev.txt
   ```

### 使用方法

#### GUI模式 (推荐)

```bash
# 启动默认界面 (Tkinter)
python main.py

# 启动PySimpleGUI界面
python main.py --gui pysimplegui
```

#### 命令行模式

```bash
# 转换单个文件
python main.py --cli input.pdf output.xlsx

# 跳过环境检查
python main.py --cli input.pdf output.xlsx --no-check
```

#### 使用优化启动脚本

```bash
# 自动检查依赖并启动
python scripts/run/run_optimized.py
```

## ⚙️ 配置说明

应用配置文件位于 `config/app_config.yaml`，可以调整以下参数：

```yaml
processing:
  batch_size: 10              # 每批处理的页数
  max_workers: 4              # 并行处理线程数
  memory_threshold: 1000      # 内存使用阈值(MB)
  timeout: 300                # 处理超时时间(秒)

output:
  excel_engine: "openpyxl"    # Excel引擎
  sheet_name_prefix: "Table_" # 工作表名称前缀
  max_sheet_name_length: 31   # 工作表名称最大长度
```

## 📊 性能优化

- **批处理**: 自动根据系统内存调整批处理大小
- **并行处理**: 支持多线程并行提取表格
- **内存管理**: 智能内存监控和释放
- **DataFrame优化**: 自动优化数据类型减少内存使用

## 🔧 开发指南

### 代码结构

- **模块化设计**: 业务逻辑、界面、工具分离
- **配置驱动**: 通过配置文件控制行为
- **日志记录**: 完整的日志系统便于调试
- **错误处理**: 优雅的错误处理和用户提示

### 添加新功能

1. 在相应模块中添加功能代码
2. 更新配置文件 (如需要)
3. 添加测试用例
4. 更新文档

### 构建可执行文件

```bash
# 使用PyInstaller构建
python scripts/build/build.py

# 或直接使用spec文件
pyinstaller scripts/build/PDF表格转Excel工具.spec
```

## 🐛 故障排除

### 常见问题

1. **Java环境问题**
   ```
   错误: 未检测到Java环境
   解决: 安装Java JRE，确保java命令可用
   ```

2. **内存不足**
   ```
   错误: 处理大文件时内存不足
   解决: 调整配置中的batch_size和memory_threshold
   ```

3. **依赖包问题**
   ```
   错误: 模块导入失败
   解决: pip install -r requirements.txt
   ```

### 日志文件

- 应用日志: `logs/pdf2excel_YYYYMMDD.log`
- 错误日志: `logs/pdf2excel_error_YYYYMMDD.log`

## 📝 更新日志

### v2.0.0 (当前版本)
- ✨ 完全重构，模块化架构
- ⚡ 性能优化，支持并行处理
- 💾 智能内存管理
- 🎯 改进的用户界面
- 📱 命令行模式支持
- ⚙️ 配置系统
- 🛡️ 完善的错误处理和日志

### v1.x
- 基础PDF表格提取功能
- 简单GUI界面

## 📄 许可证

本项目使用 MIT 许可证 - 详情请参见 [LICENSE](LICENSE) 文件

## 🤝 贡献

欢迎提交Issue和Pull Request来改进这个项目！

## 📞 支持

如果您遇到问题或有建议，请：

1. 查看 [故障排除](#故障排除) 部分
2. 检查日志文件获取详细错误信息
3. 提交Issue描述问题

---

**PDF表格转Excel工具** - 让PDF表格提取变得简单高效！ 