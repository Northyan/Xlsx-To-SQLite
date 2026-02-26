# Xlsx To SQLite

![App Icon](Icons/App.ico)

## 中文

这是一款操作简单的WPF桌面应用程序，可以一键将Excel文件转换为SQLite数据库或JSONL文件。

### 功能特性

- 支持拖拽Excel文件到应用界面
- 支持将Excel工作表转换为SQLite数据库表或JSONL格式
- 支持多种Excel格式：.xlsx, .xlsm, .xltx, .xltm
- 支持自定义输出目录
- 简单易用的图形界面

### 支持的文件格式

- **.xlsx**: Excel 2007+ 工作簿
- **.xlsm**: 启用宏的 Excel 工作簿
- **.xltx**: Excel 模板
- **.xltm**: 启用宏的 Excel 模板

**注意**: 包含宏的工作簿（.xlsm）和模板文件（.xltx/.xltm）仅支持提取数据，不保留宏代码。

### 使用方法

1. 运行应用程序
2. 将Excel文件拖拽到应用窗口，或点击选择文件按钮打开
3. 选择输出格式（SQLite 数据库或 JSONL 文件）
4. 选择输出目录（默认为桌面）
5. 点击"转换"按钮开始转换
6. 转换完成后，输出文件将保存在指定目录中

### 系统要求

- Windows操作系统
- .NET Framework

### 许可证

本项目采用Apache License 2.0许可证。

## English

This is a simple-to-use WPF desktop application that can convert Excel files to SQLite databases or JSONL files with one click.

### Features

- Supports dragging and dropping Excel files into the application interface
- Supports converting Excel worksheets to SQLite database tables or JSONL format
- Supports multiple Excel formats: .xlsx, .xlsm, .xltx, .xltm
- Supports custom output directory
- Simple and easy-to-use graphical interface

### Supported File Formats

- **.xlsx**: Excel 2007+ Workbook
- **.xlsm**: Macro-enabled Excel Workbook
- **.xltx**: Excel Template
- **.xltm**: Macro-enabled Excel Template

**Note**: Workbooks with macros (.xlsm) and template files (.xltx/.xltm) only support data extraction, macros are not preserved.

### Usage

1. Run the application
2. Drag and drop an Excel file into the application window, or click the select file button to open one
3. Choose the output format (SQLite database or JSONL file)
4. Choose the output directory (defaults to desktop)
5. Click the "Convert" button to start conversion
6. After conversion, the output file will be saved in the specified directory

### System Requirements

- Windows Operating System
- .NET Framework

### License

This project is licensed under the Apache License 2.0.
