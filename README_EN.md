# Xlsx To SQLite

This is a WPF desktop application for converting Excel files to SQLite databases or JSONL files.

## Features

- Supports dragging and dropping Excel files into the application interface
- Supports converting Excel worksheets to SQLite database tables or JSONL format
- Supports multiple Excel formats: .xlsx, .xlsm, .xltx, .xltm
- Supports custom output directory
- Simple and easy-to-use graphical interface

## Supported File Formats

- **.xlsx**: Excel 2007+ Workbook
- **.xlsm**: Macro-enabled Excel Workbook
- **.xltx**: Excel Template
- **.xltm**: Macro-enabled Excel Template

**Note**: Workbooks with macros (.xlsm) and template files (.xltx/.xltm) only support data extraction, macros are not preserved.

## Usage

1. Run the application
2. Drag and drop an Excel file into the application window, or click the select file button to open one
3. Choose the output format (SQLite database or JSONL file)
4. Choose the output directory (defaults to desktop)
5. Click the "Convert" button to start conversion
6. After conversion, the output file will be saved in the specified directory

## System Requirements

- Windows Operating System
- .NET Framework

## License

This project is licensed under the Apache License 2.0.
