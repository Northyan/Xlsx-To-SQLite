/*
 * Copyright 2026 Lucian Li
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;

using ClosedXML.Excel;
using System.Data.SQLite;
using System.Text.Json;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace XlsxToSQLite;

public partial class MainWindow : Window
{
    private string _selectedFile = string.Empty;
    private string _outputDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
    
    public MainWindow()
    {
        InitializeComponent();
        OutputPathTextBox.Text = _outputDirectory;
    }
    
    private void DropArea_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files.Length > 0)
            {
                string file = files[0];
                if (IsValidExcelFile(file))
                {
                    LoadFile(file);
                }
                else
                {
                    MessageBox.Show("请选择有效的 Excel 文件 (.xlsx, .xlsm, .xltx, .xltm)", "文件格式错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }
        ResetDropAreaStyle();
    }
    
    private void DropArea_DragEnter(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files.Length > 0 && IsValidExcelFile(files[0]))
            {
                DropArea.Background = new SolidColorBrush(Color.FromRgb(232, 245, 232));
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }
        else
        {
            e.Effects = DragDropEffects.None;
        }
    }
    
    private void DropArea_DragLeave(object sender, DragEventArgs e)
    {
        ResetDropAreaStyle();
    }
    
    private void DropArea_Click(object sender, RoutedEventArgs e)
    {
        var openFileDialog = new OpenFileDialog
        {
            Filter = "Excel 文件|*.xlsx;*.xlsm;*.xltx;*.xltm|所有文件|*.*",
            Title = "选择 Excel 文件"
        };
        
        if (openFileDialog.ShowDialog() == true)
        {
            LoadFile(openFileDialog.FileName);
        }
    }
    
    private void BrowseButton_Click(object sender, RoutedEventArgs e)
    {
        // 使用OpenFileDialog作为文件夹选择器
        var dialog = new OpenFileDialog
        {
            Title = "选择输出目录",
            Filter = "文件夹|*",
            ValidateNames = false,
            CheckFileExists = false,
            CheckPathExists = true,
            FileName = "选择文件夹"
        };
        
        if (dialog.ShowDialog() == true)
        {
            // 获取选择的文件夹路径
            string? directoryPath = System.IO.Path.GetDirectoryName(dialog.FileName);
            _outputDirectory = directoryPath ?? _outputDirectory;
            OutputPathTextBox.Text = _outputDirectory;
        }
    }
    
    private async void ConvertButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_selectedFile)) return;
        
        ConvertButton.IsEnabled = false;
        StatusText.Text = "正在转换...";
        
        try
        {
            string outputFileName = Path.GetFileNameWithoutExtension(_selectedFile);
            string outputFile;
            
            if (FormatComboBox.SelectedIndex == 0) // SQLite
            {
                outputFile = Path.Combine(_outputDirectory, $"{outputFileName}.db");
                await ExportToSQLite(_selectedFile, outputFile);
            }
            else // JSONL
            {
                outputFile = Path.Combine(_outputDirectory, $"{outputFileName}.jsonl");
                await ExportToJsonl(_selectedFile, outputFile);
            }
            
            StatusText.Text = $"转换完成: {outputFile}";
            MessageBox.Show($"文件已成功转换并保存到:\n{outputFile}", "转换完成", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            StatusText.Text = "转换失败";
            string errorMessage = ex.Message;
            
            // 提供更友好的错误信息
            if (errorMessage.Contains("Specified part does not exist"))
            {
                errorMessage = "Excel文件可能已损坏或格式不正确。请尝试用Excel打开并重新保存文件。";
            }
            else if (errorMessage.Contains("The process cannot access the file"))
            {
                errorMessage = "文件被其他程序占用，请关闭相关程序后重试。";
            }
            else if (errorMessage.Contains("Could not find file"))
            {
                errorMessage = "找不到指定的文件。";
            }
            
            MessageBox.Show($"转换过程中发生错误:\n{errorMessage}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            ConvertButton.IsEnabled = true;
        }
    }
    
    private void LoadFile(string filePath)
    {
        try
        {
            // 最简化的文件验证
            if (!File.Exists(filePath))
            {
                MessageBox.Show("文件不存在。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
                
            _selectedFile = filePath;
            var fileInfo = new FileInfo(filePath);
                    
            FileNameText.Text = Path.GetFileName(filePath);
            FileSizeText.Text = $"文件大小: {GetFileSizeString(fileInfo.Length)}";
            StatusText.Text = "文件已加载，准备转换";
            ConvertButton.IsEnabled = true;
                    
            //切换显示区域
            DropArea.Visibility = Visibility.Collapsed;
            FileInfoPanel.Visibility = Visibility.Visible;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"文件加载失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
            
    private void ClearButton_Click(object sender, RoutedEventArgs e)
    {
        _selectedFile = string.Empty;
        ConvertButton.IsEnabled = false;
                
        //切换显示区域
        DropArea.Visibility = Visibility.Visible;
        FileInfoPanel.Visibility = Visibility.Collapsed;
    }
    
    private static bool IsValidExcelFile(string filePath)
    {
        string? extension = Path.GetExtension(filePath)?.ToLower();
        return extension == ".xlsx" || extension == ".xlsm" || extension == ".xltx" || extension == ".xltm";
    }
    
    private void ResetDropAreaStyle()
    {
        DropArea.Background = new SolidColorBrush(Color.FromRgb(250, 250, 250));
    }
    
    private static string GetFileSizeString(long bytes)
    {
        if (bytes < 1024) return $"{bytes} B";
        if (bytes < 1024 * 1024) return $"{bytes / 1024.0:F1} KB";
        return $"{bytes / (1024.0 * 1024.0):F1} MB";
    }
    
    private static async Task ExportToSQLite(string inputPath, string outputPath)
    {
        await Task.Run(() =>
        {
            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }
            
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }
            
            using var workbook = new XLWorkbook(inputPath);
            var worksheet = workbook.Worksheets.First();
            var rows = worksheet.RowsUsed();
            
            if (!rows.Any())
            {
                throw new Exception("Excel文件中没有数据行。");
            }
            
            SQLiteConnection.CreateFile(outputPath);
            using var connection = new SQLiteConnection($"Data Source={outputPath};Version=3;");
            connection.Open();
            
            var headers = new List<string>();
            var firstRow = worksheet.FirstRow();
            foreach (var cell in firstRow.Cells())
            {
                var headerValue = cell.Value.ToString();
                var headerName = string.IsNullOrEmpty(headerValue) ? $"Column_{cell.Address.ColumnNumber}" : headerValue;
                headers.Add(headerName.Replace(" ", "_").Replace("-", "_"));
            }
            
            if (headers.Count == 0)
            {
                throw new Exception("Excel文件中没有有效的列标题。");
            }
            
            var createTableSql = $"CREATE TABLE IF NOT EXISTS data ({string.Join(", ", headers.Select(h => $"`{h}` TEXT"))})";
            using var createCmd = new SQLiteCommand(createTableSql, connection);
            createCmd.ExecuteNonQuery();
            
            using var transaction = connection.BeginTransaction();
            var insertSql = $"INSERT INTO data ({string.Join(", ", headers.Select(h => $"`{h}`"))}) VALUES ({string.Join(", ", headers.Select((_, i) => $"@param{i}"))})";
            using var insertCmd = new SQLiteCommand(insertSql, connection, transaction);
            
            foreach (var row in rows.Skip(1))
            {
                insertCmd.Parameters.Clear();
                
                for (int i = 0; i < headers.Count; i++)
                {
                    string cellValue;
                    try
                    {
                        var cell = row.Cell(i + 1); // 使用列号直接访问，确保空单元格也能被处理
                        cellValue = cell.Value.ToString();
                    }
                    catch
                    {
                        cellValue = string.Empty;
                    }
                    insertCmd.Parameters.AddWithValue($"@param{i}", cellValue);
                }
                
                insertCmd.ExecuteNonQuery();
            }
            
            transaction.Commit();
        });
    }
    
    private static async Task ExportToJsonl(string inputPath, string outputPath)
    {
        await Task.Run(() =>
        {
            using var workbook = new XLWorkbook(inputPath);
            var worksheet = workbook.Worksheets.First();
            var rows = worksheet.RowsUsed();
            
            if (!rows.Any())
            {
                throw new Exception("Excel文件中没有数据行。");
            }
            
            using var writer = new StreamWriter(outputPath, false, System.Text.Encoding.UTF8);
            var headers = new List<string>();
            var firstRow = worksheet.FirstRow();
            
            foreach (var cell in firstRow.Cells())
            {
                headers.Add(cell.Value.ToString() ?? string.Empty);
            }
            
            if (headers.Count == 0)
            {
                throw new Exception("Excel文件中没有有效的列标题。");
            }
            
            var jsonOptions = new JsonSerializerOptions
            {
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
                WriteIndented = false
            };
            
            foreach (var row in rows.Skip(1))
            {
                var record = new Dictionary<string, object>();
                
                for (int i = 0; i < headers.Count; i++)
                {
                    string cellValue;
                    try
                    {
                        var cell = row.Cell(i + 1); // 使用列号直接访问，确保空单元格也能被处理
                        cellValue = cell.Value.ToString();
                    }
                    catch
                    {
                        cellValue = string.Empty;
                    }
                    record[headers[i]] = cellValue;
                }
                
                writer.WriteLine(JsonSerializer.Serialize(record, jsonOptions));
            }
        });
    }
    
    private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ClickCount == 2)
        {
            if (WindowState == WindowState.Normal)
                WindowState = WindowState.Maximized;
            else
                WindowState = WindowState.Normal;
        }
        else
        {
            DragMove();
        }
    }
    
    private void MinimizeButton_Click(object sender, RoutedEventArgs e)
    {
        WindowState = WindowState.Minimized;
    }
    
    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
}