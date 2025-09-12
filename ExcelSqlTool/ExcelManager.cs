using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using ExcelSqlTool.Models;

namespace ExcelSqlTool
{
    /// <summary>
    /// Excel操作管理器
    /// </summary>
    public class ExcelManager
    {
        private readonly string _directoryPath;
        private readonly Dictionary<string, ExcelFile> _excelFiles;

        public ExcelManager(string directoryPath)
        {
            _directoryPath = directoryPath;
            _excelFiles = new Dictionary<string, ExcelFile>();
            LoadAllExcelFiles();
        }

        /// <summary>
        /// 加载目录下所有Excel文件
        /// </summary>
        private void LoadAllExcelFiles()
        {
            if (!Directory.Exists(_directoryPath))
            {
                throw new DirectoryNotFoundException($"目录 {_directoryPath} 不存在");
            }

            var excelFiles = Directory.GetFiles(_directoryPath, "*.xlsx");
            foreach (var filePath in excelFiles)
            {
                LoadExcelFile(filePath);
            }
        }

        /// <summary>
        /// 加载单个Excel文件
        /// </summary>
        /// <param name="filePath">文件路径</param>
        private void LoadExcelFile(string filePath)
        {
            try
            {
                using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    var workbook = new XSSFWorkbook(fileStream);
                    var fileName = Path.GetFileName(filePath);
                    var excelFile = new ExcelFile
                    {
                        FileName = fileName,
                        Path = filePath,
                        LastModified = File.GetLastWriteTime(filePath)
                    };

                    // 遍历所有工作表
                    for (int i = 0; i < workbook.NumberOfSheets; i++)
                    {
                        var sheet = workbook.GetSheetAt(i);
                        var worksheet = new Worksheet
                        {
                            Name = sheet.SheetName,
                            FilePath = filePath
                        };

                        // 解析表头和数据
                        ParseSheet(sheet, worksheet);
                        excelFile.Worksheets[sheet.SheetName] = worksheet;
                    }

                    _excelFiles[fileName] = excelFile;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"加载Excel文件 {filePath} 失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 解析工作表
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="worksheet">工作表模型</param>
        private void ParseSheet(ISheet sheet, Worksheet worksheet)
        {
            if (sheet.LastRowNum < 3)
            {
                // 空表或数据不足
                return;
            }

            // 解析表头（前3行）
            var headerRow = sheet.GetRow(0); // 第1行为列名
            if (headerRow == null) return;

            // 创建列定义
            for (int i = 0; i < headerRow.LastCellNum; i++)
            {
                var cell = headerRow.GetCell(i);
                if (cell != null)
                {
                    var columnName = cell.ToString();
                    if (!string.IsNullOrEmpty(columnName))
                    {
                        var column = new Column
                        {
                            Name = columnName,
                            Index = i,
                            DataType = InferColumnType(sheet, i) // 推断数据类型
                        };
                        worksheet.Headers.Add(column);
                    }
                }
            }

            // 解析数据行（从第4行开始）
            for (int i = 3; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row != null)
                {
                    var dataRow = new Dictionary<string, object>();
                    foreach (var column in worksheet.Headers)
                    {
                        var cell = row.GetCell(column.Index);
                        dataRow[column.Name] = GetCellValue(cell);
                    }
                    worksheet.DataRows.Add(dataRow);
                }
            }
        }

        /// <summary>
        /// 推断列的数据类型
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="columnIndex">列索引</param>
        /// <returns>数据类型</returns>
        private string InferColumnType(ISheet sheet, int columnIndex)
        {
            // 从第4行开始检查数据类型
            for (int i = 3; i <= Math.Min(sheet.LastRowNum, 10); i++)
            {
                var row = sheet.GetRow(i);
                if (row != null)
                {
                    var cell = row.GetCell(columnIndex);
                    if (cell != null)
                    {
                        var cellValue = GetCellValue(cell);
                        if (cellValue != null)
                        {
                            // 尝试解析为不同数据类型
                            if (int.TryParse(cellValue.ToString(), out _))
                                return "INT";
                            if (double.TryParse(cellValue.ToString(), out _))
                                return "DOUBLE";
                            if (DateTime.TryParse(cellValue.ToString(), out _))
                                return "DATE";
                            
                            return "VARCHAR";
                        }
                    }
                }
            }
            
            return "VARCHAR"; // 默认字符串类型
        }

        /// <summary>
        /// 获取单元格值
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <returns>单元格值</returns>
        private object GetCellValue(ICell cell)
        {
            if (cell == null) return null;

            switch (cell.CellType)
            {
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Numeric:
                    // 检查是否为日期
                    if (DateUtil.IsCellDateFormatted(cell))
                        return cell.DateCellValue;
                    return cell.NumericCellValue;
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.Formula:
                    try
                    {
                        return cell.NumericCellValue;
                    }
                    catch
                    {
                        return cell.StringCellValue;
                    }
                default:
                    return cell.ToString();
            }
        }

        /// <summary>
        /// 获取所有表名
        /// </summary>
        /// <returns>表名列表</returns>
        public List<string> GetTableNames()
        {
            var tableNames = new List<string>();
            foreach (var excelFile in _excelFiles.Values)
            {
                foreach (var worksheetName in excelFile.Worksheets.Keys)
                {
                    if (!tableNames.Contains(worksheetName))
                    {
                        tableNames.Add(worksheetName);
                    }
                }
            }
            return tableNames;
        }

        /// <summary>
        /// 获取建表语句
        /// </summary>
        /// <param name="tableName">表名（应为工作表名称）</param>
        /// <returns>建表语句</returns>
        public string GetCreateTableStatement(string tableName)
        {
            foreach (var excelFile in _excelFiles.Values)
            {
                if (excelFile.Worksheets.ContainsKey(tableName))
                {
                    var worksheet = excelFile.Worksheets[tableName];
                    var sql = $"CREATE TABLE {tableName} (\n";
                    
                    for (int i = 0; i < worksheet.Headers.Count; i++)
                    {
                        var column = worksheet.Headers[i];
                        sql += $"    {column.Name} {column.DataType}";
                        if (i < worksheet.Headers.Count - 1)
                            sql += ",";
                        sql += "\n";
                    }
                    
                    sql += ");";
                    return sql;
                }
            }
            
            // 增强错误提示，提供更友好的指导
            var availableTables = GetTableNames();
            var errorMessage = $"表 '{tableName}' 不存在。\n";
            errorMessage += "注意：表名应为Excel文件中的工作表名称，而不是Excel文件名。\n";
            errorMessage += "可用的表名包括：\n";
            foreach (var table in availableTables)
            {
                errorMessage += $"  - {table}\n";
            }
            
            throw new Exception(errorMessage);
        }

        /// <summary>
        /// 执行SELECT查询
        /// </summary>
        /// <param name="tableName">表名（应为工作表名称）</param>
        /// <param name="columns">要查询的列</param>
        /// <param name="whereClause">WHERE条件</param>
        /// <returns>查询结果</returns>
        public List<Dictionary<string, object>> ExecuteSelect(string tableName, List<string> columns, string whereClause = null)
        {
            foreach (var excelFile in _excelFiles.Values)
            {
                if (excelFile.Worksheets.ContainsKey(tableName))
                {
                    var worksheet = excelFile.Worksheets[tableName];
                    var result = new List<Dictionary<string, object>>();

                    // 确定要返回的列
                    List<Column> selectedColumns;
                    if (columns == null || columns.Count == 0 || (columns.Count == 1 && columns[0] == "*"))
                    {
                        selectedColumns = worksheet.Headers;
                    }
                    else
                    {
                        selectedColumns = worksheet.Headers.Where(h => columns.Contains(h.Name)).ToList();
                    }

                    // 应用WHERE条件过滤数据
                    var filteredRows = ApplyWhereClause(worksheet.DataRows, whereClause);

                    // 构建结果
                    foreach (var row in filteredRows)
                    {
                        var resultRow = new Dictionary<string, object>();
                        foreach (var column in selectedColumns)
                        {
                            row.TryGetValue(column.Name, out var value);
                            resultRow[column.Name] = value;
                        }
                        result.Add(resultRow);
                    }

                    return result;
                }
            }

            // 增强错误提示，提供更友好的指导
            var availableTables = GetTableNames();
            var errorMessage = $"表 '{tableName}' 不存在。\n";
            errorMessage += "注意：表名应为Excel文件中的工作表名称，而不是Excel文件名。\n";
            errorMessage += "可用的表名包括：\n";
            foreach (var table in availableTables)
            {
                errorMessage += $"  - {table}\n";
            }
            
            throw new Exception(errorMessage);
        }

        /// <summary>
        /// 应用WHERE条件过滤数据
        /// </summary>
        /// <param name="dataRows">数据行</param>
        /// <param name="whereClause">WHERE条件</param>
        /// <returns>过滤后的数据行</returns>
        private List<Dictionary<string, object>> ApplyWhereClause(List<Dictionary<string, object>> dataRows, string whereClause)
        {
            if (string.IsNullOrEmpty(whereClause))
                return dataRows;

            // 简单的WHERE条件解析（实际项目中需要更复杂的SQL解析器）
            // 这里只处理简单的等于条件，如：ID = 1
            var filteredRows = new List<Dictionary<string, object>>();
            
            // 移除WHERE关键字
            var condition = whereClause.Trim();
            if (condition.ToUpper().StartsWith("WHERE "))
            {
                condition = condition.Substring(6).Trim();
            }

            // 解析简单条件
            if (condition.Contains("="))
            {
                var parts = condition.Split('=');
                if (parts.Length == 2)
                {
                    var columnName = parts[0].Trim().Trim('\"', '\'');
                    var value = parts[1].Trim().Trim('\"', '\'');

                    foreach (var row in dataRows)
                    {
                        if (row.ContainsKey(columnName))
                        {
                            var cellValue = row[columnName]?.ToString();
                            if (cellValue == value)
                            {
                                filteredRows.Add(row);
                            }
                        }
                    }
                    
                    return filteredRows;
                }
            }

            // 如果无法解析条件，返回所有行
            return dataRows;
        }

        /// <summary>
        /// 刷新文件缓存
        /// </summary>
        public void Refresh()
        {
            _excelFiles.Clear();
            LoadAllExcelFiles();
        }
    }
}