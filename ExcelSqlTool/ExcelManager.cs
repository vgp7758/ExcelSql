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
        private string _directoryPath;
        private readonly Dictionary<string, ExcelFile> _excelFiles;
        private SqliteManager _sqlite; // 新增：SQLite 运行时数据库

        public ExcelManager(string directoryPath)
        {
            // 只在非MCP模式下输出调试信息
            WriteDebug("ExcelManager初始化，目录路径: " + directoryPath);

            _directoryPath = directoryPath;
            _excelFiles = new Dictionary<string, ExcelFile>();
            if(string.IsNullOrEmpty(_directoryPath))
            {
                return;
            }
            LoadAllExcelFiles();

            WriteDebug("ExcelManager加载完成");
        }

        /// <summary>
        /// 检查是否为MCP模式
        /// </summary>
        /// <returns>是否为MCP模式</returns>
        private bool IsMcpMode()
        {
            // 检查环境变量或命令行参数来判断是否为MCP模式
            return Environment.GetCommandLineArgs().Any(a => string.Equals(a, "--mcp", StringComparison.OrdinalIgnoreCase) || 
                                                             string.Equals(a, "-mcp", StringComparison.OrdinalIgnoreCase) || 
                                                             string.Equals(a, "mcp", StringComparison.OrdinalIgnoreCase));
        }

        private void WriteDebug(string message)
        {
            if (!IsMcpMode())
            {
                Console.WriteLine($"DEBUG: {message}");
            }
        }

        /// <summary>
        /// 获取当前目录路径
        /// </summary>
        public string DirectoryPath => _directoryPath;

        public bool IsDirectoryExists => !string.IsNullOrEmpty(_directoryPath) && Directory.Exists(_directoryPath);

        /// <summary>
        /// 更新目录路径并重新加载文件
        /// </summary>
        /// <param name="newDirectoryPath">新的目录路径</param>
        public void UpdateDirectoryPath(string newDirectoryPath)
        {
            if (string.IsNullOrEmpty(newDirectoryPath))
            {
                throw new ArgumentException("目录路径不能为空");
            }

            if (!Directory.Exists(newDirectoryPath))
            {
                throw new DirectoryNotFoundException($"目录 {newDirectoryPath} 不存在");
            }

            _directoryPath = newDirectoryPath;
            _excelFiles.Clear();
            LoadAllExcelFiles();
        }

        /// <summary>
        /// 加载目录下所有Excel文件
        /// </summary>
        private void LoadAllExcelFiles()
        {
            WriteDebug("开始加载Excel文件 " + _directoryPath);

            if (!Directory.Exists(_directoryPath))
            {
                WriteDebug("目录不存在，跳过加载Excel文件");
                return;
            }

            var excelFiles = Directory.GetFiles(_directoryPath, "*.xlsx");

            WriteDebug($"找到 {excelFiles.Length} 个Excel文件");

            foreach (var filePath in excelFiles)
            {
                WriteDebug("加载文件: " + filePath);
                LoadExcelFile(filePath);
            }

            // Excel载入完成后，构建/重建SQLite运行库
            BuildOrRebuildSqlite();
        }

        /// <summary>
        /// 加载单个Excel文件
        /// </summary>
        /// <param name="filePath">文件路径</param>
        private void LoadExcelFile(string filePath)
        {
            try
            {
                WriteDebug("开始解析Excel文件: " + filePath);

                using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    var workbook = new XSSFWorkbook(fileStream);
                    var fileName = Path.GetFileName(filePath);

                    WriteDebug($"文件 {fileName} 包含 {workbook.NumberOfSheets} 个工作表");

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
                        WriteDebug($"解析工作表: {sheet.SheetName}");

                        if (sheet.SheetName == "Struct") continue; // 跳过Struct表

                        var worksheet = new Worksheet
                        {
                            Name = sheet.SheetName,
                            FilePath = filePath
                        };

                        // 解析表头和数据
                        ParseSheet(sheet, worksheet);

                        WriteDebug($"工作表 {sheet.SheetName} 解析完成，包含 {worksheet.DataRows.Count} 行数据");
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
        /// 构建或重建SQLite运行数据库
        /// </summary>
        private void BuildOrRebuildSqlite()
        {
            try
            {
                _sqlite?.Dispose();
                // 可将DB放在临时目录，每次重建，简单稳妥
                _sqlite = new SqliteManager();
                _sqlite.RebuildFromExcel(_excelFiles);
            }
            catch (Exception ex)
            {
                // MCP模式下禁止写stdout
                try { Console.Error.WriteLine($"WARN: 构建SQLite数据库失败，将暂不启用SQLite: {ex.Message} {ex.StackTrace}"); } catch { }
                _sqlite = null;
            }
        }

        /// <summary>
        /// 解析工作表
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="worksheet">工作表模型</param>
        private void ParseSheet(ISheet sheet, Worksheet worksheet)
        {
            WriteDebug($"开始解析工作表 {sheet.SheetName}，总行数: {sheet.LastRowNum + 1}");

            if (sheet.LastRowNum < 3)
            {
                WriteDebug($"工作表 {sheet.SheetName} 行数不足，跳过");
                // 空表或数据不足
                return;
            }

            // 智能判断字段名称位置（第一行还是第二行）
            int headerRowIndex = DetermineHeaderRowIndex(sheet);
            int typeRowIndex = 0;
            if (headerRowIndex == 0) typeRowIndex = 1;
            int dataStartRowIndex = FindDataStartRow(sheet);

            WriteDebug($"工作表 {sheet.SheetName} - 字段名行: {headerRowIndex}, 数据开始行: {dataStartRowIndex}");

            // 解析表头
            var headerRow = sheet.GetRow(headerRowIndex);
            if (headerRow == null) return;

            var typeRow = sheet.GetRow(typeRowIndex);
            if (typeRow == null) return;

            // 第三行（描述/COMMENTS）按惯例为 headerRowIndex + 2
            var commentsRow = sheet.GetRow(headerRowIndex + 2);

            // 创建列定义
            for (int i = 0; i < headerRow.LastCellNum; i++)
            {
                var cell = headerRow.GetCell(i);
                if (cell != null)
                {
                    var columnName = cell.ToString();
                    if (!string.IsNullOrEmpty(columnName))
                    {
                        string comments = null;
                        if (commentsRow != null)
                        {
                            var ccell = commentsRow.GetCell(i);
                            if (ccell != null)
                            {
                                var cval = GetCellValue(ccell);
                                comments = cval?.ToString();
                            }
                        }

                        var column = new Column
                        {
                            Name = columnName,
                            Index = i,
                            DataType = typeRow.GetCell(i).ToString(), // 推断数据类型
                            Comments = comments
                        };
                        worksheet.Headers.Add(column);
                    }
                }
            }

            // 解析数据行
            for (int i = dataStartRowIndex; i <= sheet.LastRowNum; i++)
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
        /// 判断字段名称所在行索引
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <returns>字段名所在行索引（0或1）</returns>
        private int DetermineHeaderRowIndex(ISheet sheet)
        {
            // 检查第一行和第二行的A列
            var row1 = sheet.GetRow(0);
            var row2 = sheet.GetRow(1);

            if (row1 == null && row2 == null)
                return 0; // 默认第一行

            if (row1 == null)
                return 1; // 只有第二行，选择第二行

            if (row2 == null)
                return 0; // 只有第一行，选择第一行

            // 检查A1和A2单元格
            var cellA1 = row1.GetCell(0);
            var cellA2 = row2.GetCell(0);

            if (cellA1 == null && cellA2 == null)
                return 0; // 都为空，默认第一行

            if (cellA1 == null)
                return 1; // A1为空，选择第二行

            if (cellA2 == null)
                return 0; // A2为空，选择第一行

            // 尝试解析A1和A2的值
            var valueA1 = GetCellValue(cellA1);
            var valueA2 = GetCellValue(cellA2);

            if(valueA1?.ToString().ToLower() == "int") return 1;
            if(valueA2?.ToString().ToLower() == "int") return 0;

            // 判断哪个更像数字
            bool isA1Numeric = IsNumeric(valueA1);
            bool isA2Numeric = IsNumeric(valueA2);

            if (isA1Numeric && !isA2Numeric)
            {
                // A1是数字，A2不是，说明A2是字段名
                return 1;
            }
            else if (!isA1Numeric && isA2Numeric)
            {
                // A1不是数字，A2是，说明A1是字段名
                return 0;
            }
            else
            {
                // 无法明确判断，默认第一行
                return 0;
            }
        }

        /// <summary>
        /// 找到A列为数字的第一行（数据开始行）
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <returns>数据开始行索引</returns>
        private int FindDataStartRow(ISheet sheet)
        {
            // 从第0行开始，找到A列为数字的第一行
            for (int i = 0; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row != null)
                {
                    var cellA = row.GetCell(0);
                    if (cellA != null)
                    {
                        var valueA = GetCellValue(cellA);
                        if (IsNumeric(valueA))
                        {
                            return i;
                        }
                    }
                }
            }

            // 如果没有找到数字行，默认从字段名行的下一行开始
            return DetermineHeaderRowIndex(sheet) + 1;
        }

        /// <summary>
        /// 判断值是否为数字类型
        /// </summary>
        /// <param name="value">要判断的值</param>
        /// <returns>是否为数字</returns>
        private bool IsNumeric(object value)
        {
            if (value == null)
                return false;

            if (value is int || value is double || value is decimal)
                return true;

            if (value is string strValue)
            {
                return double.TryParse(strValue, out _) || int.TryParse(strValue, out _);
            }

            return false;
        }

        /// <summary>
        /// 推断列的数据类型
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="columnIndex">列索引</param>
        /// <param name="dataStartRowIndex">数据开始行索引</param>
        /// <returns>数据类型</returns>
        private string InferColumnType(ISheet sheet, int columnIndex, int dataStartRowIndex)
        {
            // 从数据开始行检查数据类型
            for (int i = dataStartRowIndex; i <= Math.Min(sheet.LastRowNum, dataStartRowIndex + 10); i++)
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
        /// 直接在SQLite上执行原始SQL（SELECT返回结果集，非查询返回影响行数）
        /// </summary>
        public object ExecuteSqlRaw(string sql)
        {
            if (string.IsNullOrWhiteSpace(sql)) throw new ArgumentException("SQL不能为空");
            var upper = sql.TrimStart().ToUpperInvariant();
            if (upper.StartsWith("SELECT"))
            {
                return _sqlite.ExecuteQuery(sql);
            }
            else if (upper.StartsWith("SHOW TABLES"))
            {
                return _sqlite.GetTables();
            }
            else if (upper.StartsWith("SHOW CREATE TABLE"))
            {
                var table = SqlParser.ParseShowCreateTable(sql);
                var stmt = _sqlite.GetCreateTable(table);
                return new Dictionary<string, string> { { "table", table }, { "createTable", stmt } };
            }
            else
            {
                var affected = _sqlite.ExecuteNonQuery(sql);
                return new Dictionary<string, object> { { "affectedRows", affected } };
            }
        }

        /// <summary>
        /// 获取所有表名
        /// </summary>
        /// <returns>表名列表</returns>
        public List<string> GetTableNames()
        {
            // 优先返回SQLite中的表名；若尚未构建，回退旧缓存
            try
            {
                if (_sqlite != null) return _sqlite.GetTables();
            }
            catch { }
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
            if (_sqlite != null)
            {
                try { return _sqlite.GetCreateTable(tableName); } catch { }
            }
            // fallback to old
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
        /// <param name="limit">LIMIT限制</param>
        /// <returns>查询结果</returns>
        public List<Dictionary<string, object>> ExecuteSelect(string tableName, List<string> columns, string whereClause = null, int? limit = null)
        {
            foreach (var excelFile in _excelFiles.Values)
            {
                if (excelFile.Worksheets.ContainsKey(tableName))
                {
                    var worksheet = excelFile.Worksheets[tableName];
                    var result = new List<Dictionary<string, object>>();

                    // 检查是否包含COUNT(*)
                    bool hasCountStar = columns != null && columns.Any(c => c.ToUpper() == "COUNT(*)");

                    if (hasCountStar)
                    {
                        // 处理COUNT(*)查询
                        var filteredRows = ApplyWhereClause(worksheet.DataRows, whereClause);
                        var countResult = new Dictionary<string, object>
                        {
                            { "COUNT(*)", filteredRows.Count }
                        };
                        result.Add(countResult);
                        return result;
                    }

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
                    var filteredDataRows = ApplyWhereClause(worksheet.DataRows, whereClause);

                    // 应用LIMIT限制
                    if (limit.HasValue && limit > 0)
                    {
                        filteredDataRows = filteredDataRows.Take(limit.Value).ToList();
                    }

                    // 构建结果
                    foreach (var row in filteredDataRows)
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
        /// 执行SELECT查询（支持Excel文件名）
        /// </summary>
        /// <param name="fileNameOrTableName">Excel文件名（不带.xlsx）或工作表名称</param>
        /// <param name="columns">要查询的列</param>
        /// <param name="whereClause">WHERE条件</param>
        /// <param name="limit">LIMIT限制</param>
        /// <returns>查询结果</returns>
        public List<Dictionary<string, object>> ExecuteSelectByFileName(string fileNameOrTableName, List<string> columns, string whereClause = null, int? limit = null)
        {
            var results = new List<Dictionary<string, object>>();

            // 首先尝试按Excel文件名查询
            var excelFileName = fileNameOrTableName.EndsWith(".xlsx") ? fileNameOrTableName : fileNameOrTableName + ".xlsx";

            if (_excelFiles.ContainsKey(excelFileName))
            {
                var excelFile = _excelFiles[excelFileName];

                // 遍历该Excel文件的所有工作表
                foreach (var worksheet in excelFile.Worksheets.Values)
                {
                    var sheetResults = ExecuteSelectFromWorksheet(worksheet, columns, whereClause, limit, excelFileName);
                    results.AddRange(sheetResults);
                }

                return results;
            }

            // 如果找不到Excel文件，回退到原来的工作表查询
            try
            {
                return ExecuteSelect(fileNameOrTableName, columns, whereClause, limit);
            }
            catch
            {
                // 如果也找不到工作表，提供友好的错误信息
                var availableFiles = _excelFiles.Keys.ToList();
                var availableTables = GetTableNames();

                var errorMessage = $"找不到 '{fileNameOrTableName}' 对应的Excel文件或工作表。\n\n";
                errorMessage += "可用的Excel文件包括：\n";
                foreach (var file in availableFiles)
                {
                    errorMessage += $"  - {file}\n";
                }

                errorMessage += "\n可用的工作表包括：\n";
                foreach (var table in availableTables)
                {
                    errorMessage += $"  - {table}\n";
                }

                errorMessage += "\n提示：您可以使用Excel文件名（不带.xlsx后缀）来查询该文件的所有工作表。";

                throw new Exception(errorMessage);
            }
        }

        /// <summary>
        /// 从单个工作表执行查询
        /// </summary>
        /// <param name="worksheet">工作表</param>
        /// <param name="columns">要查询的列</param>
        /// <param name="whereClause">WHERE条件</param>
        /// <param name="limit">LIMIT限制</param>
        /// <param name="sourceFileName">源Excel文件名</param>
        /// <returns>查询结果</returns>
        private List<Dictionary<string, object>> ExecuteSelectFromWorksheet(Worksheet worksheet, List<string> columns, string whereClause, int? limit, string sourceFileName)
        {
            var result = new List<Dictionary<string, object>>();

            // 检查是否包含COUNT(*)
            bool hasCountStar = columns != null && columns.Any(c => c.ToUpper() == "COUNT(*)");

            if (hasCountStar)
            {
                // 处理COUNT(*)查询
                var filteredRows = ApplyWhereClause(worksheet.DataRows, whereClause);
                var countResult = new Dictionary<string, object>
                {
                    { "COUNT(*)", filteredRows.Count },
                    { "_source_file", sourceFileName },
                    { "_source_sheet", worksheet.Name }
                };
                result.Add(countResult);
                return result;
            }

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
            var filteredDataRows = ApplyWhereClause(worksheet.DataRows, whereClause);

            // 应用LIMIT限制
            if (limit.HasValue && limit > 0)
            {
                filteredDataRows = filteredDataRows.Take(limit.Value).ToList();
            }

            // 构建结果，添加源文件和工作表信息
            foreach (var row in filteredDataRows)
            {
                var resultRow = new Dictionary<string, object>();
                foreach (var column in selectedColumns)
                {
                    row.TryGetValue(column.Name, out var value);
                    resultRow[column.Name] = value;
                }

                // 添加元数据信息，标识数据来源
                resultRow["_source_file"] = sourceFileName;
                resultRow["_source_sheet"] = worksheet.Name;

                result.Add(resultRow);
            }

            return result;
        }

        /// <summary>
        /// 执行带JOIN的SELECT查询
        /// </summary>
        /// <param name="select">SELECT语句对象</param>
        /// <returns>查询结果</returns>
        public List<Dictionary<string, object>> ExecuteSelectWithJoin(SelectStatement select)
        {
            // 获取主表数据
            var mainTableData = GetTableData(select.TableName);
            var joinedData = mainTableData;

            // 逐个应用JOIN
            foreach (var join in select.Joins)
            {
                var joinTableData = GetTableData(join.TableName);
                joinedData = PerformJoin(joinedData, joinTableData, join);
            }

            // 应用WHERE条件
            var filteredData = ApplyWhereClause(joinedData, select.WhereClause);

            // 应用LIMIT
            if (select.Limit.HasValue)
            {
                filteredData = filteredData.Take(select.Limit.Value).ToList();
            }

            // 选择指定的列
            return SelectColumns(filteredData, select.Columns);
        }

        /// <summary>
        /// 获取表数据
        /// </summary>
        /// <param name="tableName">表名</param>
        /// <returns>表数据</returns>
        private List<Dictionary<string, object>> GetTableData(string tableName)
        {
            foreach (var excelFile in _excelFiles.Values)
            {
                if (excelFile.Worksheets.ContainsKey(tableName))
                {
                    var worksheet = excelFile.Worksheets[tableName];
                    var result = new List<Dictionary<string, object>>();

                    foreach (var row in worksheet.DataRows)
                    {
                        var resultRow = new Dictionary<string, object>();
                        foreach (var column in worksheet.Headers)
                        {
                            row.TryGetValue(column.Name, out var value);
                            resultRow[column.Name] = value;
                        }
                        result.Add(resultRow);
                    }

                    return result;
                }
            }

            throw new Exception($"表 '{tableName}' 不存在");
        }

        /// <summary>
        /// 执行JOIN操作
        /// </summary>
        /// <param name="leftData">左表数据</param>
        /// <param name="rightData">右表数据</param>
        /// <param name="join">JOIN子句</param>
        /// <returns>JOIN结果</returns>
        private List<Dictionary<string, object>> PerformJoin(List<Dictionary<string, object>> leftData, List<Dictionary<string, object>> rightData, JoinClause join)
        {
            var result = new List<Dictionary<string, object>>();

            // 解析连接条件，例如 "table1.id = table2.user_id"
            var conditions = join.OnClause.Split(new[] { " AND " }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var leftRow in leftData)
            {
                var joined = false;

                foreach (var rightRow in rightData)
                {
                    var match = true;

                    // 检查所有连接条件
                    foreach (var condition in conditions)
                    {
                        if (!EvaluateJoinCondition(condition, leftRow, rightRow))
                        {
                            match = false;
                            break;
                        }
                    }

                    if (match)
                    {
                        // 合并行数据
                        var mergedRow = new Dictionary<string, object>(leftRow);

                        // 添加右表的列，使用表名前缀避免冲突
                        foreach (var kvp in rightRow)
                        {
                            mergedRow[$"{join.TableName}.{kvp.Key}"] = kvp.Value;
                        }

                        result.Add(mergedRow);
                        joined = true;

                        // 对于INNER JOIN，只需要第一个匹配
                        if (join.JoinType.ToUpper() == "INNER")
                        {
                            break;
                        }
                    }
                }

                // 对于LEFT JOIN，如果没有匹配，仍保留左表行
                if (!joined && join.JoinType.ToUpper() == "LEFT")
                {
                    var mergedRow = new Dictionary<string, object>(leftRow);

                    // 为右表列添加NULL值
                    foreach (var rightRow in rightData.FirstOrDefault() ?? new Dictionary<string, object>())
                    {
                        mergedRow[$"{join.TableName}.{rightRow.Key}"] = null;
                    }

                    result.Add(mergedRow);
                }
            }

            return result;
        }

        /// <summary>
        /// 评估JOIN条件
        /// </summary>
        /// <param name="condition">条件表达式</param>
        /// <param name="leftRow">左表行</param>
        /// <param name="rightRow">右表行</param>
        /// <returns>是否满足条件</returns>
        private bool EvaluateJoinCondition(string condition, Dictionary<string, object> leftRow, Dictionary<string, object> rightRow)
        {
            // 简单的条件解析，支持 table.column = table.column 格式
            var equalityMatch = System.Text.RegularExpressions.Regex.Match(condition, @"(\w+)\.(\w+)\s*=\s*(\w+)\.(\w+)");
            if (equalityMatch.Success)
            {
                var leftTable = equalityMatch.Groups[1].Value;
                var leftColumn = equalityMatch.Groups[2].Value;
                var rightTable = equalityMatch.Groups[3].Value;
                var rightColumn = equalityMatch.Groups[4].Value;

                // 获取对应的值
                object leftValue = null;
                object rightValue = null;

                if (leftRow.ContainsKey(leftColumn))
                {
                    leftValue = leftRow[leftColumn];
                }
                else if (leftRow.ContainsKey($"{leftTable}.{leftColumn}"))
                {
                    leftValue = leftRow[$"{leftTable}.{leftColumn}"];
                }

                if (rightRow.ContainsKey(rightColumn))
                {
                    rightValue = rightRow[rightColumn];
                }
                else if (rightRow.ContainsKey($"{rightTable}.{rightColumn}"))
                {
                    rightValue = rightRow[$"{rightTable}.{rightColumn}"];
                }

                // 比较值
                if (leftValue == null && rightValue == null)
                    return true;
                if (leftValue == null || rightValue == null)
                    return false;

                return leftValue.Equals(rightValue);
            }

            // 更复杂的条件可以在这里添加
            throw new ArgumentException($"不支持的JOIN条件: {condition}");
        }

        /// <summary>
        /// 选择指定的列
        /// </summary>
        /// <param name="data">数据</param>
        /// <param name="columns">列名列表</param>
        /// <returns>选择结果</returns>
        private List<Dictionary<string, object>> SelectColumns(List<Dictionary<string, object>> data, List<string> columns)
        {
            if (columns == null || columns.Count == 0 || (columns.Count == 1 && columns[0] == "*"))
            {
                return data;
            }

            var result = new List<Dictionary<string, object>>();

            foreach (var row in data)
            {
                var resultRow = new Dictionary<string, object>();

                foreach (var column in columns)
                {
                    // 处理带表名前缀的列
                    if (column.Contains("."))
                    {
                        if (row.ContainsKey(column))
                        {
                            resultRow[column] = row[column];
                        }
                    }
                    else
                    {
                        // 不带表名前缀的列，查找匹配的列
                        var matchingKey = row.Keys.FirstOrDefault(k => k.EndsWith($".{column}") || k == column);
                        if (matchingKey != null)
                        {
                            resultRow[column] = row[matchingKey];
                        }
                    }
                }

                result.Add(resultRow);
            }

            return result;
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

            var filteredRows = new List<Dictionary<string, object>>();

            // 移除WHERE关键字
            var condition = whereClause.Trim();
            if (condition.ToUpper().StartsWith("WHERE "))
            {
                condition = condition.Substring(6).Trim();
            }

            // 使用新的数学表达式求值器
            foreach (var row in dataRows)
            {
                try
                {
                    if (SqlParser.EvaluateWhereClause(condition, row))
                    {
                        filteredRows.Add(row);
                    }
                }
                catch (Exception ex)
                {
                    // 如果表达式求值失败，记录错误但不中断处理
                    Console.WriteLine($"WHERE条件求值失败，跳过该行: {ex.Message}");
                }
            }

            return filteredRows;
        }

        /// <summary>
        /// 刷新文件缓存（重建SQLite）
        /// </summary>
        public void Refresh()
        {
            _excelFiles.Clear();
            LoadAllExcelFiles();
        }

        /// <summary>
        /// 获取Excel文件中所有工作表的建表语句
        /// </summary>
        /// <param name="fileNameOrTableName">Excel文件名（不带.xlsx）或工作表名称</param>
        /// <returns>建表语句列表</returns>
        public List<Dictionary<string, string>> GetCreateTableStatementsByFileName(string fileNameOrTableName)
        {
            var results = new List<Dictionary<string, string>>();

            // 首先尝试按Excel文件名查询
            var excelFileName = fileNameOrTableName.EndsWith(".xlsx") ? fileNameOrTableName : fileNameOrTableName + ".xlsx";

            if (_excelFiles.ContainsKey(excelFileName))
            {
                var excelFile = _excelFiles[excelFileName];

                // 遍历该Excel文件的所有工作表
                foreach (var worksheet in excelFile.Worksheets.Values)
                {
                    try
                    {
                        var createTableSql = GenerateCreateTableSql(worksheet);
                        var result = new Dictionary<string, string>
                        {
                            { "table", worksheet.Name },
                            { "createTable", createTableSql },
                            { "source_file", excelFileName }
                        };
                        results.Add(result);
                    }
                    catch (Exception ex)
                    {
                        var result = new Dictionary<string, string>
                        {
                            { "table", worksheet.Name },
                            { "createTable", $"-- 无法生成建表语句: {ex.Message}" },
                            { "source_file", excelFileName }
                        };
                        results.Add(result);
                    }
                }

                return results;
            }

            // 如果找不到Excel文件，回退到原来的单个工作表查询
            try
            {
                var createTableStatement = GetCreateTableStatement(fileNameOrTableName);
                return new List<Dictionary<string, string>>
                {
                    new Dictionary<string, string>
                    {
                        { "table", fileNameOrTableName },
                        { "createTable", createTableStatement },
                        { "source_file", "单个工作表" }
                    }
                };
            }
            catch
            {
                // 如果也找不到工作表，提供友好的错误信息
                var availableFiles = _excelFiles.Keys.ToList();
                var availableTables = GetTableNames();

                var errorMessage = $"找不到 '{fileNameOrTableName}' 对应的Excel文件或工作表。\n\n";
                errorMessage += "可用的Excel文件包括：\n";
                foreach (var file in availableFiles)
                {
                    errorMessage += $"  - {file}\n";
                }

                errorMessage += "\n可用的工作表包括：\n";
                foreach (var table in availableTables)
                {
                    errorMessage += $"  - {table}\n";
                }

                throw new Exception(errorMessage);
            }
        }

        /// <summary>
        /// 生成工作表的建表语句
        /// </summary>
        /// <param name="worksheet">工作表</param>
        /// <returns>建表语句</returns>
        private string GenerateCreateTableSql(Worksheet worksheet)
        {
            var sql = $"CREATE TABLE {worksheet.Name} (\n";

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

        /// <summary>
        /// 执行UPDATE语句
        /// </summary>
        /// <param name="fileNameOrTableName">Excel文件名或工作表名称</param>
        /// <param name="setValues">要更新的字段和值</param>
        /// <param name="whereClause">WHERE条件</param>
        /// <returns>更新的行数</returns>
        public int ExecuteUpdate(string fileNameOrTableName, Dictionary<string, string> setValues, string whereClause = null)
        {
            if (_sqlite == null) throw new Exception("SQLite尚未初始化");
            // 构造UPDATE语句（简单版）
            var sb = new System.Text.StringBuilder();
            sb.Append("UPDATE ").Append('"').Append(fileNameOrTableName).Append('"').Append(" SET ");
            bool first = true;
            foreach (var kv in setValues)
            {
                if (!first) sb.Append(", ");
                first = false;
                sb.Append('"').Append(kv.Key).Append('"').Append(" = ");
                sb.Append('"').Append(kv.Value.Replace("\"", "\"\"")).Append('"');
            }
            if (!string.IsNullOrWhiteSpace(whereClause))
            {
                sb.Append(" WHERE ").Append(whereClause);
            }
            var res = _sqlite.ExecuteNonQuery(sb.ToString());
            return res;
        }

        /// <summary>
        /// 执行DELETE语句
        /// </summary>
        /// <param name="fileNameOrTableName">Excel文件名或工作表名称</param>
        /// <param name="whereClause">WHERE条件</param>
        /// <returns>删除的行数</returns>
        public int ExecuteDelete(string fileNameOrTableName, string whereClause = null)
        {
            if (_sqlite == null) throw new Exception("SQLite尚未初始化");
            var sql = new System.Text.StringBuilder();
            sql.Append("DELETE FROM ").Append('"').Append(fileNameOrTableName).Append('"');
            if (!string.IsNullOrWhiteSpace(whereClause))
            {
                sql.Append(" WHERE ").Append(whereClause);
            }
            return _sqlite.ExecuteNonQuery(sql.ToString());
        }

        /// <summary>
        /// 保存所有修改回Excel：从SQLite导出
        /// </summary>
        public int SaveAllChanges()
        {
            var savedFiles = 0;
            foreach (var excelFile in _excelFiles.Values)
            {
                if (ExportExcelFileFromSqlite(excelFile)) savedFiles++;
            }
            return savedFiles;
        }

        public bool SaveChanges(string fileName)
        {
            var excelFileName = fileName.EndsWith(".xlsx") ? fileName : fileName + ".xlsx";
            if (_excelFiles.ContainsKey(excelFileName))
            {
                var excelFile = _excelFiles[excelFileName];
                return ExportExcelFileFromSqlite(excelFile);
            }
            return false;
        }

        private bool ExportExcelFileFromSqlite(ExcelFile excelFile)
        {
            try
            {
                var tempFilePath = Path.GetTempFileName();
                File.Move(tempFilePath, tempFilePath + ".xlsx");
                tempFilePath += ".xlsx";

                XSSFWorkbook workbook = null;
                // 优先以原文件为基础，保留头部样式/列宽/合并等
                if (File.Exists(excelFile.Path))
                {
                    using (var srcStream = new FileStream(excelFile.Path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        workbook = new XSSFWorkbook(srcStream);
                    }
                }
                else
                {
                    workbook = new XSSFWorkbook();
                }

                foreach (var worksheet in excelFile.Worksheets.Values)
                {
                    var sheet = workbook.GetSheet(worksheet.Name) ?? workbook.CreateSheet(worksheet.Name);

                    // 如果工作表存在，尽量保留前3行头部（样式/合并/列宽等保持不变）
                    // 确保存在3行头部，不存在则补齐（补齐的样式无法还原，仅写入文本）
                    for (int i = 0; i < 3; i++)
                    {
                        var row = sheet.GetRow(i) ?? sheet.CreateRow(i);
                        if (row.Cells.Count == 0)
                        {
                            // 当头部不存在时，按当前模型写入文本（样式不可还原）
                            for (int c = 0; c < worksheet.Headers.Count; c++)
                            {
                                var cell = row.GetCell(c) ?? row.CreateCell(c);
                                if (i == 0) cell.SetCellValue(worksheet.Headers[c].Name);
                                else if (i == 1) cell.SetCellValue(worksheet.Headers[c].DataType);
                                else cell.SetCellValue("");
                            }
                        }
                    }

                    // 清除旧数据行（保留0..2行）
                    for (int r = sheet.LastRowNum; r >= 3; r--)
                    {
                        var oldRow = sheet.GetRow(r);
                        if (oldRow != null)
                        {
                            sheet.RemoveRow(oldRow);
                        }
                    }

                    // 从SQLite读取该表最新数据
                    var rows = _sqlite.ExecuteQuery($"SELECT * FROM \"{worksheet.Name}\"");
                    worksheet.DataRows = rows;

                    // 写入数据行（从第4行开始）
                    WriteDataRows(sheet, worksheet);
                }

                // 写出到临时文件
                using (var outStream = new FileStream(tempFilePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(outStream);
                }

                // 备份原文件并替换
                var backupFilePath = excelFile.Path + ".backup";
                if (File.Exists(excelFile.Path))
                {
                    if (File.Exists(backupFilePath)) File.Delete(backupFilePath);
                    File.Copy(excelFile.Path, backupFilePath);
                }
                if (File.Exists(excelFile.Path)) File.Delete(excelFile.Path);
                File.Move(tempFilePath, excelFile.Path);
                excelFile.LastModified = File.GetLastWriteTime(excelFile.Path);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"保存Excel文件失败: {excelFile.Path}, 错误: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 获取修改统计信息
        /// </summary>
        /// <returns>统计信息</returns>
        public Dictionary<string, object> GetModificationStats()
        {
            var stats = new Dictionary<string, object>
            {
                { "loaded_files", _excelFiles.Count },
                { "total_worksheets", _excelFiles.Values.Sum(f => f.Worksheets.Count) },
                { "total_rows", _excelFiles.Values.Sum(f => f.Worksheets.Values.Sum(w => w.DataRows.Count)) }
            };

            return stats;
        }

        /// <summary>
        /// 撤销所有未保存的修改（重新加载文件）
        /// </summary>
        public void UndoChanges()
        {
            Refresh();
        }

        private void WriteHeaders(ISheet sheet, Worksheet worksheet)
        {
            // 第一行：列名
            var headerRow = sheet.GetRow(0) ?? sheet.CreateRow(0);
            for (int i = 0; i < worksheet.Headers.Count; i++)
            {
                var cell = headerRow.GetCell(i) ?? headerRow.CreateCell(i);
                cell.SetCellValue(worksheet.Headers[i].Name);
            }

            // 第二行：数据类型
            var typeRow = sheet.GetRow(1) ?? sheet.CreateRow(1);
            for (int i = 0; i < worksheet.Headers.Count; i++)
            {
                var cell = typeRow.GetCell(i) ?? typeRow.CreateCell(i);
                cell.SetCellValue(worksheet.Headers[i].DataType);
            }

            // 第三行：描述（COMMENTS）
            var descRow = sheet.GetRow(2) ?? sheet.CreateRow(2);
            for (int i = 0; i < worksheet.Headers.Count; i++)
            {
                var cell = descRow.GetCell(i) ?? descRow.CreateCell(i);
                var comments = worksheet.Headers[i].Comments;
                cell.SetCellValue(string.IsNullOrEmpty(comments) ? "" : comments);
            }
        }

        private void WriteDataRows(ISheet sheet, Worksheet worksheet)
        {
            for (int rowIndex = 0; rowIndex < worksheet.DataRows.Count; rowIndex++)
            {
                var dataRow = worksheet.DataRows[rowIndex];
                var sheetRow = sheet.CreateRow(rowIndex + 3); // 从第4行开始（前3行是表头）

                for (int colIndex = 0; colIndex < worksheet.Headers.Count; colIndex++)
                {
                    var columnName = worksheet.Headers[colIndex].Name;
                    var cell = sheetRow.CreateCell(colIndex);

                    if (dataRow.ContainsKey(columnName))
                    {
                        var value = dataRow[columnName];
                        SetCellValue(cell, value);
                    }
                    else
                    {
                        cell.SetCellValue("");
                    }
                }
            }
        }

        private void SetCellValue(ICell cell, object value)
        {
            if (value == null || value == DBNull.Value)
            {
                cell.SetCellValue("");
            }
            else if (value is string)
            {
                cell.SetCellValue((string)value);
            }
            else if (value is int)
            {
                cell.SetCellValue((int)value);
            }
            else if (value is long)
            {
                cell.SetCellValue(Convert.ToDouble(value));
            }
            else if (value is double)
            {
                cell.SetCellValue((double)value);
            }
            else if (value is float)
            {
                cell.SetCellValue(Convert.ToDouble(value));
            }
            else if (value is bool)
            {
                cell.SetCellValue((bool)value);
            }
            else if (value is DateTime)
            {
                cell.SetCellValue((DateTime)value);
            }
            else
            {
                cell.SetCellValue(value.ToString());
            }
        }
    }
}