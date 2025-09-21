using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using ExcelSqlTool.Tools;

namespace ExcelSqlTool
{
    /// <summary>
    /// MCP服务器实现，直接提供MCP工具服务
    /// </summary>
    public class McpServer
    {
        private readonly ExcelManager _excelManager;
        private readonly StreamReader _input;
        private readonly StreamWriter _output;
        private bool _isRunning = false;
        private readonly Dictionary<string, ToolBase> _tools = new Dictionary<string, ToolBase>(StringComparer.OrdinalIgnoreCase);

        public McpServer(ExcelManager excelManager, Stream input, Stream output)
        {
            _excelManager = excelManager;
            _input = new StreamReader(input, new UTF8Encoding(false)); // 禁用BOM
            _output = new StreamWriter(output, new UTF8Encoding(false)) { AutoFlush = true }; // 禁用BOM

            // 注册工具
            RegisterTools();
        }

        private void RegisterTools()
        {
            // 这里集中注册所有工具实例
            AddTool(new ShowTablesTool(_excelManager));
            AddTool(new QueryTool(_excelManager));
            AddTool(new GetTableSchemaTool(_excelManager));
            AddTool(new RefreshCacheTool(_excelManager));
            AddTool(new ListSheetsTool(_excelManager));
            AddTool(new ChangeDirectoryTool(_excelManager));
            AddTool(new SaveAllTool(_excelManager));
            AddTool(new SaveFileTool(_excelManager));
        }

        private void AddTool(ToolBase tool)
        {
            if (tool == null || string.IsNullOrEmpty(tool.name)) return;
            _tools[tool.name] = tool;
        }

        /// <summary>
        /// 移除字符串中的BOM标记
        /// </summary>
        /// <param name="str">输入字符串</param>
        /// <returns>移除BOM后的字符串</returns>
        private string RemoveBOM(string str)
        {
            if (string.IsNullOrEmpty(str))
                return str;

            // UTF-8 BOM is EF BB BF
            if (str.Length > 0 && str[0] == '\uFEFF')
                return str.Substring(1);

            return str;
        }

        /// <summary>
        /// 启动MCP服务器
        /// </summary>
        public async Task StartAsync()
        {
            _isRunning = true;

            try
            {
                while (_isRunning)
                {
                    try
                    {
                        // 读取JSON-RPC请求
                        var requestLine = await _input.ReadLineAsync();
                        if (string.IsNullOrEmpty(requestLine))
                        {
                            await Task.Delay(100); // 避免CPU占用过高
                            continue;
                        }

                        // 移除BOM标记
                        requestLine = RemoveBOM(requestLine);

                        // 解析请求
                        var request = JObject.Parse(requestLine);
                        var response = await ProcessRequestAsync(request);

                        // 发送响应
                        var responseJson = JsonConvert.SerializeObject(response, Formatting.None);
                        responseJson = RemoveBOM(responseJson);
                        await _output.WriteLineAsync(responseJson);
                        await _output.FlushAsync();
                    }
                    catch (Exception ex)
                    {
                        // 发送错误响应，避免日志输出
                        var errorResponse = new
                        {
                            jsonrpc = "2.0",
                            id = (object)null,
                            error = new
                            {
                                code = -32603,
                                message = "Internal error",
                                data = ex.Message
                            }
                        };

                        var errorJson = JsonConvert.SerializeObject(errorResponse, Formatting.None);
                        errorJson = RemoveBOM(errorJson);
                        await _output.WriteLineAsync(errorJson);
                        await _output.FlushAsync();
                    }
                }
            }
            catch
            {
                // 静默处理错误，避免输出干扰
            }
        }

        /// <summary>
        /// 处理MCP请求
        /// </summary>
        /// <param name="request">请求对象</param>
        /// <returns>响应对象</returns>
        private async Task<object> ProcessRequestAsync(JObject request)
        {
            var method = request["method"]?.ToString();
            var id = request["id"];

            try
            {
                switch (method)
                {
                    case "initialize":
                        return await HandleInitializeAsync(request, id);
                    case "initialized":
                        return new { jsonrpc = "2.0", id, result = (object)null };
                    case "shutdown":
                        return await HandleShutdownAsync(id);
                    case "list_tools":
                    case "tools/list":
                        return await HandleListToolsAsync(id);
                    case "call_tool":
                    case "tools/call":
                        return await HandleCallToolAsync(request, id);
                    default:
                        return new
                        {
                            jsonrpc = "2.0",
                            id,
                            error = new
                            {
                                code = -32601,
                                message = "Method not found",
                                data = $"未知方法: {method}"
                            }
                        };
                }
            }
            catch (Exception ex)
            {
                return new
                {
                    jsonrpc = "2.0",
                    id,
                    error = new
                    {
                        code = -32603,
                        message = "Internal error",
                        data = ex.Message
                    }
                };
            }
        }

        /// <summary>
        /// 处理初始化请求
        /// </summary>
        /// <param name="request">请求对象</param>
        /// <param name="id">请求ID</param>
        /// <returns>响应对象</returns>
        private Task<object> HandleInitializeAsync(JObject request, object id)
        {
            var protocolVersion = request["params"]?["protocolVersion"]?.ToString();
            var capabilities = request["params"]?["capabilities"]?.ToObject<JObject>();

            var response = new
            {
                jsonrpc = "2.0",
                id,
                result = new
                {
                    protocolVersion = "2024-11-05",
                    capabilities = new
                    {
                        tools = new
                        {
                            listChanged = false
                        }
                    },
                    serverInfo = new
                    {
                        name = "excel-sql-tool",
                        version = "1.0.0"
                    }
                }
            };

            return Task.FromResult<object>(response);
        }

        /// <summary>
        /// 处理关闭请求
        /// </summary>
        /// <param name="id">请求ID</param>
        /// <returns>响应对象</returns>
        private Task<object> HandleShutdownAsync(object id)
        {
            _isRunning = false;
            var response = new
            {
                jsonrpc = "2.0",
                id,
                result = (object)null
            };

            return Task.FromResult<object>(response);
        }

        /// <summary>
        /// 处理列出工具请求
        /// </summary>
        /// <param name="id">请求ID</param>
        /// <returns>响应对象</returns>
        private Task<object> HandleListToolsAsync(object id)
        {
            var tools = new List<object>();
            foreach (var kv in _tools)
            {
                tools.Add(((ToolBase)kv.Value).Info);
            }

            var response = new
            {
                jsonrpc = "2.0",
                id,
                result = new { tools = tools.ToArray() }
            };

            return Task.FromResult<object>(response);
        }

        /// <summary>
        /// 处理调用工具请求
        /// </summary>
        /// <param name="request">请求对象</param>
        /// <param name="id">请求ID</param>
        /// <returns>响应对象</returns>
        private async Task<object> HandleCallToolAsync(JObject request, object id)
        {
            var name = request["params"]?["name"]?.ToString();
            var arguments = request["params"]?["arguments"] as JObject;

            try
            {
                // 智能解析参数
                var parsedArguments = ParameterHelper.SmartParseParameters(arguments);

                if (string.IsNullOrEmpty(name) || !_tools.ContainsKey(name))
                {
                    throw new ArgumentException($"未知工具: {name}");
                }

                // 只有更换目录工具在目录不存在时可以调用
                if (!string.Equals(name, "excel_change_directory", StringComparison.OrdinalIgnoreCase) && !_excelManager.IsDirectoryExists)
                {
                    return new
                    {
                        jsonrpc = "2.0",
                        id,
                        error = new
                        {
                            code = -32000,
                            message = "Excel文件目录不存在或未设置",
                            data = "请先使用'excel_change_directory'工具设置有效的Excel文件目录"
                        }
                    };
                }

                var tool = _tools[name];
                var result = await tool.CallAsync(parsedArguments ?? new JObject());

                var response = new
                {
                    jsonrpc = "2.0",
                    id,
                    result = new
                    {
                        content = new[]
                        {
                            new
                            {
                                type = "text",
                                text = JsonConvert.SerializeObject(result, Formatting.Indented)
                            }
                        }
                    }
                };

                return response;
            }
            catch (Exception ex)
            {
                var errorResponse = new
                {
                    jsonrpc = "2.0",
                    id,
                    error = new
                    {
                        code = -32603,
                        message = "Internal error",
                        data = ex.Message
                    }
                };

                return errorResponse;
            }
        }

        // 以下保留原有的具体实现方法以兼容旧逻辑（若后续不需要可移除）
        private async Task<object> ExecuteSqlAsync(string sql)
        {
            return await Task.Run(() =>
            {
                var statementType = SqlParser.ParseStatementType(sql);

                switch (statementType)
                {
                    case SqlParser.SqlStatementType.Select:
                        var selectStatement = SqlParser.ParseSelect(sql);

                        // 检查是否包含JOIN
                        if (selectStatement.Joins != null && selectStatement.Joins.Count > 0)
                        {
                            return (object)_excelManager.ExecuteSelectWithJoin(selectStatement);
                        }
                        else
                        {
                            return (object)_excelManager.ExecuteSelectByFileName(
                                selectStatement.TableName,
                                selectStatement.Columns,
                                selectStatement.WhereClause,
                                selectStatement.Limit);
                        }
                    case SqlParser.SqlStatementType.Update:
                        var updateStatement = SqlParser.ParseUpdate(sql);
                        var updateResult = _excelManager.ExecuteUpdate(
                            updateStatement.TableName,
                            updateStatement.SetValues,
                            updateStatement.WhereClause);
                        return (object)new Dictionary<string, object>
                        {
                            { "affectedRows", updateResult },
                            { "message", $"成功更新 {updateResult} 行数据" }
                        };
                    case SqlParser.SqlStatementType.Delete:
                        var deleteStatement = SqlParser.ParseDelete(sql);
                        var deleteResult = _excelManager.ExecuteDelete(
                            deleteStatement.TableName,
                            deleteStatement.WhereClause);
                        return (object)new Dictionary<string, object>
                        {
                            { "affectedRows", deleteResult },
                            { "message", $"成功删除 {deleteResult} 行数据" }
                        };
                    case SqlParser.SqlStatementType.ShowTables:
                        return (object)_excelManager.GetTableNames();
                    case SqlParser.SqlStatementType.ShowCreateTable:
                        var tableName = SqlParser.ParseShowCreateTable(sql);
                        var createTableStatement = _excelManager.GetCreateTableStatement(tableName);
                        return (object)new Dictionary<string, string>
                        {
                            { "table", tableName },
                            { "createTable", createTableStatement }
                        };
                    default:
                        throw new ArgumentException($"不支持的SQL语句类型: {statementType}");
                }
            });
        }

        private async Task<object> GetTablesAsync()
        {
            return await Task.Run(() => (object)_excelManager.GetTableNames());
        }

        private async Task<object> GetTableSchemaAsync(string tableName)
        {
            return await Task.Run(() =>
            {
                var results = _excelManager.GetCreateTableStatementsByFileName(tableName);

                // 如果只有一个结果，返回原来的格式以保持兼容性
                if (results.Count == 1)
                {
                    return (object)new Dictionary<string, string>
                    {
                        { "table", results[0]["table"] },
                        { "createTable", results[0]["createTable"] }
                    };
                }

                // 如果有多个结果（Excel文件中有多个工作表），返回列表格式
                return (object)results;
            });
        }

        private async Task<object> RefreshCacheAsync()
        {
            return await Task.Run(() => (object)"缓存已刷新");
        }

        private async Task<object> ListSheetsAsync()
        {
            return await Task.Run(() => (object)_excelManager.GetTableNames());
        }

        private async Task<object> ChangeDirectoryAsync(string newDirectory)
        {
            return await Task.Run(() =>
            {
                var oldDirectory = _excelManager.DirectoryPath;
                _excelManager.UpdateDirectoryPath(newDirectory);

                return (object)new Dictionary<string, string>
                {
                    { "old_directory", oldDirectory },
                    { "new_directory", newDirectory },
                    { "message", "目录已成功更改" }
                };
            });
        }

        private async Task<object> SaveAllAsync()
        {
            return await Task.Run(() =>
            {
                var savedFiles = _excelManager.SaveAllChanges();
                return (object)new Dictionary<string, object>
                {
                    { "saved_files", savedFiles },
                    { "message", $"成功保存 {savedFiles} 个Excel文件的修改" }
                };
            });
        }

        private async Task<object> SaveFileAsync(string fileName)
        {
            return await Task.Run(() =>
            {
                var success = _excelManager.SaveChanges(fileName);
                return (object)new Dictionary<string, object>
                {
                    { "success", success },
                    { "message", success ? $"成功保存文件 {fileName} 的修改" : $"保存文件 {fileName} 失败" }
                };
            });
        }

        /// <summary>
        /// 停止服务器
        /// </summary>
        public void Stop()
        {
            _isRunning = false;
        }
    }
}