using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelSqlTool
{
    /// <summary>
    /// 参数处理助手类
    /// </summary>
    public static class ParameterHelper
    {
        /// <summary>
        /// 智能解析工具参数，处理可能的包装问题
        /// </summary>
        /// <param name="parameters">原始参数</param>
        /// <returns>解析后的参数</returns>
        public static JObject SmartParseParameters(JObject parameters)
        {
            if (parameters == null)
                return new JObject();

            // 检查是否是被包装的参数格式
            // 情况1: 参数被包装在"args"键中
            if (parameters.ContainsKey("args") && parameters["args"] is JObject argsObj
                && parameters.ContainsKey("server_name") && parameters.ContainsKey("tool_name"))
            {
                //Console.WriteLine("C#端: 检测到参数被包装在'args'中，自动解包...");
                return SmartParseParameters(argsObj);
            }

            // 情况2: 参数被包装在其他常见键中
            string[] wrapperKeys = { "parameters", "params", "arguments" };
            if (parameters.Count == 1)
            {
                var key = parameters.Properties().First().Name;
                if (wrapperKeys.Contains(key) && parameters[key] is JObject wrappedObj)
                {
                    //Console.WriteLine($"C#端: 检测到参数被包装在'{key}'中，自动解包...");
                    return wrappedObj;
                }
            }

            // 情况3: 参数是正确的格式，直接返回
            return parameters;
        }
    }

    /// <summary>
    /// MCP协议处理器
    /// </summary>
    public class McpHandler
    {
        private readonly ExcelManager _excelManager;

        public McpHandler(ExcelManager excelManager)
        {
            _excelManager = excelManager;
        }

        /// <summary>
        /// 处理MCP请求
        /// </summary>
        /// <param name="requestJson">请求JSON</param>
        /// <returns>响应JSON</returns>
        public string HandleRequest(string requestJson)
        {
            try
            {
                Console.WriteLine($"收到请求: {requestJson}");
                var request = JObject.Parse(requestJson);
                var method = request["method"]?.ToString();
                var parameters = request["params"] as JObject;

                // 智能解析参数
                var parsedParameters = ParameterHelper.SmartParseParameters(parameters);
                Console.WriteLine($"原始参数: {parameters}");
                Console.WriteLine($"解析后参数: {parsedParameters}");

                switch (method?.ToLower())
                {
                    case "execute_sql":
                        return HandleExecuteSql(parsedParameters);
                    case "get_tables":
                        return HandleGetTables();
                    case "get_create_table":
                        return HandleGetCreateTable(parsedParameters);
                    case "refresh":
                        return HandleRefresh();
                    case "change_directory":
                        return HandleChangeDirectory(parsedParameters);
                    default:
                        return CreateErrorResponse($"不支持的方法: {method}");
                }
            }
            catch (Exception ex)
            {
                return CreateErrorResponse($"处理请求时发生错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理执行SQL请求
        /// </summary>
        /// <param name="parameters">参数</param>
        /// <returns>响应JSON</returns>
        private string HandleExecuteSql(JObject parameters)
        {
            try
            {
                var sql = parameters["sql"]?.ToString();

                // 只在非MCP模式下输出调试信息
                if (!IsMcpMode())
                {
                    Console.WriteLine($"DEBUG: 收到的SQL语句: {sql}");
                }

                if (string.IsNullOrEmpty(sql))
                {
                    return CreateErrorResponse("SQL语句不能为空");
                }

                var statementType = SqlParser.ParseStatementType(sql);

                if (!IsMcpMode())
                {
                    Console.WriteLine($"DEBUG: 解析的SQL类型: {statementType}");
                }

                switch (statementType)
                {
                    case SqlParser.SqlStatementType.Select:
                        return HandleSelect(sql);
                    case SqlParser.SqlStatementType.Update:
                        return HandleUpdate(sql);
                    case SqlParser.SqlStatementType.Delete:
                        return HandleDelete(sql);
                    case SqlParser.SqlStatementType.ShowTables:
                        return HandleShowTables();
                    case SqlParser.SqlStatementType.ShowCreateTable:
                        return HandleShowCreateTable(sql);
                    default:
                        return CreateErrorResponse($"不支持的SQL语句类型: {statementType}");
                }
            }
            catch (Exception ex)
            {
                return CreateErrorResponse($"执行SQL时发生错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理SELECT语句
        /// </summary>
        /// <param name="sql">SELECT语句</param>
        /// <returns>响应JSON</returns>
        private string HandleSelect(string sql)
        {
            try
            {
                if (!IsMcpMode())
                {
                    Console.WriteLine($"DEBUG: 开始处理SELECT语句: {sql}");
                }

                var selectStatement = SqlParser.ParseSelect(sql);

                if (!IsMcpMode())
                {
                    Console.WriteLine($"DEBUG: 解析的表名: {selectStatement.TableName}");
                    Console.WriteLine($"DEBUG: 解析的列数: {selectStatement.Columns?.Count ?? 0}");
                }

                // 检查是否包含JOIN
                if (selectStatement.Joins != null && selectStatement.Joins.Count > 0)
                {
                    if (!IsMcpMode())
                    {
                        Console.WriteLine($"DEBUG: 执行JOIN查询，JOIN数量: {selectStatement.Joins.Count}");
                    }

                    var result = _excelManager.ExecuteSelectWithJoin(selectStatement);

                    if (!IsMcpMode())
                    {
                        Console.WriteLine($"DEBUG: JOIN查询结果行数: {result?.Count ?? 0}");
                    }

                    return CreateToolSuccessResponse(result);
                }
                else
                {
                    if (!IsMcpMode())
                    {
                        Console.WriteLine($"DEBUG: 执行普通查询");
                    }

                    var result = _excelManager.ExecuteSelectByFileName(
                        selectStatement.TableName,
                        selectStatement.Columns,
                        selectStatement.WhereClause,
                        selectStatement.Limit);

                    if (!IsMcpMode())
                    {
                        Console.WriteLine($"DEBUG: 普通查询结果行数: {result?.Count ?? 0}");
                    }

                    return CreateToolSuccessResponse(result);
                }
            }
            catch (Exception ex)
            {
                if (!IsMcpMode())
                {
                    Console.WriteLine($"DEBUG: SELECT语句执行异常: {ex.Message}");
                    Console.WriteLine($"DEBUG: 异常堆栈: {ex.StackTrace}");
                }

                return CreateToolErrorResponse($"执行SELECT语句时发生错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理UPDATE语句
        /// </summary>
        /// <param name="sql">UPDATE语句</param>
        /// <returns>响应JSON</returns>
        private string HandleUpdate(string sql)
        {
            try
            {
                var updateStatement = SqlParser.ParseUpdate(sql);
                var result = _excelManager.ExecuteUpdate(
                    updateStatement.TableName,
                    updateStatement.SetValues,
                    updateStatement.WhereClause);

                var response = new Dictionary<string, object>
                {
                    { "affectedRows", result },
                    { "message", $"成功更新 {result} 行数据" }
                };

                return CreateToolSuccessResponse(response);
            }
            catch (Exception ex)
            {
                return CreateToolErrorResponse($"执行UPDATE语句时发生错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理DELETE语句
        /// </summary>
        /// <param name="sql">DELETE语句</param>
        /// <returns>响应JSON</returns>
        private string HandleDelete(string sql)
        {
            try
            {
                var deleteStatement = SqlParser.ParseDelete(sql);
                var result = _excelManager.ExecuteDelete(
                    deleteStatement.TableName,
                    deleteStatement.WhereClause);

                var response = new Dictionary<string, object>
                {
                    { "affectedRows", result },
                    { "message", $"成功删除 {result} 行数据" }
                };

                return CreateToolSuccessResponse(response);
            }
            catch (Exception ex)
            {
                return CreateToolErrorResponse($"执行DELETE语句时发生错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理SHOW TABLES请求
        /// </summary>
        /// <returns>响应JSON</returns>
        private string HandleShowTables()
        {
            try
            {
                var tables = _excelManager.GetTableNames();
                return CreateToolSuccessResponse(tables);
            }
            catch (Exception ex)
            {
                return CreateToolErrorResponse($"获取表名时发生错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理SHOW CREATE TABLE请求
        /// </summary>
        /// <param name="sql">SHOW CREATE TABLE语句</param>
        /// <returns>响应JSON</returns>
        private string HandleShowCreateTable(string sql)
        {
            try
            {
                var tableName = SqlParser.ParseShowCreateTable(sql);
                var createTableStatement = _excelManager.GetCreateTableStatement(tableName);
                
                var result = new Dictionary<string, string>
                {
                    { "table", tableName },
                    { "createTable", createTableStatement }
                };
                
                return CreateToolSuccessResponse(result);
            }
            catch (Exception ex)
            {
                return CreateToolErrorResponse($"获取建表语句时发生错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理获取所有表名请求
        /// </summary>
        /// <returns>响应JSON</returns>
        private string HandleGetTables()
        {
            return HandleShowTables();
        }

        /// <summary>
        /// 处理获取建表语句请求
        /// </summary>
        /// <param name="parameters">参数</param>
        /// <returns>响应JSON</returns>
        private string HandleGetCreateTable(JObject parameters)
        {
            try
            {
                var tableName = parameters["table"]?.ToString();
                if (string.IsNullOrEmpty(tableName))
                {
                    return CreateToolErrorResponse("表名不能为空");
                }

                var createTableStatement = _excelManager.GetCreateTableStatement(tableName);
                
                var result = new Dictionary<string, string>
                {
                    { "table", tableName },
                    { "createTable", createTableStatement }
                };
                
                return CreateToolSuccessResponse(result);
            }
            catch (Exception ex)
            {
                return CreateToolErrorResponse($"获取建表语句时发生错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理刷新请求
        /// </summary>
        /// <returns>响应JSON</returns>
        private string HandleRefresh()
        {
            try
            {
                _excelManager.Refresh();
                return CreateToolSuccessResponse("缓存已刷新");
            }
            catch (Exception ex)
            {
                return CreateToolErrorResponse($"刷新缓存时发生错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理更改目录请求
        /// </summary>
        /// <param name="parameters">参数</param>
        /// <returns>响应JSON</returns>
        private string HandleChangeDirectory(JObject parameters)
        {
            try
            {
                var newDirectoryPath = parameters["directory"]?.ToString();
                if (string.IsNullOrEmpty(newDirectoryPath))
                {
                    return CreateToolErrorResponse("新目录路径不能为空");
                }

                var oldDirectoryPath = _excelManager.DirectoryPath;
                _excelManager.UpdateDirectoryPath(newDirectoryPath);

                var result = new Dictionary<string, string>
                {
                    { "old_directory", oldDirectoryPath },
                    { "new_directory", newDirectoryPath },
                    { "message", "目录已成功更改" }
                };

                return CreateToolSuccessResponse(result);
            }
            catch (Exception ex)
            {
                return CreateToolErrorResponse($"更改目录时发生错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 检查是否为MCP模式
        /// </summary>
        /// <returns>是否为MCP模式</returns>
        private bool IsMcpMode()
        {
            return Environment.GetCommandLineArgs().Contains("--mcp") ||
                   Environment.GetCommandLineArgs().Contains("-mcp");
        }

        /// <summary>
        /// 创建工具调用成功响应（符合MCP格式）
        /// </summary>
        /// <param name="result">结果数据</param>
        /// <returns>响应JSON</returns>

        private string CreateToolSuccessResponse(object result)
        {
            var response = new
            {
                meta = (object)null,
                content = new[] {
                    new {
                        type = "text",
                        text = JsonConvert.SerializeObject(result, Formatting.Indented),
                        annotations = (object)null,
                        meta = (object)null
                    }
                },
                structuredContent = (object)null,
                isError = false
            };

            return JsonConvert.SerializeObject(response, Formatting.Indented);
        }

        /// <summary>
        /// 创建工具调用错误响应（符合MCP格式）
        /// </summary>
        /// <param name="errorMessage">错误信息</param>
        /// <returns>响应JSON</returns>
        private string CreateToolErrorResponse(string errorMessage)
        {
            var response = new
            {
                meta = (object)null,
                content = new[] {
                    new {
                        type = "text",
                        text = errorMessage,
                        annotations = (object)null,
                        meta = (object)null
                    }
                },
                structuredContent = (object)null,
                isError = true
            };
            
            return JsonConvert.SerializeObject(response, Formatting.Indented);
        }

        /// <summary>
        /// 创建通用成功响应
        /// </summary>
        /// <param name="result">结果数据</param>
        /// <returns>响应JSON</returns>
        private string CreateSuccessResponse(object result)
        {
            var response = new
            {
                result = result
            };
            
            return JsonConvert.SerializeObject(response, Formatting.Indented);
        }

        /// <summary>
        /// 创建通用错误响应
        /// </summary>
        /// <param name="errorMessage">错误信息</param>
        /// <returns>响应JSON</returns>
        private string CreateErrorResponse(string errorMessage)
        {
            var response = new
            {
                error = new
                {
                    message = errorMessage
                }
            };
            
            return JsonConvert.SerializeObject(response, Formatting.Indented);
        }
    }
}