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
            AddTool(new ShowTablesTool(_excelManager));
            AddTool(new ListSheetsTool(_excelManager)); // 与show_tables功能重复，但更直观
            AddTool(new ListFilesTool(_excelManager));
            AddTool(new GetDirectoryTool(_excelManager));
            AddTool(new GetStatsTool(_excelManager));
            AddTool(new UndoChangesTool(_excelManager));

            AddTool(new QueryTool(_excelManager));
            AddTool(new GetTableSchemaTool(_excelManager));
            AddTool(new RefreshCacheTool(_excelManager));
            AddTool(new ChangeDirectoryTool(_excelManager));
            AddTool(new SaveAllTool(_excelManager));
            AddTool(new SaveFileTool(_excelManager));
        }

        private void AddTool(ToolBase tool)
        {
            if (tool == null || string.IsNullOrEmpty(tool.name)) return;
            _tools[tool.name] = tool;
        }

        private string RemoveBOM(string str)
        {
            if (string.IsNullOrEmpty(str)) return str;
            if (str.Length > 0 && str[0] == '\uFEFF') return str.Substring(1);
            return str;
        }

        public async Task StartAsync()
        {
            _isRunning = true;
            try
            {
                while (_isRunning)
                {
                    try
                    {
                        var requestLine = await _input.ReadLineAsync();
                        if (string.IsNullOrEmpty(requestLine))
                        {
                            await Task.Delay(50);
                            continue;
                        }
                        if(requestLine == "quit" || requestLine == "exit" || requestLine == "shutdown")
                        {
                            requestLine = JsonConvert.SerializeObject(new
                            {
                                jsonrpc = "2.0",
                                method = requestLine,
                                id = Guid.NewGuid().ToString()
                            });
                        }
                        requestLine = RemoveBOM(requestLine);
                        var request = JObject.Parse(requestLine);
                        var response = await ProcessRequestAsync(request);
                        var responseJson = JsonConvert.SerializeObject(response, Formatting.None);
                        responseJson = RemoveBOM(responseJson);
                        await _output.WriteLineAsync(responseJson);
                        await _output.FlushAsync();
                    }
                    catch (Exception ex)
                    {
                        var errorResponse = new
                        {
                            jsonrpc = "2.0",
                            id = (object)null,
                            error = new { code = -32603, message = "Internal error", data = ex.Message }
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
                // 静默
            }
        }

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
                        return new { jsonrpc = "2.0", id, result = new { ok = true } };
                    case "exit":
                    case "quit":
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
                            error = new { code = -32601, message = "Method not found", data = $"未知方法: {method}" }
                        };
                }
            }
            catch (Exception ex)
            {
                return new { jsonrpc = "2.0", id, error = new { code = -32603, message = "Internal error", data = ex.Message } };
            }
        }

        private Task<object> HandleInitializeAsync(JObject request, object id)
        {
            var response = new
            {
                jsonrpc = "2.0",
                id,
                result = new
                {
                    protocolVersion = "2024-11-05",
                    capabilities = new { tools = new { listChanged = false } },
                    serverInfo = new { name = "excel-sql", version = "1.0.0" }
                }
            };
            return Task.FromResult<object>(response);
        }

        private Task<object> HandleShutdownAsync(object id)
        {
            _isRunning = false;
            return Task.FromResult<object>(new { jsonrpc = "2.0", id, result = (object)null });
        }

        private Task<object> HandleListToolsAsync(object id)
        {
            var tools = new List<object>();
            foreach (var kv in _tools) tools.Add(((ToolBase)kv.Value).Info);
            return Task.FromResult<object>(new { jsonrpc = "2.0", id, result = new { tools = tools.ToArray() } });
        }

        private async Task<object> HandleCallToolAsync(JObject request, object id)
        {
            var name = request["params"]?["name"]?.ToString();
            var arguments = request["params"]?["arguments"] as JObject;
            try
            {
                var parsedArguments = ParameterHelper.SmartParseParameters(arguments);
                if (string.IsNullOrEmpty(name) || !_tools.ContainsKey(name))
                    throw new ArgumentException($"未知工具: {name}");
                if (!string.Equals(name, "excel_change_directory", StringComparison.OrdinalIgnoreCase) && !_excelManager.IsDirectoryExists)
                {
                    return new { jsonrpc = "2.0", id, error = new { code = -32000, message = "Excel文件目录不存在或未设置", data = "请先使用'excel_change_directory'工具设置有效的Excel文件目录" } };
                }
                var tool = _tools[name];
                var result = await tool.CallAsync(parsedArguments ?? new JObject());
                return new { jsonrpc = "2.0", id, result = new { content = new[] { new { type = "text", text = JsonConvert.SerializeObject(result, Formatting.Indented) } } } };
            }
            catch (Exception ex)
            {
                return new { jsonrpc = "2.0", id, error = new { code = -32603, message = "Internal error", data = ex.Message } };
            }
        }
    }
}