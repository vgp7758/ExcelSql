using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelSqlTool
{
    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                // UTF-8 无 BOM，避免向 stdout 输出非协议字符
                Console.OutputEncoding = new UTF8Encoding(false);
                Console.InputEncoding = new UTF8Encoding(false);

                var directoryPath = "d:/Projects/BunkerProject/TableTools/XLSX";
                bool isMcpMode = false;

                // 先解析开关，再解析目录，避免把"mcp"当作目录
                foreach (var a in args)
                {
                    var lower = a.ToLowerInvariant();
                    if (lower == "--mcp" || lower == "-mcp" || lower == "mcp")
                    {
                        isMcpMode = true;
                        continue;
                    }
                    if (lower.StartsWith("-dir=") || lower.StartsWith("--dir=") || lower.StartsWith("dir=")||lower.StartsWith("--cwd"))
                    {
                        directoryPath = a.Substring(a.IndexOf('=') + 1).Trim('"');
                        continue;
                    }
                }
                // 兜底：允许把第一个非开关且不含'='的参数当目录
                if (string.IsNullOrEmpty(directoryPath))
                {
                    foreach (var a in args)
                    {
                        var lower = a.ToLowerInvariant();
                        if (!lower.StartsWith("-") && !lower.Contains("=") && lower != "mcp")
                        {
                            directoryPath = a;
                            break;
                        }
                    }
                }

                var excelManager = new ExcelManager(directoryPath);

                if (isMcpMode)
                {
                    // MCP服务器模式：仅输出JSON-RPC
                    var mcpServer = new McpServer(excelManager, Console.OpenStandardInput(), Console.OpenStandardOutput());
                    var cts = new CancellationTokenSource();
                    Console.CancelKeyPress += (s, e) => { e.Cancel = true; cts.Cancel(); };
                    try { await mcpServer.StartAsync(); } catch (OperationCanceledException) { }
                    return;
                }

                // 传统模式（调试用）
                Console.WriteLine("Excel SQL工具已启动，等待MCP请求...");
                Console.WriteLine("输入 'quit' 或 'exit' 退出程序");
                var handler = new McpHandler(excelManager);
                string input;
                while ((input = Console.ReadLine()) != null)
                {
                    if (input.Equals("quit", StringComparison.OrdinalIgnoreCase) || input.Equals("exit", StringComparison.OrdinalIgnoreCase))
                        break;
                    if (string.IsNullOrWhiteSpace(input)) continue;
                    try { Console.WriteLine(handler.HandleRequest(input)); }
                    catch (Exception ex) { Console.WriteLine($"处理请求时发生错误: {ex.Message}"); }
                }
            }
            catch (Exception ex)
            {
                // 避免在stdout写异常
                try { Console.Error.WriteLine(ex.Message); } catch { }
            }
        }
    }
}