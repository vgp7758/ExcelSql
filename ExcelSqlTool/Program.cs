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
                // 设置控制台编码为UTF-8
                Console.OutputEncoding = Encoding.UTF8;
                Console.InputEncoding = Encoding.UTF8;

                // 检查命令行参数
                if (args.Length == 0)
                {
                    Console.WriteLine("用法:");
                    Console.WriteLine("  ExcelSqlTool.exe <目录路径>           # 传统模式");
                    Console.WriteLine("  ExcelSqlTool.exe <目录路径> --mcp    # MCP服务器模式");
                    Console.WriteLine("例如:");
                    Console.WriteLine("  ExcelSqlTool.exe ./XLSX");
                    Console.WriteLine("  ExcelSqlTool.exe ./XLSX --mcp");
                    return;
                }

                var directoryPath = "";
                bool isMcpMode = false;

                // 如果没有特殊参数，第一个参数就是目录路径
                if (args.Length > 0 && !args[0].StartsWith("-"))
                {
                    directoryPath = args[0];
                }

                // 解析其他参数
                for (int i = 0; i < args.Length; i++)
                {
                    var lowerArg = args[i].ToLower();
                    if (lowerArg == "--mcp" || lowerArg == "-mcp"|| lowerArg == "mcp")
                    {
                        isMcpMode = true;
                        continue;
                    }
                    if (lowerArg.StartsWith("-dir=") || lowerArg.StartsWith("--dir=")|| lowerArg.StartsWith("dir="))
                    {
                        directoryPath = args[i].Substring(lowerArg.IndexOf('=') + 1).Trim('"');
                        continue;
                    }
                }

                // 初始化Excel管理器
                var excelManager = new ExcelManager(directoryPath);

                if (isMcpMode)
                {
                    // MCP服务器模式 - 静默启动，避免输出干扰JSON通信
                    var mcpServer = new McpServer(excelManager, Console.OpenStandardInput(), Console.OpenStandardOutput());
                    
                    // 使用CancellationToken来处理Ctrl+C
                    var cts = new CancellationTokenSource();
                    Console.CancelKeyPress += (sender, e) =>
                    {
                        e.Cancel = true;
                        cts.Cancel();
                    };

                    try
                    {
                        await mcpServer.StartAsync();
                    }
                    catch (OperationCanceledException)
                    {
                        // 静默退出
                    }
                }
                else
                {
                    // 传统控制台模式
                    Console.WriteLine("Excel SQL工具已启动，等待MCP请求...");
                    Console.WriteLine("输入 'quit' 或 'exit' 退出程序");

                    // 初始化MCP处理器
                    var mcpHandler = new McpHandler(excelManager);

                    // 读取标准输入并处理MCP请求
                    string input;
                    while ((input = Console.ReadLine()) != null)
                    {
                        if (input.ToLower() == "quit" || input.ToLower() == "exit")
                        {
                            break;
                        }

                        if (!string.IsNullOrWhiteSpace(input))
                        {
                            try
                            {
                                var response = mcpHandler.HandleRequest(input);
                                Console.WriteLine(response);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"处理请求时发生错误: {ex.Message}");
                            }
                        }
                    }

                    Console.WriteLine("Excel SQL工具已退出");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"程序执行过程中发生错误: {ex.Message}");
                Console.WriteLine($"详细信息: {ex.StackTrace}");
            }
        }
    
    
    }
}