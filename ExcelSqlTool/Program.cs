using System;
using System.IO;
using System.Text;

namespace ExcelSqlTool
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 设置控制台编码为UTF-8
                Console.OutputEncoding = Encoding.UTF8;
                Console.InputEncoding = Encoding.UTF8;
                
                // 检查命令行参数
                if (args.Length == 0)
                {
                    Console.WriteLine("用法: ExcelSqlTool.exe <目录路径>");
                    Console.WriteLine("例如: ExcelSqlTool.exe ./XLSX");
                    return;
                }

                var directoryPath = args[0];
                
                // 检查目录是否存在
                if (!Directory.Exists(directoryPath))
                {
                    // 尝试转换为绝对路径
                    directoryPath = Path.GetFullPath(directoryPath);
                    if (!Directory.Exists(directoryPath))
                    {
                        Console.WriteLine($"错误: 目录 {directoryPath} 不存在");
                        return;
                    }
                }

                Console.WriteLine($"Excel SQL工具启动中...");
                Console.WriteLine($"监控目录: {directoryPath}");

                // 初始化Excel管理器
                var excelManager = new ExcelManager(directoryPath);
                
                // 初始化MCP处理器
                var mcpHandler = new McpHandler(excelManager);

                Console.WriteLine("Excel SQL工具已启动，等待MCP请求...");
                Console.WriteLine("输入 'quit' 或 'exit' 退出程序");

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
            catch (Exception ex)
            {
                Console.WriteLine($"程序执行过程中发生错误: {ex.Message}");
                Console.WriteLine($"详细信息: {ex.StackTrace}");
            }
        }
    }
}