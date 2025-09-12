#!/usr/bin/env python3
"""
MCP参数处理解决方案
解决IDE agent错误包装参数的问题
"""

import json
from typing import Dict, Any

def analyze_parameter_issue():
    """分析参数问题"""
    print("参数问题分析:")
    print("=" * 50)
    print("问题描述:")
    print("IDE中的agent有时会错误地将参数包装在一个额外的JSON结构中")
    print("而不是直接发送参数。")
    print()
    print("示例:")
    print("正确格式: {\"directory\": \"path/to/directory\"}")
    print("错误格式: {\"args\": {\"directory\": \"path/to/directory\"}}")
    print()

def demonstrate_parameter_handling():
    """演示参数处理"""
    print("参数处理演示:")
    print("=" * 50)
    
    # 正确的参数格式
    correct_params = {
        "directory": "d:\\Projects\\Bunker\\TableTools\\XLSX"
    }
    
    # 错误的参数格式（被包装在args中）
    incorrect_params = {
        "args": {
            "directory": "d:\\Projects\\Bunker\\TableTools\\XLSX"
        }
    }
    
    print("正确参数格式:")
    print(json.dumps(correct_params, ensure_ascii=False, indent=2))
    print()
    
    print("错误参数格式:")
    print(json.dumps(incorrect_params, ensure_ascii=False, indent=2))
    print()

def solution_approach():
    """解决方案方法"""
    print("解决方案方法:")
    print("=" * 50)
    print("1. 在服务器端实现智能参数解析")
    print("2. 检查参数是否被错误包装")
    print("3. 如果发现包装，自动解包参数")
    print("4. 保持向后兼容性")
    print()

def implement_parameter_parser():
    """实现参数解析器"""
    print("实现参数解析器:")
    print("=" * 50)
    
    def parse_tool_arguments(args: Dict[str, Any]) -> Dict[str, Any]:
        """
        智能解析工具参数，处理可能的包装问题
        
        Args:
            args: 原始参数字典
            
        Returns:
            解析后的参数字典
        """
        # 如果参数是空的，直接返回
        if not args:
            return args
            
        # 检查是否是被包装的参数格式
        # 情况1: 参数被包装在"args"键中
        if "args" in args and isinstance(args["args"], dict):
            print("检测到参数被包装在'args'中，自动解包...")
            return args["args"]
            
        # 情况2: 参数被包装在其他键中（如"parameters"）
        if len(args) == 1:
            key = next(iter(args))
            if isinstance(args[key], dict) and key in ["parameters", "params", "arguments"]:
                print(f"检测到参数被包装在'{key}'中，自动解包...")
                return args[key]
                
        # 情况3: 参数是正确的格式，直接返回
        return args
    
    # 测试解析器
    test_cases = [
        # 正确格式
        {
            "name": "正确格式",
            "params": {"directory": "d:\\Projects\\Bunker\\TableTools\\XLSX"}
        },
        # 包装在args中
        {
            "name": "包装在args中",
            "params": {"args": {"directory": "d:\\Projects\\Bunker\\TableTools\\XLSX"}}
        },
        # 包装在parameters中
        {
            "name": "包装在parameters中",
            "params": {"parameters": {"directory": "d:\\Projects\\Bunker\\TableTools\\XLSX"}}
        },
        # 包装在params中
        {
            "name": "包装在params中",
            "params": {"params": {"directory": "d:\\Projects\\Bunker\\TableTools\\XLSX"}}
        },
        # 包装在arguments中
        {
            "name": "包装在arguments中",
            "params": {"arguments": {"directory": "d:\\Projects\\Bunker\\TableTools\\XLSX"}}
        },
        # 空参数
        {
            "name": "空参数",
            "params": {}
        }
    ]
    
    print("测试参数解析器:")
    for test_case in test_cases:
        print(f"\n测试: {test_case['name']}")
        print(f"输入: {json.dumps(test_case['params'], ensure_ascii=False)}")
        result = parse_tool_arguments(test_case['params'])
        print(f"输出: {json.dumps(result, ensure_ascii=False)}")

def implement_in_fastmcp_server():
    """在FastMCP服务器中实现解决方案"""
    print("\n\n在FastMCP服务器中实现解决方案:")
    print("=" * 50)
    
    fastmcp_solution = '''
import json
from typing import Dict, Any

def smart_parse_arguments(args: Dict[str, Any]) -> Dict[str, Any]:
    """
    智能解析工具参数，处理IDE agent可能的参数包装问题
    
    Args:
        args: 原始参数字典
        
    Returns:
        解析后的参数字典
    """
    # 如果参数为空，直接返回
    if not args:
        return args
        
    # 检查是否是被包装的参数格式
    # 情况1: 参数被包装在"args"键中
    if "args" in args and isinstance(args["args"], dict):
        print("FastMCP服务器: 检测到参数被包装在'args'中，自动解包...")
        return args["args"]
        
    # 情况2: 参数被包装在其他常见键中
    wrapper_keys = ["parameters", "params", "arguments"]
    if len(args) == 1:
        key = next(iter(args))
        if isinstance(args[key], dict) and key in wrapper_keys:
            print(f"FastMCP服务器: 检测到参数被包装在'{key}'中，自动解包...")
            return args[key]
            
    # 情况3: 参数是正确的格式，直接返回
    return args

# 在FastMCP工具函数中使用
@mcp.tool()
def excel_show_tables(directory: str = None) -> str:
    """显示Excel中所有可用的表名"""
    try:
        # 注意：在FastMCP中，参数会自动解析，但我们可以在工具内部再次检查
        # 这里只是示例，实际FastMCP会自动处理参数
        print(f"excel_show_tables接收到参数: directory={directory}")
        # 实际实现...
        return "表名列表"
    except Exception as e:
        return f"错误: {str(e)}"

# 在标准MCP服务器中使用
class ExcelSqlMcpServer:
    async def call_tool(self, name: str, arguments: Dict[str, Any]) -> CallToolResult:
        """调用指定的工具"""
        logger.info(f"调用工具: {name}，原始参数: {arguments}")
        try:
            # 智能解析参数
            parsed_arguments = smart_parse_arguments(arguments)
            logger.info(f"解析后参数: {parsed_arguments}")
            
            if name == "excel_show_tables":
                return await self._get_tables(parsed_arguments.get("directory"))
            elif name == "excel_query":
                sql = parsed_arguments.get("sql")
                if not sql:
                    raise ValueError("SQL查询语句不能为空")
                return await self._execute_sql(sql, parsed_arguments.get("directory"))
            elif name == "excel_get_table_schema":
                table_name = parsed_arguments.get("table_name")
                if not table_name:
                    raise ValueError("表名不能为空")
                return await self._get_create_table(table_name, parsed_arguments.get("directory"))
            # ... 其他工具
            else:
                raise ValueError(f"未知工具: {name}")
                
        except Exception as e:
            logger.error(f"工具调用失败: {e}")
            return CallToolResult(
                content=[TextContent(
                    type="text",
                    text=f"错误: {str(e)}"
                )],
                isError=True
            )
    '''
    
    print("FastMCP服务器中的参数处理实现:")
    print(fastmcp_solution)

def implement_in_csharp():
    """在C#端实现解决方案"""
    print("\n在C#端实现解决方案:")
    print("=" * 50)
    
    csharp_solution = '''
using System;
using System.Collections.Generic;
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
            if (parameters.ContainsKey("args") && parameters["args"] is JObject argsObj)
            {
                Console.WriteLine("C#端: 检测到参数被包装在'args'中，自动解包...");
                return argsObj;
            }

            // 情况2: 参数被包装在其他常见键中
            string[] wrapperKeys = { "parameters", "params", "arguments" };
            if (parameters.Count == 1)
            {
                var key = parameters.Properties().First().Name;
                if (Array.IndexOf(wrapperKeys, key) >= 0 && parameters[key] is JObject wrappedObj)
                {
                    Console.WriteLine($"C#端: 检测到参数被包装在'{key}'中，自动解包...");
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
                if (string.IsNullOrEmpty(sql))
                {
                    return CreateErrorResponse("SQL语句不能为空");
                }

                var statementType = SqlParser.ParseStatementType(sql);

                switch (statementType)
                {
                    case SqlParser.SqlStatementType.Select:
                        return HandleSelect(sql);
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

        // ... 其他方法保持不变
    }
}
    '''
    
    print("C#端的参数处理实现:")
    print(csharp_solution)

def main():
    """主函数"""
    print("MCP参数处理解决方案")
    print("=" * 60)
    
    analyze_parameter_issue()
    demonstrate_parameter_handling()
    solution_approach()
    implement_parameter_parser()
    implement_in_fastmcp_server()
    implement_in_csharp()
    
    print("\n\n总结:")
    print("=" * 60)
    print("通过以上解决方案，我们可以有效处理IDE agent错误包装参数的问题:")
    print("1. 在服务器端实现智能参数解析器")
    print("2. 自动检测和解包被包装的参数")
    print("3. 保持向后兼容性，支持多种包装格式")
    print("4. 在FastMCP和标准MCP服务器中都实现该功能")
    print("5. 在C#端也实现相应的参数处理逻辑")

if __name__ == "__main__":
    main()