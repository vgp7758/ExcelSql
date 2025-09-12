#!/usr/bin/env python3
"""
Excel SQL Tool MCP Server
MCP服务器，将Excel SQL工具暴露给IDE使用
"""

import asyncio
import json
import subprocess
import sys
import os
from typing import Any, Dict, List, Optional
import logging

# 设置更详细的日志
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 添加当前目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 检查mcp模块是否存在
try:
    import mcp
    print(f"成功导入mcp模块，路径: {mcp.__file__}")
except ImportError as e:
    print(f"导入mcp模块失败: {e}")
    print("错误: 缺少必要的依赖包 'mcp'。请运行 'pip install mcp' 安装。")
    sys.exit(1)

# 导入mcp模块的组件
try:
    from mcp.server import Server
    from mcp.server.stdio import stdio_server
    from mcp.types import (
        CallToolRequest,
        CallToolResult,
        ListToolsRequest,
        Tool,
        TextContent,
        InitializeRequest,
        InitializeResult,
        ServerCapabilities
    )
    print("成功导入mcp模块的组件")
except ImportError as e:
    print(f"导入mcp模块组件失败: {e}")
    sys.exit(1)

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
        logger.info("检测到参数被包装在'args'中，自动解包...")
        return args["args"]
        
    # 情况2: 参数被包装在其他常见键中
    wrapper_keys = ["parameters", "params", "arguments"]
    if len(args) == 1:
        key = next(iter(args))
        if isinstance(args[key], dict) and key in wrapper_keys:
            logger.info(f"检测到参数被包装在'{key}'中，自动解包...")
            return args[key]
            
    # 情况3: 参数是正确的格式，直接返回
    return args

class ExcelSqlMcpServer:
    def __init__(self, excel_directory: str = "./XLSX"):
        self.excel_directory = excel_directory
        # 使用更灵活的方式确定Excel工具路径
        script_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 首先检查根目录是否有ExcelSqlTool.exe
        self.excel_tool_path = os.path.join(script_dir, "ExcelSqlTool.exe")
        
        # 如果根目录没有，检查标准构建路径
        if not os.path.exists(self.excel_tool_path):
            self.excel_tool_path = os.path.join(script_dir, "ExcelSqlTool", "bin", "Debug", "net48", "ExcelSqlTool.exe")
            
        # 如果还是没有，尝试在上级目录查找
        if not os.path.exists(self.excel_tool_path):
            parent_dir = os.path.dirname(script_dir)
            self.excel_tool_path = os.path.join(parent_dir, "ExcelSqlTool", "bin", "Debug", "net48", "ExcelSqlTool.exe")
            
        # 最后检查是否存在
        if not os.path.exists(self.excel_tool_path):
            print(f"警告: Excel工具未找到: {self.excel_tool_path}")
            print("请确保已构建Excel SQL工具项目")
        
        self.server = Server("excel-sql-tool")
        
        # 注册处理程序
        self.server.list_tools()(self.list_tools)
        self.server.call_tool()(self.call_tool)
        
    async def list_tools(self) -> List[Tool]:
        """返回可用工具列表"""
        logger.info("列出可用工具")
        tools = [
            Tool(
                name="excel_show_tables",
                description="显示Excel中所有可用的表名（这些名称在SQL查询中用作表名）",
                inputSchema={
                    "type": "object",
                    "properties": {},
                    "required": []
                }
            ),
            Tool(
                name="excel_query",
                description="执行SQL查询Excel数据，表名应为工作表名称而非文件名",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "sql": {
                            "type": "string",
                            "description": "SQL查询语句，支持SELECT、SHOW TABLES、SHOW CREATE TABLE等。注意：表名应为工作表名称"
                        }
                    },
                    "required": ["sql"]
                }
            ),
            Tool(
                name="excel_get_table_schema",
                description="获取指定表的结构定义，表名应为工作表名称而非文件名",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "table_name": {
                            "type": "string",
                            "description": "表名（应为工作表名称，不是Excel文件名）"
                        }
                    },
                    "required": ["table_name"]
                }
            ),
            Tool(
                name="excel_refresh_cache",
                description="刷新Excel文件缓存，重新加载所有文件",
                inputSchema={
                    "type": "object",
                    "properties": {},
                    "required": []
                }
            ),
            Tool(
                name="excel_list_sheets",
                description="列出所有Excel工作表",
                inputSchema={
                    "type": "object",
                    "properties": {},
                    "required": []
                }
            )
        ]
        logger.info(f"返回 {len(tools)} 个工具")
        return tools
    
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
            elif name == "excel_refresh_cache":
                return await self._refresh_cache(parsed_arguments.get("directory"))
            elif name == "excel_list_sheets":
                return await self._get_tables(parsed_arguments.get("directory"))
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
    
    async def _execute_sql(self, sql: str, directory: str = None) -> CallToolResult:
        """执行SQL语句"""
        try:
            request = {
                "method": "execute_sql",
                "params": {"sql": sql}
            }
            
            result = await self._send_request_to_excel_tool(request, directory)
            return self._safe_create_call_tool_result(result)
        except Exception as e:
            return self._safe_create_call_tool_result({
                "content": [{"type": "text", "text": f"执行SQL失败: {str(e)}"}],
                "isError": True
            })
    
    async def _get_create_table(self, table_name: str, directory: str = None) -> CallToolResult:
        """获取表结构"""
        try:
            request = {
                "method": "get_create_table",
                "params": {"table": table_name}
            }
            
            result = await self._send_request_to_excel_tool(request, directory)
            return self._safe_create_call_tool_result(result)
        except Exception as e:
            return self._safe_create_call_tool_result({
                "content": [{"type": "text", "text": f"获取表结构失败: {str(e)}"}],
                "isError": True
            })
    
    async def _get_tables(self, directory: str = None) -> CallToolResult:
        """获取所有表"""
        try:
            request = {
                "method": "get_tables",
                "params": {}
            }
            
            result = await self._send_request_to_excel_tool(request, directory)
            return self._safe_create_call_tool_result(result)
        except Exception as e:
            return self._safe_create_call_tool_result({
                "content": [{"type": "text", "text": f"获取表列表失败: {str(e)}"}],
                "isError": True
            })
    
    async def _refresh_cache(self, directory: str = None) -> CallToolResult:
        """刷新缓存"""
        try:
            request = {
                "method": "refresh",
                "params": {}
            }
            
            result = await self._send_request_to_excel_tool(request, directory)
            return self._safe_create_call_tool_result(result)
        except Exception as e:
            return self._safe_create_call_tool_result({
                "content": [{"type": "text", "text": f"刷新缓存失败: {str(e)}"}],
                "isError": True
            })
    
    async def _send_request_to_excel_tool(self, request: Dict[str, Any], directory: str = None) -> Dict[str, Any]:
        """发送请求到Excel工具"""
        try:
            # 检查Excel工具是否存在
            if not os.path.exists(self.excel_tool_path):
                raise Exception(f"Excel工具未找到: {self.excel_tool_path}")
            
            # 启动Excel工具进程
            logger.info(f"启动Excel工具进程: {self.excel_tool_path} {directory or self.excel_directory}")
            process = await asyncio.create_subprocess_exec(
                self.excel_tool_path,
                directory or self.excel_directory,
                stdin=asyncio.subprocess.PIPE,
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE
            )
            
            # 发送请求
            request_json = json.dumps(request, ensure_ascii=False)
            logger.info(f"发送请求到Excel工具: {request_json}")
            
            # 发送请求并读取响应
            try:
                stdout, stderr = await asyncio.wait_for(
                    process.communicate(input=(request_json + '\nquit\n').encode('utf-8')),
                    timeout=30
                )
            except asyncio.TimeoutError:
                process.terminate()
                raise
            
            if stderr:
                error_text = stderr.decode('utf-8', errors='ignore')
                logger.error(f"Excel工具错误: {error_text}")
                # 如果是编码错误，尝试使用其他编码解码
                if 'codec can\'t encode' in error_text or 'illegal multibyte sequence' in error_text:
                    try:
                        error_text = stderr.decode('gbk', errors='ignore')
                        logger.error(f"Excel工具错误 (GBK解码): {error_text}")
                    except:
                        pass
            
            # 解析响应
            response_text = stdout.decode('utf-8', errors='ignore')
            logger.info(f"Excel工具响应文本: {response_text}")
            
            # 查找JSON响应 - 更严格的解析，支持多行JSON
            lines = response_text.split('\n')
            json_lines = []
            in_json = False
            brace_count = 0
            
            for line in lines:
                line = line.strip()
                # 开始JSON对象
                if line.startswith('{') and not in_json:
                    in_json = True
                    json_lines = [line]
                    brace_count = line.count('{') - line.count('}')
                # JSON对象的一部分
                elif in_json:
                    json_lines.append(line)
                    brace_count += line.count('{') - line.count('}')
                    
                    # JSON对象结束（大括号平衡）
                    if brace_count == 0:
                        json_text = '\n'.join(json_lines)
                        try:
                            parsed_response = json.loads(json_text)
                            logger.info(f"成功解析JSON响应: {json.dumps(parsed_response, indent=2, ensure_ascii=False)}")
                            return parsed_response
                        except json.JSONDecodeError as e:
                            logger.warning(f"JSON解析失败: {e}, JSON文本: {json_text[:200]}...")
                            in_json = False
                            json_lines = []
                            brace_count = 0
                            continue
            
            # 如果没有找到JSON响应，返回原始响应
            if response_text.strip():
                logger.warning(f"未找到有效的JSON响应，返回原始响应: {response_text}")
                return {"raw_response": response_text}
            
            raise Exception("无法解析Excel工具响应")
            
        except asyncio.TimeoutError:
            raise Exception("Excel工具响应超时")
        except Exception as e:
            logger.error(f"调用Excel工具失败: {str(e)}")
            raise Exception(f"调用Excel工具失败: {str(e)}")

    def _safe_create_call_tool_result(self, response_data: Dict[str, Any]) -> CallToolResult:
        """安全地创建CallToolResult对象，处理可能的格式错误"""
        try:
            logger.info(f"创建CallToolResult，输入数据: {json.dumps(response_data, indent=2, ensure_ascii=False)}")
            
            # 检查是否是C#工具的标准响应格式 {"result": ...}
            if "result" in response_data:
                logger.info("检测到C#工具的标准响应格式")
                result_data = response_data["result"]
                # 将结果转换为JSON字符串显示
                if isinstance(result_data, (list, dict)):
                    text_content = json.dumps(result_data, indent=2, ensure_ascii=False)
                else:
                    text_content = str(result_data)
                return CallToolResult(
                    content=[TextContent(type="text", text=text_content)],
                    isError=False
                )
            
            # 检查是否是错误格式
            if "error" in response_data:
                error_msg = response_data["error"].get("message", "未知错误")
                logger.info(f"检测到错误格式，错误信息: {error_msg}")
                return CallToolResult(
                    content=[TextContent(type="text", text=error_msg)],
                    isError=True
                )
            
            # 检查是否是原始响应格式
            if "raw_response" in response_data:
                logger.info("检测到原始响应格式")
                return CallToolResult(
                    content=[TextContent(type="text", text=response_data["raw_response"])],
                    isError=True
                )
            
            # 检查是否是元组格式（这是MCP库序列化导致的问题）
            content = response_data.get("content", [])
            if isinstance(content, list) and len(content) > 0 and isinstance(content[0], (list, tuple)):
                logger.warning("检测到元组格式的content数据，尝试修复")
                return self._fix_tuple_format_response(response_data)
            
            # 检查是否已经是正确的MCP响应格式
            if "content" in response_data and isinstance(response_data["content"], list):
                logger.info("检测到正确的MCP响应格式")
                content = response_data["content"]
                is_error = response_data.get("isError", False)
                
                # 转换content中的字典为TextContent对象
                corrected_content = []
                for item in content:
                    if isinstance(item, dict) and "type" in item:
                        corrected_content.append(TextContent(**item))
                    elif isinstance(item, str):
                        corrected_content.append(TextContent(type="text", text=item))
                    else:
                        corrected_content.append(TextContent(type="text", text=str(item)))
                
                return CallToolResult(
                    content=corrected_content,
                    isError=is_error,
                    meta=response_data.get("meta"),
                    structuredContent=response_data.get("structuredContent")
                )
            
            # 如果没有识别的格式，将整个响应作为文本返回
            logger.warning("无法识别响应格式，将整个响应作为文本返回")
            text_content = json.dumps(response_data, indent=2, ensure_ascii=False)
            return CallToolResult(
                content=[TextContent(type="text", text=text_content)],
                isError=False
            )
                
        except Exception as e:
            logger.error(f"创建CallToolResult时发生错误: {e}")
            # 返回错误信息
            return CallToolResult(
                content=[TextContent(type="text", text=f"创建CallToolResult失败: {str(e)}")],
                isError=True
            )
    
    def _fix_tuple_format_response(self, response_data: Dict[str, Any]) -> CallToolResult:
        """修复元组格式的响应数据"""
        logger.info(f"修复元组格式的响应数据: {response_data}")
        
        try:
            # 提取字段值
            meta = None
            content = []
            structured_content = None
            is_error = False
            
            # 遍历元组格式的数据
            for item in response_data.get("content", []):
                if isinstance(item, (list, tuple)) and len(item) == 2:
                    key, value = item
                    if key == "meta":
                        meta = value
                    elif key == "content":
                        content = value
                    elif key == "structuredContent":
                        structured_content = value
                    elif key == "isError":
                        is_error = value
            
            # 确保content是正确的格式
            corrected_content = []
            for item in content:
                if isinstance(item, dict) and "type" in item:
                    corrected_content.append(TextContent(**item))
                elif isinstance(item, str):
                    corrected_content.append(TextContent(type="text", text=item))
                else:
                    corrected_content.append(TextContent(type="text", text=str(item)))
            
            logger.info("成功修复元组格式数据")
            return CallToolResult(
                content=corrected_content,
                isError=is_error,
                meta=meta,
                structuredContent=structured_content
            )
            
        except Exception as e:
            logger.error(f"修复元组格式数据时发生错误: {e}")
            # 返回错误信息
            return CallToolResult(
                content=[TextContent(type="text", text=f"修复数据格式失败: {str(e)}")],
                isError=True
            )
  
async def main():
    """主函数"""
    # 从命令行参数获取Excel目录
    excel_directory = "./XLSX"
    if len(sys.argv) > 1:
        excel_directory = sys.argv[1]
    
    # 创建MCP服务器
    server_instance = ExcelSqlMcpServer(excel_directory)
    
    # 打印注册的处理程序
    print("注册的请求处理程序:")
    for req_type in server_instance.server.request_handlers:
        print(f"  {req_type}")
    
    # 使用stdio_server启动服务器
    async with stdio_server() as (read_stream, write_stream):
        # 创建初始化选项
        initialization_options = server_instance.server.create_initialization_options()
        # 运行服务器
        await server_instance.server.run(read_stream, write_stream, initialization_options)

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("MCP服务器已停止")
    except Exception as e:
        print(f"MCP服务器错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)