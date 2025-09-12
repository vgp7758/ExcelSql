#!/usr/bin/env python3
"""
Excel SQL Tool FastMCP Server
使用FastMCP框架实现的MCP服务器，将Excel SQL工具暴露给IDE使用
"""

import asyncio
import json
import subprocess
import sys
import os
from typing import Any, Dict, List, Optional
import logging
from pathlib import Path
import concurrent.futures

# 添加当前目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 导入FastMCP
try:
    from fastmcp import FastMCP
    from fastmcp.server.middleware import Middleware, MiddlewareContext
    from mcp.types import (
        TextContent,
        CallToolResult,
    )
    print("成功导入FastMCP和MCP模块")
except ImportError as e:
    print(f"导入FastMCP或MCP模块失败: {e}")
    print("请运行 'pip install fastmcp' 安装FastMCP")
    sys.exit(1)

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 创建 MCP Server
mcp = FastMCP("excel-sql-tool")

# 默认Excel目录
default_excel_directory = "./XLSX"

class NonStandardRequestMiddleware(Middleware):
    """处理非标准请求格式的中间件"""
    
    async def on_request(self, context: MiddlewareContext, call_next):
        """处理请求消息"""
        logger.info(f"处理请求: {context.method}")
        
        # 检查是否是工具调用请求
        if context.method == "tools/call" and hasattr(context, 'message') and hasattr(context.message, 'params'):
            logger.info(f"原始参数: {context.message.params}")
            
            # 智能解析工具调用参数
            parsed_params = smart_parse_tool_call_params(context.message.params)
            logger.info(f"解析后参数: {parsed_params}")
            
            # 直接修改context.message.params
            context.message.params = parsed_params
            
            # 同时确保name和arguments字段存在
            if "name" not in context.message.params:
                context.message.params["name"] = ""
            if "arguments" not in context.message.params:
                context.message.params["arguments"] = {}
        
        return await call_next(context)

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

def smart_parse_tool_call_params(params: Dict[str, Any]) -> Dict[str, Any]:
    """
    智能解析工具调用参数，处理IDE agent可能的非标准格式
    
    Args:
        params: 原始工具调用参数字典
        
    Returns:
        解析后的标准格式参数字典，包含'name'和'arguments'键
    """
    logger.info(f"解析工具调用参数: {params}")
    
    # 如果参数为空，返回默认值
    if not params:
        logger.warning("工具调用参数为空")
        return {"name": "", "arguments": {}}
        
    # 检查是否是标准MCP格式: {name: "...", arguments: {...}}
    if "name" in params and "arguments" in params:
        logger.info("检测到标准MCP工具调用格式")
        return params
        
    # 检查是否是非标准格式: {tool_name: "...", args: {...}}
    if "tool_name" in params and "args" in params:
        logger.info("检测到非标准工具调用格式，自动转换为标准格式...")
        result = {
            "name": params["tool_name"],
            "arguments": params["args"] if isinstance(params["args"], dict) else {}
        }
        logger.info(f"转换结果: {result}")
        return result
        
    # 检查是否是另一种非标准格式: {server_name: "...", tool_name: "...", args: {...}}
    if "server_name" in params and "tool_name" in params and "args" in params:
        logger.info("检测到另一种非标准工具调用格式，自动转换为标准格式...")
        result = {
            "name": params["tool_name"],
            "arguments": params["args"] if isinstance(params["args"], dict) else {}
        }
        logger.info(f"转换结果: {result}")
        return result
        
    # 如果无法识别格式，返回原始参数作为arguments
    logger.warning(f"无法识别工具调用参数格式，将整个参数作为arguments: {params}")
    return {"name": "", "arguments": params}

@mcp.tool
def set_excel_directory(directory: str = None) -> str:  # pyright: ignore[reportArgumentType]
    """设置Excel工作目录，请求参数不需要包装成包含server_name和tool_name的结构，而是直接传递啊Args
    """
    global default_excel_directory
    try:
        # 智能解析参数
        parsed_args = smart_parse_arguments({"directory": directory} if directory is not None else {})
        actual_directory = parsed_args.get("directory", directory)
        
        # 如果没有提供目录参数，返回当前目录
        if actual_directory is None:
            return f"当前Excel工作目录: {default_excel_directory}"
        
        # 验证目录是否存在
        dir_path = Path(actual_directory)
        if not dir_path.exists():
            return f"错误: 目录 '{actual_directory}' 不存在"
        if not dir_path.is_dir():
            return f"错误: '{actual_directory}' 不是一个有效的目录"
        
        default_excel_directory = actual_directory
        return f"Excel工作目录已设置为: {actual_directory}"
    except Exception as e:
        return f"设置Excel工作目录失败: {str(e)}"

@mcp.tool
def get_excel_directory() -> str:
    """获取当前Excel工作目录"""
    return default_excel_directory

@mcp.tool
def excel_show_tables(directory: str = None) -> str:
    """显示Excel中所有可用的表名（这些名称在SQL查询中用作表名），请求参数不需要包装成包含server_name和tool_name的结构，而是直接传递啊Args

    Args/arguments:
        directory: Excel工作目录，默认为当前目录
    """
    try:
        # 检查是否是非标准参数格式
        if isinstance(directory, dict) and "args" in directory:
            # 非标准格式: {args: {directory: "..."}}
            actual_directory = directory.get("args", {}).get("directory", default_excel_directory)
        elif isinstance(directory, dict) and "directory" in directory:
            # 另一种非标准格式: {directory: "..."}
            actual_directory = directory.get("directory", default_excel_directory)
        else:
            # 标准格式
            actual_directory = directory if directory is not None else default_excel_directory
        
        # 使用线程池执行器运行异步代码
        with concurrent.futures.ThreadPoolExecutor() as executor:
            future = executor.submit(_execute_sql_sync, "SHOW TABLES", actual_directory)
            return future.result()
    except Exception as e:
        return f"错误: {str(e)}"

@mcp.tool
def excel_query(sql: str = None, directory: str = None) -> str:  # pyright: ignore[reportArgumentType]
    """执行SQL查询Excel数据，表名应为工作表名称而非文件名，请求参数不需要包装成包含server_name和tool_name的结构，而是直接传递啊Args
    
    Args/arguments:
        sql: SQL查询语句，支持SELECT、SHOW TABLES、SHOW CREATE TABLE等。注意：表名应为工作表名称
        directory: Excel文件所在的目录路径（可选，默认使用已设置的目录）
    """
    try:
        # 检查是否是非标准参数格式
        actual_sql = sql
        actual_directory = directory if directory is not None else default_excel_directory
        
        # 处理sql参数的非标准格式
        if isinstance(sql, dict) and "args" in sql:
            actual_sql = sql.get("args", {}).get("sql")
            actual_directory = sql.get("args", {}).get("directory", actual_directory)
        elif isinstance(sql, dict) and "sql" in sql:
            actual_sql = sql.get("sql")
            actual_directory = sql.get("directory", actual_directory)
            
        # 处理directory参数的非标准格式
        if isinstance(directory, dict) and "args" in directory:
            actual_directory = directory.get("args", {}).get("directory", actual_directory)
        elif isinstance(directory, dict) and "directory" in directory:
            actual_directory = directory.get("directory", actual_directory)
        
        if actual_sql is None:
            return "错误: SQL查询语句不能为空"
            
        # 使用线程池执行器运行异步代码
        with concurrent.futures.ThreadPoolExecutor() as executor:
            future = executor.submit(_execute_sql_sync, actual_sql, actual_directory)
            return future.result()
    except Exception as e:
        return f"错误: {str(e)}"

@mcp.tool
def excel_get_table_schema(sheet_name: str = None, directory: str = None) -> str:
    """获取指定表的结构定义，表名应为工作表名称而非文件名，请求参数不需要包装成包含server_name和tool_name的结构，而是直接传递啊Args
    
    Args/arguments:
        sheet_name: 表名（应为Sheet名称，不是Excel文件名）
        directory: Excel文件所在的目录路径（可选，默认使用已设置的目录）
    """
    try:
        # 检查是否是非标准参数格式
        actual_table_name = sheet_name
        actual_directory = directory if directory is not None else default_excel_directory
        
        # 处理table_name参数的非标准格式
        if isinstance(sheet_name, dict) and "args" in sheet_name:
            actual_table_name = sheet_name.get("args", {}).get("table_name")
            actual_directory = sheet_name.get("args", {}).get("directory", actual_directory)
        elif isinstance(sheet_name, dict) and "table_name" in sheet_name:
            actual_table_name = sheet_name.get("table_name")
            actual_directory = sheet_name.get("directory", actual_directory)
            
        # 处理directory参数的非标准格式
        if isinstance(directory, dict) and "args" in directory:
            actual_directory = directory.get("args", {}).get("directory", actual_directory)
        elif isinstance(directory, dict) and "directory" in directory:
            actual_directory = directory.get("directory", actual_directory)
        
        if actual_table_name is None:
            return "错误: 表名不能为空"
            
        # 使用线程池执行器运行异步代码
        with concurrent.futures.ThreadPoolExecutor() as executor:
            future = executor.submit(_get_create_table_sync, actual_table_name, actual_directory)
            return future.result()
    except Exception as e:
        return f"错误: {str(e)}"

@mcp.tool
def excel_refresh_cache(directory: str = None) -> str:
    """刷新Excel文件缓存，重新加载所有文件，请求参数不需要包装成包含server_name和tool_name的结构，而是直接传递啊Args
    
    Args/arguments:
        directory: Excel文件所在的目录路径（可选，默认使用已设置的目录）
    """
    try:
        # 智能解析参数
        parsed_args = smart_parse_arguments({"directory": directory} if directory is not None else {})
        actual_directory = parsed_args.get("directory", directory)
        
        excel_dir = actual_directory if actual_directory is not None else default_excel_directory
        # 使用线程池执行器运行异步代码
        with concurrent.futures.ThreadPoolExecutor() as executor:
            future = executor.submit(_refresh_cache_sync, excel_dir)
            return future.result()
    except Exception as e:
        return f"错误: {str(e)}"

@mcp.tool
def excel_list_sheets(directory: str = None) -> str:
    """列出所有Excel工作表，请求参数不需要包装成包含server_name和tool_name的结构，而是直接传递啊Args
    
    Args/arguments:
        directory: Excel文件所在的目录路径（可选，默认使用已设置的目录）
    """
    try:
        # 智能解析参数
        parsed_args = smart_parse_arguments({"directory": directory} if directory is not None else {})
        actual_directory = parsed_args.get("directory", directory)
        
        excel_dir = actual_directory if actual_directory is not None else default_excel_directory
        # 使用线程池执行器运行异步代码
        with concurrent.futures.ThreadPoolExecutor() as executor:
            future = executor.submit(_get_tables_sync, excel_dir)
            return future.result()
    except Exception as e:
        return f"错误: {str(e)}"

def _execute_sql_sync(sql: str, directory: str) -> str:
    """同步执行SQL语句"""
    return _run_async_task(_execute_sql_internal(sql, directory))

def _get_create_table_sync(table_name: str, directory: str) -> str:
    """同步获取表结构"""
    return _run_async_task(_get_create_table_internal(table_name, directory))

def _get_tables_sync(directory: str) -> str:
    """同步获取所有表"""
    return _run_async_task(_get_tables_internal(directory))

def _refresh_cache_sync(directory: str) -> str:
    """同步刷新缓存"""
    return _run_async_task(_refresh_cache_internal(directory))

def _run_async_task(coro):
    """运行异步任务"""
    try:
        # 创建新的事件循环
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        result = loop.run_until_complete(coro)
        loop.close()
        return _format_result(result)
    except Exception as e:
        return f"执行异步任务失败: {str(e)}"

async def _execute_sql_internal(sql: str, directory: str) -> Dict[str, Any]:
    """执行SQL语句"""
    request = {
        "method": "execute_sql",
        "params": {"sql": sql}
    }
    return await _send_request_to_excel_tool(request, directory)

async def _get_create_table_internal(table_name: str, directory: str) -> Dict[str, Any]:
    """获取表结构"""
    request = {
        "method": "get_create_table",
        "params": {"table": table_name}
    }
    return await _send_request_to_excel_tool(request, directory)

async def _get_tables_internal(directory: str) -> Dict[str, Any]:
    """获取所有表"""
    request = {
        "method": "get_tables",
        "params": {}
    }
    return await _send_request_to_excel_tool(request, directory)

async def _refresh_cache_internal(directory: str) -> Dict[str, Any]:
    """刷新缓存"""
    request = {
        "method": "refresh",
        "params": {}
    }
    return await _send_request_to_excel_tool(request, directory)

async def _send_request_to_excel_tool(request: Dict[str, Any], directory: str) -> Dict[str, Any]:
    """发送请求到Excel工具"""
    try:
        # 确定Excel工具路径
        script_dir = Path(__file__).parent
        excel_tool_path = script_dir / "ExcelSqlTool" / "bin" / "Debug" / "net48" / "ExcelSqlTool.exe"
        
        # 检查Excel工具是否存在
        if not excel_tool_path.exists():
            # 如果默认路径不存在，尝试在上级目录查找
            parent_dir = script_dir.parent
            excel_tool_path = parent_dir / "ExcelSqlTool" / "bin" / "Debug" / "net48" / "ExcelSqlTool.exe"
            
        if not excel_tool_path.exists():
            raise Exception(f"Excel工具未找到: {excel_tool_path}")
        
        # 启动Excel工具进程
        logger.info(f"启动Excel工具进程: {excel_tool_path} {directory}")
        
        # 使用subprocess.run而不是asyncio.create_subprocess_exec以避免事件循环问题
        request_json = json.dumps(request, ensure_ascii=False)
        input_data = (request_json + '\nquit\n').encode('utf-8')
        
        # 运行进程
        result = subprocess.run(
            [str(excel_tool_path), directory],
            input=input_data,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=30
        )
        
        # 处理错误输出
        if result.stderr:
            error_text = result.stderr.decode('utf-8', errors='ignore')
            logger.error(f"Excel工具错误: {error_text}")
        
        # 解析响应
        response_text = result.stdout.decode('utf-8', errors='ignore')
        logger.info(f"Excel工具响应: {response_text}")
        
        # 查找JSON响应
        lines = response_text.split('\n')
        for line in lines:
            line = line.strip()
            if line.startswith('{') and line.endswith('}'):
                try:
                    return json.loads(line)
                except json.JSONDecodeError:
                    continue
        
        # 如果没有找到JSON响应，返回原始响应
        if response_text.strip():
            return {"raw_response": response_text}
        
        raise Exception("无法解析Excel工具响应")
        
    except subprocess.TimeoutExpired:
        raise Exception("Excel工具响应超时")
    except Exception as e:
        logger.error(f"调用Excel工具失败: {str(e)}")
        raise Exception(f"调用Excel工具失败: {str(e)}")

def _format_result(result: Any) -> str:
    """格式化结果"""
    try:
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as e:
        return f"格式化结果失败: {str(e)}"

if __name__ == "__main__":
    try:
        # 运行FastMCP服务器
        mcp.run()
    except KeyboardInterrupt:
        print("FastMCP服务器已停止")
    except Exception as e:
        print(f"FastMCP服务器错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)