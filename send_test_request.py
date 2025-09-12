#!/usr/bin/env python3
"""
发送测试请求到FastMCP服务器
"""

import json
import sys

def send_initialize_request():
    """发送初始化请求"""
    request = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "initialize",
        "params": {
            "protocolVersion": "2024-11-05",
            "capabilities": {},
            "clientInfo": {
                "name": "test-client",
                "version": "1.0.0"
            }
        }
    }
    print(json.dumps(request))
    sys.stdout.flush()

def send_tool_list_request():
    """发送工具列表请求"""
    request = {
        "jsonrpc": "2.0",
        "id": 2,
        "method": "tools/list",
        "params": {}
    }
    print(json.dumps(request))
    sys.stdout.flush()

def send_nonstandard_tool_call_request():
    """发送非标准格式的工具调用请求"""
    request = {
        "jsonrpc": "2.0",
        "id": 3,
        "method": "tools/call",
        "params": {
            "server_name": "mcp.config.usrlocalmcp.excel-sql-tool",
            "tool_name": "excel_show_tables",
            "args": {
                "directory": "d:\\Projects\\Bunker\\TableTools\\XLSX"
            }
        }
    }
    print(json.dumps(request))
    sys.stdout.flush()

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python send_test_request.py [initialize|list|call]")
        sys.exit(1)
    
    command = sys.argv[1]
    
    if command == "initialize":
        send_initialize_request()
    elif command == "list":
        send_tool_list_request()
    elif command == "call":
        send_nonstandard_tool_call_request()
    else:
        print("未知命令")
        sys.exit(1)