#!/usr/bin/env python3
"""
测试MCP请求
"""
import json
import subprocess
import sys
import os

def send_mcp_request(request):
    """发送MCP请求到ExcelSqlTool"""
    try:
        # 将JSON请求写入临时文件
        with open("temp_request.json", "w", encoding="utf-8") as f:
            f.write(json.dumps(request, ensure_ascii=False))

        # 使用PowerShell发送请求
        ps_script = f'''
        $request = Get-Content "temp_request.json" -Raw | ConvertFrom-Json
        $jsonRequest = $request | ConvertTo-Json -Depth 10
        Write-Output $jsonRequest
        '''

        result = subprocess.run(["powershell", "-Command", ps_script],
                              capture_output=True, text=True, encoding="utf-8")

        # 清理临时文件
        if os.path.exists("temp_request.json"):
            os.remove("temp_request.json")

        return result.stdout

    except Exception as e:
        return f"Error: {e}"

def test_show_tables():
    """测试显示表"""
    print("=== 测试显示所有表 ===")

    request = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "tools/call",
        "params": {
            "name": "excel_show_tables",
            "arguments": {}
        }
    }

    print("请求:")
    print(json.dumps(request, indent=2, ensure_ascii=False))
    print("\n响应:")
    response = send_mcp_request(request)
    print(response)

def test_query():
    """测试查询"""
    print("\n=== 测试查询ActionType表 ===")

    request = {
        "jsonrpc": "2.0",
        "id": 2,
        "method": "tools/call",
        "params": {
            "name": "excel_query",
            "arguments": {
                "sql": "SELECT COUNT(*) FROM ActionType"
            }
        }
    }

    print("请求:")
    print(json.dumps(request, indent=2, ensure_ascii=False))
    print("\n响应:")
    response = send_mcp_request(request)
    print(response)

def test_select_data():
    """测试选择数据"""
    print("\n=== 测试选择ActionType表数据 ===")

    request = {
        "jsonrpc": "2.0",
        "id": 3,
        "method": "tools/call",
        "params": {
            "name": "excel_query",
            "arguments": {
                "sql": "SELECT * FROM ActionType LIMIT 3"
            }
        }
    }

    print("请求:")
    print(json.dumps(request, indent=2, ensure_ascii=False))
    print("\n响应:")
    response = send_mcp_request(request)
    print(response)

if __name__ == "__main__":
    print("MCP请求测试")
    print("==========")
    print("注意：ExcelSqlTool.exe应该在另一个窗口中运行")
    print("当前加载的表：ActionType, Config, Preload")

    test_show_tables()
    test_query()
    test_select_data()