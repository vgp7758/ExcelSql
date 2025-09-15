#!/usr/bin/env python3
"""
简单的MCP测试
"""
import json
import subprocess
import sys

def test_mcp_communication():
    """测试MCP通信"""
    print("=== 测试MCP通信 ===")

    # MCP请求
    request = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "tools/call",
        "params": {
            "name": "excel_show_tables",
            "arguments": {}
        }
    }

    request_json = json.dumps(request, ensure_ascii=False)
    print(f"发送请求:")
    print(request_json)

    print("\n使用方法:")
    print("1. 在一个窗口运行: ExcelSqlTool.exe ../XLSX --mcp")
    print("2. 将上面的JSON复制粘贴到该窗口")
    print("3. 观察返回的响应")
    print("4. 查看控制台的DEBUG输出")

def test_queries():
    """测试查询"""
    print("\n=== 测试查询 ===")

    queries = [
        "SELECT COUNT(*) FROM ActionType",
        "SELECT * FROM ActionType LIMIT 3",
        "SELECT * FROM Config"
    ]

    for i, sql in enumerate(queries, 1):
        request = {
            "jsonrpc": "2.0",
            "id": i + 1,
            "method": "tools/call",
            "params": {
                "name": "excel_query",
                "arguments": {
                    "sql": sql
                }
            }
        }

        print(f"\n查询 {i}: {sql}")
        print("请求JSON:")
        print(json.dumps(request, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    print("ExcelDB MCP 测试")
    print("================")

    print("当前状态:")
    print("- Excel文件已成功加载")
    print("- ActionType表: 11行数据")
    print("- Config表: 3行数据")
    print("- Preload表: 1行数据")

    test_mcp_communication()
    test_queries()

    print("\n预期结果:")
    print("1. COUNT(*) 查询应返回: [{\"COUNT(*)\": 11}]")
    print("2. LIMIT查询应返回前3行数据")
    print("3. Config查询应返回3行配置数据")