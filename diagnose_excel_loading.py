#!/usr/bin/env python3
"""
诊断Excel加载问题
"""

import json

def create_diagnostic_request():
    """创建诊断请求"""
    request = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "tools/call",
        "params": {
            "name": "excel_show_tables",
            "arguments": {}
        }
    }

    print("=== 诊断请求 ===")
    print("显示所有表的请求:")
    print(json.dumps(request, indent=2))

    # 简单的SELECT查询
    query_request = {
        "jsonrpc": "2.0",
        "id": 2,
        "method": "tools/call",
        "params": {
            "name": "excel_query",
            "arguments": {
                "sql": "SELECT COUNT(*) FROM Sheet1"
            }
        }
    }

    print("\n简单的COUNT查询:")
    print(json.dumps(query_request, indent=2))

    # 查询所有数据
    select_all_request = {
        "jsonrpc": "2.0",
        "id": 3,
        "method": "tools/call",
        "params": {
            "name": "excel_query",
            "arguments": {
                "sql": "SELECT * FROM Sheet1 LIMIT 10"
            }
        }
    }

    print("\n查询所有数据:")
    print(json.dumps(select_all_request, indent=2))

if __name__ == "__main__":
    print("ExcelDB 诊断工具")
    print("================")
    print("使用方法:")
    print("1. 启动 ExcelSqlTool.exe")
    print("2. 将上述JSON请求发送到服务器")
    print("3. 检查响应以确定问题")

    create_diagnostic_request()

    print("\n故障排除提示:")
    print("- 如果没有表返回，检查Excel文件路径")
    print("- 如果查询返回空数据，检查字段名判断逻辑")
    print("- 如果出现错误，检查Excel文件格式")