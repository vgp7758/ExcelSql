#!/usr/bin/env python3
"""
调试测试脚本
"""
import json
import sys
import subprocess
import os

def create_test_request():
    """创建测试请求"""
    request = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "tools/call",
        "params": {
            "name": "excel_show_tables",
            "arguments": {}
        }
    }

    return request

def create_query_request():
    """创建查询请求"""
    request = {
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

    return request

def main():
    """主函数"""
    print("ExcelDB 调试测试")
    print("================")

    print("\n1. 显示所有表的请求:")
    print(json.dumps(create_test_request(), indent=2, ensure_ascii=False))

    print("\n2. 查询COUNT(*)的请求:")
    print(json.dumps(create_query_request(), indent=2, ensure_ascii=False))

    print("\n使用方法:")
    print("1. 运行 start_debug_test.bat 启动服务")
    print("2. 观察控制台的DEBUG输出")
    print("3. 将上述JSON复制到MCP客户端发送")
    print("4. 检查响应结果")

if __name__ == "__main__":
    main()