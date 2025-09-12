#!/usr/bin/env python3
"""
Pydantic模型测试
"""

from mcp.types import CallToolRequestParams

def test_pydantic_model():
    """测试Pydantic模型"""
    # 测试正确的参数格式
    params = {
        "name": "add",
        "arguments": {
            "a": 5,
            "b": 3
        }
    }
    
    try:
        # 尝试创建CallToolRequestParams对象
        request_params = CallToolRequestParams(**params)
        print(f"成功创建CallToolRequestParams对象: {request_params}")
        print(f"名称: {request_params.name}")
        print(f"参数: {request_params.arguments}")
    except Exception as e:
        print(f"创建CallToolRequestParams对象时出错: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_pydantic_model()