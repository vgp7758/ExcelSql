# MCP功能使用说明

## 配置步骤

### 1. 确保项目已正确构建

首先需要确保Excel SQL Tool项目已正确构建：

1. 使用Visual Studio打开`ExcelDB.sln`
2. 构建整个解决方案
3. 确认`ExcelSqlTool\bin\Debug\net48\`目录下生成了可执行文件

### 2. 验证Python环境

确保已安装必要的Python依赖：

```bash
pip install -r requirements.txt
```

### 3. 配置MCP服务器

编辑`mcp_config.json`文件，确保配置正确：

```json
{
  "mcpServers": {
    "excel-sql-tool": {
      "command": "python",
      "args": ["mcp_server.py", "./XLSX"]
    }
  }
}
```

## 测试MCP功能

### 1. 手动启动MCP服务器

```bash
python mcp_server.py ./XLSX
```

### 2. 使用测试脚本验证

```bash
python test_mcp_functionality.py
```

## 在IDE中配置MCP

### Qoder IDE配置

1. 打开Qoder IDE
2. 进入设置 -> MCP服务器配置
3. 添加新的MCP服务器配置：
   - 名称：excel-sql-tool
   - 命令：python
   - 参数：mcp_server.py ./XLSX
4. 保存配置并重启IDE

### 其他IDE配置

根据具体IDE的MCP配置方式进行相应配置，通常需要指向`mcp_config.json`文件。

## 常见问题及解决方案

### 1. MCP服务器启动失败

**问题**：MCP服务器无法启动或报错

**解决方案**：
- 检查Python环境和依赖是否正确安装
- 确认Excel SQL Tool项目已正确构建
- 检查路径配置是否正确

### 2. 工具列表为空

**问题**：IDE中无法看到Excel SQL Tool的工具列表

**解决方案**：
- 确认MCP服务器已正确启动
- 检查IDE的MCP配置是否正确
- 查看MCP服务器日志获取详细错误信息

### 3. 工具调用失败

**问题**：调用Excel SQL Tool工具时返回错误

**解决方案**：
- 检查Excel文件是否存在且格式正确
- 确认SQL语句语法正确
- 查看错误信息进行针对性解决

## 调试技巧

### 1. 查看MCP服务器日志

MCP服务器会输出详细的日志信息，可以帮助诊断问题：

```bash
python mcp_server.py ./XLSX
```

### 2. 使用调试脚本

项目提供了多个调试脚本：

- `debug_mcp.py` - 调试MCP通信问题
- `test_mcp_functionality.py` - 测试MCP功能
- `simple_mcp_test.py` - 简单的MCP测试

### 3. 手动测试工具调用

可以手动发送JSON-RPC请求测试工具调用：

```bash
# 启动MCP服务器
python mcp_server.py ./XLSX

# 在另一个终端中发送请求
echo '{"jsonrpc":"2.0","id":1,"method":"tools/list","params":{}}' | nc localhost 3000
```

## 性能优化建议

### 1. Excel文件优化

- 避免过大的Excel文件，建议单文件不超过10MB
- 合理设计表结构，避免过多的空行和空列
- 使用合适的数据类型，避免混合存储不同类型数据

### 2. 查询优化

- 尽量使用具体的WHERE条件过滤数据
- 避免SELECT *，只选择需要的列
- 对于复杂查询，考虑预先处理数据

### 3. 缓存管理

- 定期使用`excel_refresh_cache`工具刷新缓存
- 对于频繁修改的Excel文件，增加刷新频率
- 对于静态数据，可以减少刷新次数以提高性能

## 扩展开发

### 添加新的工具

1. 在`mcp_server.py`的`list_tools`方法中添加新的工具定义
2. 实现对应的处理方法
3. 在`call_tool`方法中添加路由逻辑

### 扩展SQL支持

1. 修改`SqlParser.cs`添加新的SQL语句解析逻辑
2. 在`McpHandler.cs`中添加相应的处理方法
3. 在`ExcelManager.cs`中实现具体的业务逻辑

## 反馈和支持

如果在使用过程中遇到问题，请提供以下信息：

1. 错误信息和日志
2. 使用的IDE和版本
3. Excel文件示例（如果可能）
4. 具体的操作步骤

可以通过以下方式获取支持：

- 提交GitHub Issue
- 联系项目维护者
- 查阅相关文档和FAQ