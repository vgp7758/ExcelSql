# Excel SQL Tool MCP Server

这是一个MCP服务器，将Excel SQL工具暴露给支持MCP协议的IDE使用。

## 功能特性

IDE可以通过MCP协议使用以下工具：

### 1. excel_show_tables
显示Excel中所有可用的表名
```json
{
  "name": "excel_show_tables",
  "description": "显示Excel中所有可用的表名",
  "inputSchema": {
    "type": "object",
    "properties": {},
    "required": []
  }
}
```

### 2. excel_query
执行SQL查询Excel数据
```json
{
  "name": "excel_query",
  "description": "执行SQL查询Excel数据",
  "inputSchema": {
    "type": "object",
    "properties": {
      "sql": {
        "type": "string",
        "description": "SQL查询语句，支持SELECT、SHOW TABLES、SHOW CREATE TABLE等"
      }
    },
    "required": ["sql"]
  }
}
```

### 3. excel_get_table_schema
获取指定表的结构定义
```json
{
  "name": "excel_get_table_schema",
  "description": "获取指定表的结构定义",
  "inputSchema": {
    "type": "object",
    "properties": {
      "table_name": {
        "type": "string",
        "description": "表名（工作表名称）"
      }
    },
    "required": ["table_name"]
  }
}
```

### 4. excel_refresh_cache
刷新Excel文件缓存，重新加载所有文件
```json
{
  "name": "excel_refresh_cache",
  "description": "刷新Excel文件缓存，重新加载所有文件",
  "inputSchema": {
    "type": "object",
    "properties": {},
    "required": []
  }
}
```

### 5. excel_list_sheets
列出所有Excel工作表
```json
{
  "name": "excel_list_sheets",
  "description": "列出所有Excel工作表",
  "inputSchema": {
    "type": "object",
    "properties": {},
    "required": []
  }
}
```

## IDE配置

### Claude Desktop配置
将以下配置添加到Claude Desktop的配置文件中：

```json
{
  "mcpServers": {
    "excel-sql-tool": {
      "command": "python",
      "args": ["D:\\Projects\\ExcelDB\\mcp_server.py"],
      "env": {
        "PYTHONPATH": "D:\\Projects\\ExcelDB"
      }
    }
  }
}
```

### 其他支持MCP的IDE
根据IDE的具体配置方式，使用`mcp_config.json`中的配置。

## 使用示例

### 在IDE中使用
1. 配置完成后，IDE会自动发现Excel SQL工具
2. 可以通过自然语言调用工具，例如：
   - "显示所有Excel表"
   - "查询ActionType表中的所有数据"
   - "获取Config表的结构"

### 手动测试
```bash
# 安装依赖
pip install mcp

# 测试MCP服务器
python mcp_server.py
```

## 支持的SQL语句

- `SHOW TABLES` - 显示所有表
- `SHOW CREATE TABLE <table_name>` - 显示表结构
- `SELECT * FROM <table_name>` - 查询所有数据
- `SELECT <columns> FROM <table_name> WHERE <condition>` - 条件查询

## 工作原理

1. IDE通过MCP协议发送请求
2. Python MCP服务器接收请求
3. 服务器调用Excel SQL工具（C#程序）
4. Excel工具处理Excel文件并返回结果
5. MCP服务器将结果返回给IDE

## 依赖

- Python 3.8+
- mcp Python包
- .NET Framework 4.8 (Excel SQL工具)
- Excel文件在XLSX目录中