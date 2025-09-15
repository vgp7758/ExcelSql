# Excel SQL工具 MCP服务器

这个项目现在支持两种运行模式：

## 1. 传统控制台模式

这是原有的运行模式，通过标准输入输出接收JSON格式的请求。

### 启动方式
```bash
ExcelSqlTool.exe ./XLSX
```

### 使用方式
程序启动后，可以通过标准输入发送JSON格式的请求，例如：
```json
{
  "method": "execute_sql",
  "params": {
    "sql": "SELECT * FROM Sheet1"
  }
}
```

## 2. MCP服务器模式

这是新增的运行模式，直接作为MCP服务器运行，可以通过stdio与IDE或其他MCP客户端通信。

### 启动方式
```bash
ExcelSqlTool.exe ./XLSX --mcp
```

或者使用批处理文件：
```bash
run_mcp_server.bat
```

### 支持的MCP工具

#### 1. excel_show_tables
显示Excel中所有可用的表名（这些名称在SQL查询中用作表名）

**参数：** 无

**示例：**
```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "call_tool",
  "params": {
    "name": "excel_show_tables",
    "arguments": {}
  }
}
```

#### 2. excel_query
执行SQL查询Excel数据，表名应为工作表名称而非文件名

**参数：**
- `sql`: SQL查询语句，支持SELECT、SHOW TABLES、SHOW CREATE TABLE等

**示例：**
```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "call_tool",
  "params": {
    "name": "excel_query",
    "arguments": {
      "sql": "SELECT * FROM Sheet1 WHERE id > 10"
    }
  }
}
```

#### 3. excel_get_table_schema
获取指定表的结构定义，表名应为工作表名称而非文件名

**参数：**
- `table_name`: 表名（应为工作表名称，不是Excel文件名）

**示例：**
```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "call_tool",
  "params": {
    "name": "excel_get_table_schema",
    "arguments": {
      "table_name": "Sheet1"
    }
  }
}
```

#### 4. excel_refresh_cache
刷新Excel文件缓存，重新加载所有文件

**参数：** 无

**示例：**
```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "call_tool",
  "params": {
    "name": "excel_refresh_cache",
    "arguments": {}
  }
}
```

#### 5. excel_list_sheets
列出所有Excel工作表

**参数：** 无

**示例：**
```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "call_tool",
  "params": {
    "name": "excel_list_sheets",
    "arguments": {}
  }
}
```

## MCP协议支持

### 初始化
客户端首先需要发送初始化请求：

```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "initialize",
  "params": {
    "protocolVersion": "2024-11-05",
    "capabilities": {
      "roots": {
        "listChanged": true
      }
    },
    "clientInfo": {
      "name": "example-client",
      "version": "1.0.0"
    }
  }
}
```

### 获取工具列表
```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "list_tools",
  "params": {}
}
```

### 调用工具
如上所述，使用`call_tool`方法调用具体的工具。

## 与Python版本的对比

C#版本的MCP服务器提供了与Python版本相同的功能，但有以下优势：

1. **性能更好**：直接在.NET环境中运行，无需跨进程调用
2. **更稳定**：减少了进程间通信的复杂性
3. **更易于部署**：只需要一个可执行文件，无需Python环境
4. **更好的集成**：可以直接在C#项目中调试和维护

## 注意事项

1. 确保XLSX目录存在并包含Excel文件
2. 在MCP服务器模式下，程序会持续运行直到收到关闭信号或用户按下Ctrl+C
3. 所有SQL查询中的表名应该是Excel工作表的名称，而不是文件名
4. 程序支持标准的SQL SELECT语句，以及SHOW TABLES和SHOW CREATE TABLE语句