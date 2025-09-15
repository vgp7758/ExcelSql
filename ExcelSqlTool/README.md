# Excel SQL工具

Excel SQL工具是一个允许通过SQL语句操作Excel文件的命令行工具。它将Excel文件视为数据库表，支持常见的SQL操作。

**最新更新**: v1.1.0 - 已修复MCP响应格式和SQL WHERE条件问题，现在完全支持复杂的SQL查询操作。

## 功能特性

- 通过SQL语句操作Excel文件
- 支持SELECT查询操作
- 支持SHOW TABLES显示所有工作表
- 支持SHOW CREATE TABLE获取建表语句
- 使用MCP协议与外部程序通信
- 自动监控指定目录下的Excel文件

## 技术栈

- .NET Framework 4.8
- NPOI库（用于Excel文件操作）
- Newtonsoft.Json（用于JSON处理）

## 安装和构建

```bash
# 克隆项目后，进入项目目录
cd ExcelSqlTool

# 构建项目
dotnet build
```

## 使用方法

### 启动工具

```bash
ExcelSqlTool.exe <excel_directory>
```

例如：
```bash
ExcelSqlTool.exe ./XLSX
```

### MCP协议接口

工具启动后会等待MCP协议请求。可以通过标准输入发送JSON格式的请求。

#### 支持的MCP方法

1. `get_tables` - 获取所有表名
2. `execute_sql` - 执行SQL语句
3. `get_create_table` - 获取建表语句
4. `refresh` - 刷新文件缓存

#### 请求格式

```json
{
  "method": "方法名",
  "params": {
    // 参数
  }
}
```

#### 响应格式

```json
{
  "result": // 结果数据
}
```

或

```json
{
  "error": {
    "message": "错误信息"
  }
}
```

### 支持的SQL语句

1. `SHOW TABLES` - 显示所有表
2. `SHOW CREATE TABLE <table_name>` - 获取建表语句
3. `SELECT * FROM <table_name>` - 查询表中所有数据
4. `SELECT <columns> FROM <table_name> WHERE <condition>` - 条件查询

### 数据模型

#### Excel文件结构

- 工具监控指定目录下的所有.xlsx文件
- 每个Excel文件可以包含多个工作表
- 每个工作表作为一个独立的表进行操作

#### 表结构规则

- 前3行为表头信息：
  - 第1行：列名
  - 第2行：数据类型（可选）
  - 第3行：描述信息（可选）
- 第4行开始为数据行

## 示例

### 启动工具

```bash
ExcelSqlTool.exe ./XLSX
```

### 获取所有表名

```json
{
  "method": "get_tables",
  "params": {}
}
```

### 执行SHOW TABLES

```json
{
  "method": "execute_sql",
  "params": {
    "sql": "SHOW TABLES"
  }
}
```

### 查询数据

```json
{
  "method": "execute_sql",
  "params": {
    "sql": "SELECT * FROM Actions"
  }
}
```

### 获取建表语句

```json
{
  "method": "execute_sql",
  "params": {
    "sql": "SHOW CREATE TABLE Actions"
  }
}
```

或

```json
{
  "method": "get_create_table",
  "params": {
    "table": "Actions"
  }
}
```

## 项目结构

```
ExcelSqlTool/
├── ExcelManager.cs       # Excel文件操作管理器
├── McpHandler.cs         # MCP协议处理器
├── McpServer.cs          # 原生MCP服务器实现
├── Models.cs             # 数据模型定义
├── Program.cs            # 程序入口
├── SqlParser.cs          # SQL解析器
├── ExcelSqlTool.csproj   # 项目文件
├── run_mcp_server.bat    # MCP服务器启动脚本
├── MCP_SERVER_README.md  # MCP服务器详细文档
└── README.md             # 说明文档
```

## 依赖库

- NPOI.dll - 核心Excel操作库
- NPOI.OOXML.dll - Office Open XML格式支持
- NPOI.OpenXml4Net.dll - Open XML格式支持
- NPOI.OpenXmlFormats.dll - Open XML格式定义
- ICSharpCode.SharpZipLib.dll - 压缩库
- Newtonsoft.Json.dll - JSON处理库

## 注意事项

1. 工具只读取.xlsx格式的Excel文件
2. 表名使用工作表名称，不使用文件名
3. 工具会自动推断列的数据类型
4. WHERE条件现在支持完整的比较操作（=, >, <, >=, <=, !=）
5. 支持复杂的列名，包括包含特殊字符的列名（使用反引号引用）

## 版本更新

### v1.1.0 最新改进
- **修复**: MCP响应双重序列化问题，现在正确返回SQL查询结果
- **修复**: SQL WHERE条件操作符转换，支持完整的比较操作
- **新增**: 原生C# MCP服务器实现，性能更优
- **新增**: 反引号列名支持，可处理复杂列名
- **改进**: 表名大小写敏感处理，保留原始大小写
- **新增**: 完整的测试套件和详细文档