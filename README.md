# Excel SQL Tool

Excel SQL Tool是一个允许使用SQL语法查询Excel文件的工具。它将Excel文件作为数据库使用，使开发者能够像操作数据库一样对Excel文件进行结构化查询。

## 核心概念

### FileName vs SheetName

在使用ExcelDB时，需要理解两个重要概念的区别：

1. **FileName（文件名）**: Excel文件的名称，如 `Data.xlsx`
2. **SheetName（工作表名）**: Excel文件中的工作表名称，如 `Sheet1`，这些名称在SQL查询中用作表名

**重要提示**: SQL查询应使用SheetName作为表名，而不是FileName。
例如：`SELECT * FROM Sheet1` 而不是 `SELECT * FROM Data.xlsx`

## 项目结构

```
ExcelDB/
├── ExcelSqlTool/           # C#核心实现
│   ├── ExcelManager.cs     # Excel文件管理
│   ├── SqlParser.cs        # SQL解析器
│   ├── McpHandler.cs       # MCP协议处理
│   ├── Models.cs           # 数据模型
│   └── Program.cs          # 程序入口
├── XLSX/                   # Excel文件目录
├── mcp_server.py           # Python MCP服务器
├── fastmcp_server.py       # FastMCP服务器实现
├── run.bat                 # 启动脚本
└── test_excel_sql.ps1      # 测试脚本
```

## 快速开始

### 1. 准备Excel文件

将Excel文件（.xlsx格式）放置在 `XLSX` 目录中。确保Excel文件满足以下格式要求：
- 第1行为列名
- 第2-3行为预留行
- 第4行开始为数据行

### 2. 构建项目

使用Visual Studio或命令行构建C#项目：

```bash
# 使用MSBuild构建
msbuild ExcelSqlTool/ExcelSqlTool.csproj
```

### 3. 启动MCP服务器

ExcelDB支持两种MCP服务器实现：

#### 方案1: 使用标准MCP服务器

```bash
python mcp_server.py ./XLSX
```

#### 方案2: 使用FastMCP服务器（推荐）

```bash
python fastmcp_server.py ./XLSX
```

## 使用示例

### 显示所有表名

```
SHOW TABLES
```
这将显示所有可用的工作表名称。

### 查询数据

```
SELECT * FROM ActionType LIMIT 5
```
这里 `ActionType` 是工作表名称，不是Excel文件名。

### 获取表结构

```
SHOW CREATE TABLE Config
```
这里 `Config` 是工作表名称。

## 支持的SQL语句

- `SELECT * FROM Sheet1` - 查询工作表数据
- `SHOW TABLES` - 显示所有工作表
- `SHOW CREATE TABLE Sheet1` - 显示工作表结构

## MCP服务器使用

ExcelDB支持两种MCP服务器实现：

### 标准MCP服务器 (mcp_server.py)

提供以下工具：

#### excel_show_tables() -> str
显示Excel中所有可用的表名（这些名称在SQL查询中用作表名）
- **返回**: 表名列表的JSON格式

#### excel_query(sql: str, directory: str = None) -> str
执行SQL查询Excel数据，表名应为工作表名称而非文件名
- **参数**: 
  - sql - SQL查询语句，支持SELECT、SHOW TABLES、SHOW CREATE TABLE等。注意：表名应为工作表名称
  - directory - Excel文件所在的目录路径（可选）
- **返回**: 查询结果的JSON格式

#### excel_get_table_schema(table_name: str, directory: str = None) -> str
获取指定表的结构定义，表名应为工作表名称而非文件名
- **参数**: 
  - table_name - 表名（应为工作表名称，不是Excel文件名）
  - directory - Excel文件所在的目录路径（可选）
- **返回**: 表结构定义的JSON格式

#### excel_refresh_cache(directory: str = None) -> str
刷新Excel文件缓存，重新加载所有文件
- **参数**: 
  - directory - Excel文件所在的目录路径（可选）
- **返回**: 操作结果

#### excel_list_sheets(directory: str = None) -> str
列出所有Excel工作表
- **参数**: 
  - directory - Excel文件所在的目录路径（可选）
- **返回**: 工作表列表

### FastMCP服务器 (fastmcp_server.py)

FastMCP服务器提供了更简洁的API和更好的参数处理：

#### set_excel_directory(directory: str = None) -> str
设置Excel工作目录
- **参数**: 
  - directory - Excel文件所在的目录路径
- **返回**: 操作结果

#### get_excel_directory() -> str
获取当前Excel工作目录
- **返回**: 当前目录路径

#### excel_show_tables(directory: str = None) -> str
显示Excel中所有可用的表名（这些名称在SQL查询中用作表名）
- **参数**: 
  - directory - Excel文件所在的目录路径（可选）
- **返回**: 表名列表

#### excel_query(sql: str = None, directory: str = None) -> str
执行SQL查询Excel数据，表名应为工作表名称而非文件名
- **参数**: 
  - sql - SQL查询语句，支持SELECT、SHOW TABLES、SHOW CREATE TABLE等。注意：表名应为工作表名称
  - directory - Excel文件所在的目录路径（可选）
- **返回**: 查询结果

#### excel_get_table_schema(table_name: str = None, directory: str = None) -> str
获取指定表的结构定义，表名应为工作表名称而非文件名
- **参数**: 
  - table_name - 表名（应为工作表名称，不是Excel文件名）
  - directory - Excel文件所在的目录路径（可选）
- **返回**: 表结构定义

#### excel_refresh_cache(directory: str = None) -> str
刷新Excel文件缓存，重新加载所有文件
- **参数**: 
  - directory - Excel文件所在的目录路径（可选）
- **返回**: 操作结果

#### excel_list_sheets(directory: str = None) -> str
列出所有Excel工作表
- **参数**: 
  - directory - Excel文件所在的目录路径（可选）
- **返回**: 工作表列表

## 参数处理

### 智能参数解析

ExcelDB的MCP服务器实现了智能参数解析功能，能够自动处理IDE agent可能的参数包装问题：

1. **标准参数格式**:
   ```json
   {
     "directory": "d:\\Projects\\Bunker\\TableTools\\XLSX"
   }
   ```

2. **被包装的参数格式** (IDE agent有时会错误地发送这种格式):
   ```json
   {
     "args": {
       "directory": "d:\\Projects\\Bunker\\TableTools\\XLSX"
     }
   }
   ```

3. **支持的包装键**:
   - `args`
   - `parameters`
   - `params`
   - `arguments`

服务器会自动检测并解包这些被包装的参数，确保工具能够正确接收参数。

## 常见问题

### 1. 表名不存在错误

如果遇到"表不存在"的错误，请确认：
- 使用的是工作表名称（SheetName）而不是Excel文件名（FileName）
- 工作表名称拼写正确
- Excel文件已正确放置在XLSX目录中

### 2. MCP服务器连接问题

确保：
- Python环境已安装所需依赖
- C#项目已成功构建
- Excel文件格式符合要求

### 3. 参数包装问题

某些IDE agent可能会错误地将参数包装在额外的JSON结构中。ExcelDB的智能参数解析功能会自动处理这个问题，但如果遇到参数传递问题，请检查IDE的MCP配置。

## 开发指南

### 构建和测试

1. 构建C#项目：
   ```bash
   msbuild ExcelSqlTool/ExcelSqlTool.csproj
   ```

2. 运行测试脚本：
   ```bash
   # PowerShell测试
   .\test_excel_sql.ps1
   
   # Python测试
   python test_sql_queries.py
   ```

### 扩展功能

可以根据需要扩展以下功能：
- 支持更多SQL语句类型
- 增加数据写入功能
- 支持更多Excel格式
- 添加数据验证功能