# CODEBUDDY.md

This file provides guidance to CodeBuddy Code when working with this repository.

## Project Overview

ExcelDB is a C# command-line tool that allows SQL operations on Excel files using the MCP (Model Context Protocol) communication standard. The tool treats Excel worksheets as database tables and supports basic SQL queries. It consists of a C# Excel processor and a Python MCP server that bridges IDE communication.

## Build and Development Commands

### Build Commands
```bash
# Build the entire solution
dotnet build ExcelDB.sln

# Build specific project
dotnet build ExcelSqlTool/ExcelSqlTool.csproj

# Run the C# tool directly
dotnet run --project ExcelSqlTool/ExcelSqlTool.csproj -- ./XLSX

# Or run the compiled executable
ExcelSqlTool.exe ./XLSX
```

### Testing Commands
```bash
# Install Python dependencies
pip install -r requirements.txt

# Test Python MCP server
python test_mcp_server.py

# Test C# tool with PowerShell
.\test_excel_sql.ps1

# Test C# tool with Python
python test_tool.py
```

### MCP Server Commands
```bash
# Start MCP server for IDE integration
python mcp_server.py

# Test MCP server functionality
python test_mcp_server.py
```

## Architecture Overview

### Two-Layer Architecture
The system uses a two-layer architecture:
1. **Python MCP Server** (`mcp_server.py`) - Handles MCP protocol communication with IDEs
2. **C# Excel Tool** (`ExcelSqlTool.exe`) - Processes Excel files and executes SQL queries

Communication flow: `IDE ↔ MCP Protocol ↔ Python Server ↔ C# Tool ↔ Excel Files`

### Core C# Components

**ExcelManager.cs** - Central Excel file management
- Loads and caches .xlsx files from specified directory
- Monitors file changes and handles refresh operations
- Maintains in-memory representation of Excel data

**McpHandler.cs** - MCP protocol communication
- Processes JSON requests via stdin/stdout
- Supports methods: `execute_sql`, `get_tables`, `get_create_table`, `refresh`
- Handles JSON serialization/deserialization

**SqlParser.cs** - SQL statement processing
- Supports: `SELECT`, `SHOW TABLES`, `SHOW CREATE TABLE`  
- Implements basic WHERE clause filtering
- Converts Excel data to tabular query results

**Models.cs** - Data structures
- `ExcelFile`, `Worksheet`, `Column` classes
- Represents Excel structure as relational data model

### MCP Server Tools
The Python server exposes 5 tools to IDEs:
- `excel_show_tables` - List all available tables
- `excel_query` - Execute SQL queries
- `excel_get_table_schema` - Get table structure  
- `excel_refresh_cache` - Reload Excel files
- `excel_list_sheets` - List worksheets

### Excel File Structure Requirements
- Only .xlsx format supported
- First 3 rows reserved for metadata:
  - Row 1: Column names
  - Row 2: Data types  
  - Row 3: Descriptions
- Actual data starts from row 4
- Each worksheet represents a database table

### Dependencies and Framework
- **C# Project**: .NET Framework 4.8, NPOI (Excel manipulation), Newtonsoft.Json, ICSharpCode.SharpZipLib
- **Python Server**: mcp>=1.0.0
- **Build Output**: `ExcelSqlTool/bin/Debug/net48/ExcelSqlTool.exe`

### Configuration
- `config.xml` - Main configuration file
- Excel files expected in `XLSX/` directory
- MCP configuration in `mcp_config.json` for IDE integration

## Key File Locations
- Main executable: `ExcelSqlTool/bin/Debug/net48/ExcelSqlTool.exe`
- Test Excel files: `XLSX/` directory
- MCP server: `mcp_server.py`
- Configuration: `config.xml`, `mcp_config.json`