# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

ExcelDB is a C# command-line tool that allows SQL operations on Excel files using the MCP (Model Context Protocol) communication standard. The tool treats Excel worksheets as database tables and supports basic SQL queries.

## Build and Development

### Build Commands
```bash
# Build the project
dotnet build ExcelDB.sln

# Build specific project
dotnet build ExcelSqlTool/ExcelSqlTool.csproj

# Run the application
dotnet run --project ExcelSqlTool/ExcelSqlTool.csproj -- <excel_directory>
```

### Testing
- Use `test_tool.py` for Python-based testing
- Use `test_excel_sql.ps1` for PowerShell-based testing  
- Test files are located in the `XLSX/` directory

## Architecture

### Core Components

1. **ExcelManager.cs** - Manages Excel file operations and caching
   - Loads and parses .xlsx files from specified directory
   - Maintains in-memory cache of Excel data
   - Handles file monitoring and refresh operations

2. **McpHandler.cs** - MCP protocol communication handler
   - Processes JSON requests via standard input
   - Supports methods: execute_sql, get_tables, get_create_table, refresh
   - Returns JSON responses following MCP format

3. **SqlParser.cs** - SQL statement parser and executor
   - Supports basic SQL operations: SELECT, SHOW TABLES, SHOW CREATE TABLE
   - Implements simple WHERE clause filtering
   - Converts Excel data to tabular format for queries

4. **Models.cs** - Data models for Excel representation
   - ExcelFile, Worksheet, Column classes
   - Represents Excel structure as relational data

### Dependencies
- **NPOI** - Excel file manipulation library
- **Newtonsoft.Json** - JSON serialization
- **ICSharpCode.SharpZipLib** - ZIP file handling
- **.NET Framework 4.8** - Target framework

### Configuration
- `config.xml` - Main configuration file with paths and settings
- Excel files expected in `XLSX/` directory
- Output binaries in `ExcelSqlTool/bin/Debug/net48/`

## Usage Patterns

### Starting the Tool
```bash
ExcelSqlTool.exe ./XLSX
```

### MCP Communication
The tool reads JSON requests from stdin and writes responses to stdout. Supported methods:
- `execute_sql` - Execute SQL statements
- `get_tables` - List all available tables
- `get_create_table` - Get table schema
- `refresh` - Refresh file cache

### Excel Structure Requirements
- .xlsx format only
- First 3 rows reserved for headers (column name, data type, description)
- Data starts from row 4
- Each worksheet represents a table

## Testing Commands

```bash
# Python test script
python test_tool.py

# PowerShell test script
.\test_excel_sql.ps1
```

## File Structure
```
ExcelDB/
├── ExcelSqlTool/           # Main C# project
│   ├── ExcelManager.cs     # Excel operations
│   ├── McpHandler.cs       # MCP protocol handler
│   ├── SqlParser.cs        # SQL parsing
│   ├── Models.cs           # Data models
│   └── Program.cs          # Entry point
├── XLSX/                   # Excel files for testing
├── config.xml              # Configuration
├── test_tool.py           # Python test client
└── test_excel_sql.ps1     # PowerShell test script
```