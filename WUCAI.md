# Excel SQL Tool Project Context

## Project Overview

This is an Excel SQL Tool project that allows querying Excel files using SQL syntax. The project consists of:

1. A C# Excel SQL tool (`ExcelSqlTool.exe`) that processes Excel files and executes SQL queries
2. A Python MCP (Model Context Protocol) server (`mcp_server.py`) that exposes the Excel tool to IDEs
3. Excel files stored in the `XLSX` directory

The main purpose is to enable developers to interact with Excel data using familiar SQL syntax directly from their IDE.

## Key Technologies

- **C# (.NET Framework 4.8)**: Core Excel processing tool
- **Python 3.8+**: MCP server implementation
- **NPOI**: Library for reading Excel files
- **MCP (Model Context Protocol)**: Protocol for IDE integration
- **JSON**: Communication format between components

## Project Structure

```
ExcelDB/
├── mcp_server.py          # MCP server that exposes Excel tools to IDEs
├── ExcelSqlTool/          # C# project for Excel processing
│   ├── bin/Debug/net48/   # Built executable
│   └── ExcelSqlTool.exe   # Main Excel tool executable
├── XLSX/                  # Directory containing Excel files
├── requirements.txt       # Python dependencies
├── config.xml            # Configuration file
└── MCP_README.md         # Documentation
```

## Building and Running

### Prerequisites
- .NET Framework 4.8
- Python 3.8+
- Excel files in the `XLSX` directory

### Building the C# Tool
```bash
dotnet build ExcelDB.sln
```

### Installing Python Dependencies
```bash
pip install -r requirements.txt
```

### Running the MCP Server
```bash
python mcp_server.py
```

Or as a module:
```bash
python -m mcp_server
```

### IDE Configuration
To use this tool in an IDE that supports MCP:

1. **Claude Desktop**: Add this configuration to Claude's config file:
   ```json
   {
     "mcpServers": {
       "excel-sql-tool": {
         "command": "python",
         "args": ["-m", "mcp_server", "D:\\Projects\\Bunker\\TableTools\\XLSX"],
         "env": {
           "PYTHONPATH": "D:\\Projects\\ExcelDB"
         }
       }
     }
   }
   ```

2. **Other IDEs**: Use the configuration in `mcp_config.json`

## Available Tools

The MCP server exposes 5 tools to the IDE:

1. **excel_show_tables** - Display all available table names in Excel
2. **excel_query** - Execute SQL queries on Excel data
3. **excel_get_table_schema** - Get the schema definition of a specified table
4. **excel_refresh_cache** - Refresh Excel file cache and reload all files
5. **excel_list_sheets** - List all Excel worksheets

## Supported SQL Statements

- `SHOW TABLES` - Show all tables
- `SHOW CREATE TABLE <table_name>` - Show table structure
- `SELECT * FROM <table_name>` - Query all data
- `SELECT <columns> FROM <table_name> WHERE <condition>` - Conditional queries

## Development Workflow

1. Place Excel files in the `XLSX` directory
2. Build the C# tool with `dotnet build`
3. Run the Python MCP server
4. Configure your IDE to connect to the MCP server
5. Use natural language in your IDE to query Excel data:
   - "Show all Excel tables"
   - "Query all data in the ActionType table"
   - "Get the structure of the Config table"

## Configuration

The `config.xml` file contains paths and settings:
- `xlsPath`: Path to Excel files (default: `./XLSX`)
- `jsonPath`: Path for JSON output
- `csPath`: Path for C# script generation
- `bytePath`: Path for binary output
- Various flags for code generation and read-only mode

## Recent Fixes and Improvements

- Fixed tool list not showing in IDE by correctly registering tool handlers using decorators
- Fixed MCP server initialization issues
- Improved error handling and logging
- Made Excel tool path resolution more flexible to work when called from other repositories
- Added support for running as a Python module (`python -m mcp_server`)
- Added better error messages when Excel tool is not found
- Fixed Pydantic validation errors when constructing CallToolResult objects
- Improved error handling for all tool responses to ensure they follow the MCP protocol specification