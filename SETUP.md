# Excel SQL Tool Project - Configuration and Usage Guide

## Project Overview

This project provides a tool for querying Excel files using SQL syntax. It consists of:
1. A C# Excel SQL tool (`ExcelSqlTool.exe`) that processes Excel files and executes SQL queries
2. A Python MCP (Model Context Protocol) server (`mcp_server.py`) that exposes the Excel tool to IDEs
3. Excel files stored in the `XLSX` directory

## Configuration

### MCP Server Configuration
The MCP server configuration is defined in `mcp_config.json`:
```json
{
  "mcpServers": {
    "excel-sql-tool": {
      "command": "python",
      "args": ["-m", "mcp_server", "D:\\Projects\\Bunker\\TableTools\\XLSX"]
    }
  }
}
```

### IDE Configuration

#### Claude Desktop Configuration
To use this tool with Claude Desktop, add the following configuration to Claude's config file:
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

#### Other IDEs
For other IDEs that support MCP, use the configuration in `mcp_config.json`.

## Running the Server

To run the MCP server directly:
```bash
python -m mcp_server D:\Projects\Bunker\TableTools\XLSX
```

Or if you're in the project directory:
```bash
python mcp_server.py D:\Projects\Bunker\TableTools\XLSX
```

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

## Troubleshooting

1. **Ensure Excel tool is built**:
   ```bash
   dotnet build ExcelDB.sln
   ```

2. **Check Python dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Test the connection**:
   ```bash
   python test_mcp_server.py
   ```

4. **Tool list not showing in IDE**:
   - Make sure the MCP server is running correctly
   - Check that the IDE is properly configured to connect to the MCP server
   - Verify that the tools are correctly registered in the server code

5. **Path issues when using from other repositories**:
   - The server now uses flexible path resolution for the Excel tool
   - Make sure the Excel SQL tool is built before running the server
   - If you encounter path issues, you can set the `PYTHONPATH` environment variable to point to this project directory

6. **Pydantic validation errors**:
   - Fixed incorrect format when constructing CallToolResult objects
   - All tool responses now correctly follow the MCP protocol specification