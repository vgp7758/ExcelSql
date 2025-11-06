# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

ExcelDB is a C# command-line tool that allows SQL operations on Excel files using the MCP (Model Context Protocol) communication standard. The tool treats Excel worksheets as database tables and supports basic SQL queries. It operates in two modes: traditional CLI mode and MCP server mode for AI integration.

## Build and Development

### Build Commands
```bash
# Build the solution
dotnet build ExcelDB.sln

# Build specific project
dotnet build ExcelSqlTool/ExcelSqlTool.csproj

# Run the application (traditional mode)
dotnet run --project ExcelSqlTool/ExcelSqlTool.csproj -- <excel_directory>

# Run in MCP server mode
dotnet run --project ExcelSqlTool/ExcelSqlTool.csproj -- <excel_directory> --mcp
```

### Dependencies
The project uses external DLL dependencies referenced in the project file:
- **NPOI.dll** - Core Excel file manipulation
- **NPOI.OOXML.dll** - OOXML format support
- **NPOI.OpenXml4Net.dll** - OpenXML packaging
- **NPOI.OpenXmlFormats.dll** - OpenXML format handling
- **ICSharpCode.SharpZipLib.dll** - ZIP compression
- **Newtonsoft.Json.dll** - JSON serialization

### Configuration
- `config.xml` - Main configuration with paths for Excel files, JSON output, C# generation, and Unity integration
- Excel files expected in `XLSX/` directory by default
- Output binaries in `ExcelSqlTool/bin/Debug/net48/`

## Architecture

### Core Components

1. **Program.cs** - Entry point with dual mode support
   - Handles command-line argument parsing
   - Supports traditional CLI and MCP server modes
   - Sets up UTF-8 encoding for proper character handling

2. **ExcelManager.cs** - Manages Excel file operations and caching
   - Loads and parses .xlsx files from specified directory
   - Maintains in-memory cache of Excel data
   - Handles file monitoring and refresh operations
   - Supports Unity integration paths from config.xml

3. **McpHandler.cs** - MCP protocol communication handler
   - Processes JSON requests via standard input
   - Supports methods: execute_sql, get_tables, get_create_table, refresh
   - Returns JSON responses following MCP format
   - Handles both synchronous and asynchronous operations

4. **SqlParser.cs** - SQL statement parser and executor
   - Supports basic SQL operations: SELECT, SHOW TABLES, SHOW CREATE TABLE
   - Implements simple WHERE clause filtering with basic operators (=, !=, >, <, >=, <=, LIKE)
   - Converts Excel data to tabular format for queries
   - Handles column projection and row filtering

5. **Models.cs** - Data models for Excel representation
   - ExcelFile, Worksheet, Column classes
   - Represents Excel structure as relational data
   - Supports data type mapping from Excel to SQL types

### Key Design Patterns
- **Dual Mode Architecture** - Single binary supporting both CLI and MCP modes
- **In-Memory Caching** - Excel data loaded once and cached for performance
- **Protocol-Based Communication** - MCP JSON protocol for AI tool integration
- **Header-Driven Schema** - Excel first 3 rows define column metadata (name, type, description)

## Usage Patterns

### Traditional Mode
```bash
ExcelSqlTool.exe ./XLSX
```

### MCP Server Mode
```bash
ExcelSqlTool.exe ./XLSX --mcp
```

### MCP Communication Protocol
The tool reads JSON requests from stdin and writes responses to stdout in MCP format:

Request format:
```json
{
  "method": "execute_sql|get_tables|get_create_table|refresh",
  "params": { "sql": "SELECT * FROM table" }
}
```

Supported methods:
- `execute_sql` - Execute SQL statements
- `get_tables` - List all available tables (worksheets)
- `get_create_table` - Get table schema as CREATE TABLE statement
- `refresh` - Refresh file cache and reload Excel data

### Excel Structure Requirements
- .xlsx format only
- First 3 rows reserved for headers:
  - Row 1: Column names
  - Row 2: Data types (string, int, float, bool, datetime)
  - Row 3: Descriptions (optional)
- Data starts from row 4
- Each worksheet represents a database table
- Empty rows and columns are automatically trimmed

## Testing

### Test Scripts
- Python test client: `test_tool.py`
- PowerShell test script: `test_excel_sql.ps1`
- Test Excel files in `XLSX/` directory

### Manual Testing
```bash
# Test basic functionality
dotnet run --project ExcelSqlTool/ExcelSqlTool.csproj -- ./XLSX

# Test MCP mode
dotnet run --project ExcelSqlTool/ExcelSqlTool.csproj -- ./XLSX --mcp
```

## Development Notes

### .NET Framework 4.8 Targeting
- Project targets .NET Framework 4.8 for Windows compatibility
- Uses external DLL references instead of NuGet packages
- All dependencies must be present in the root directory

### Unicode Support
- Console input/output configured for UTF-8 encoding
- Proper handling of international characters in Excel data

### Error Handling
- Comprehensive try-catch blocks in all major operations
- Graceful degradation for malformed Excel files
- Detailed error messages in MCP responses