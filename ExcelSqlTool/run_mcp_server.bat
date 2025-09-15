@echo off
echo 启动Excel SQL工具 MCP服务器...
echo.

REM 检查是否存在XLSX目录
if not exist "XLSX" (
    echo 错误: 未找到XLSX目录
    echo 请确保在项目根目录运行此脚本
    pause
    exit /b 1
)

REM 启动MCP服务器模式
ExcelSqlTool.exe XLSX --mcp

echo.
echo MCP服务器已停止
pause