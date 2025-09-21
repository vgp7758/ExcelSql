@echo off
echo 启动ExcelSqlTool调试测试...
echo ================================

cd /d "D:\Projects\ExcelDB\ExcelSqlTool"

echo 启动ExcelSqlTool.exe...
echo 请在另一个窗口中发送MCP请求来测试功能
echo 按Ctrl+C停止服务

".\bin\Debug\net48\ExcelSqlTool.exe" "..\XLSX"

pause