using System;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace ExcelSqlTool.Tools
{
    public class QueryTool : ToolBase
    {
        private readonly ExcelManager _excelManager;
        public QueryTool(ExcelManager excelManager)
        {
            _excelManager = excelManager;
        }

        public override string name => "excel_query";
        public override string description => "执行SQL查询Excel数据，表名应为工作表名称而非文件名";
        public override object inputSchema => new
        {
            type = "object",
            properties = new
            {
                sql = new
                {
                    type = "string",
                    description = "SQL查询语句，支持SELECT、SHOW TABLES、SHOW CREATE TABLE等。注意：表名应为工作表名称"
                }
            },
            required = new[] { "sql" }
        };

        public override Task<object> CallAsync(JObject arguments)
        {
            var sql = arguments?["sql"]?.ToString();
            if (string.IsNullOrEmpty(sql))
            {
                throw new ArgumentException("SQL查询语句不能为空");
            }

            return Task.Run<object>(() =>
            {
                var statementType = SqlParser.ParseStatementType(sql);
                switch (statementType)
                {
                    case SqlParser.SqlStatementType.Select:
                        var selectStatement = SqlParser.ParseSelect(sql);
                        if (selectStatement.Joins != null && selectStatement.Joins.Count > 0)
                        {
                            return (object)_excelManager.ExecuteSelectWithJoin(selectStatement);
                        }
                        else
                        {
                            return (object)_excelManager.ExecuteSelectByFileName(
                                selectStatement.TableName,
                                selectStatement.Columns,
                                selectStatement.WhereClause,
                                selectStatement.Limit);
                        }
                    case SqlParser.SqlStatementType.ShowTables:
                        return (object)_excelManager.GetTableNames();
                    case SqlParser.SqlStatementType.ShowCreateTable:
                        var tableName = SqlParser.ParseShowCreateTable(sql);
                        var createTableStatement = _excelManager.GetCreateTableStatement(tableName);
                        return (object)new System.Collections.Generic.Dictionary<string, string>
                        {
                            { "table", tableName },
                            { "createTable", createTableStatement }
                        };
                    case SqlParser.SqlStatementType.Update:
                        var updateStatement = SqlParser.ParseUpdate(sql);
                        var updateResult = _excelManager.ExecuteUpdate(
                            updateStatement.TableName,
                            updateStatement.SetValues,
                            updateStatement.WhereClause);
                        return (object)new System.Collections.Generic.Dictionary<string, object>
                        {
                            { "affectedRows", updateResult },
                            { "message", $"成功更新 {updateResult} 行数据" }
                        };
                    case SqlParser.SqlStatementType.Delete:
                        var deleteStatement = SqlParser.ParseDelete(sql);
                        var deleteResult = _excelManager.ExecuteDelete(
                            deleteStatement.TableName,
                            deleteStatement.WhereClause);
                        return (object)new System.Collections.Generic.Dictionary<string, object>
                        {
                            { "affectedRows", deleteResult },
                            { "message", $"成功删除 {deleteResult} 行数据" }
                        };
                    default:
                        throw new ArgumentException($"不支持的SQL语句类型: {statementType}");
                }
            });
        }
    }
}
