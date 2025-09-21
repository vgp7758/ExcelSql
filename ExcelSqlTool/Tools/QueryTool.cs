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
        public override string description => "ִ��SQL��ѯExcel���ݣ�����ӦΪ���������ƶ����ļ���";
        public override object inputSchema => new
        {
            type = "object",
            properties = new
            {
                sql = new
                {
                    type = "string",
                    description = "SQL��ѯ��䣬֧��SELECT��SHOW TABLES��SHOW CREATE TABLE�ȡ�ע�⣺����ӦΪ����������"
                }
            },
            required = new[] { "sql" }
        };

        public override Task<object> CallAsync(JObject arguments)
        {
            var sql = arguments?["sql"]?.ToString();
            if (string.IsNullOrEmpty(sql))
            {
                throw new ArgumentException("SQL��ѯ��䲻��Ϊ��");
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
                            { "message", $"�ɹ����� {updateResult} ������" }
                        };
                    case SqlParser.SqlStatementType.Delete:
                        var deleteStatement = SqlParser.ParseDelete(sql);
                        var deleteResult = _excelManager.ExecuteDelete(
                            deleteStatement.TableName,
                            deleteStatement.WhereClause);
                        return (object)new System.Collections.Generic.Dictionary<string, object>
                        {
                            { "affectedRows", deleteResult },
                            { "message", $"�ɹ�ɾ�� {deleteResult} ������" }
                        };
                    default:
                        throw new ArgumentException($"��֧�ֵ�SQL�������: {statementType}");
                }
            });
        }
    }
}
