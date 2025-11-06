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
        public override string description => "执行SQL查询Excel（经SQLite映射），表名应为工作表名或使用标准SQL。";
        public override object inputSchema => new
        {
            type = "object",
            properties = new
            {
                sql = new
                {
                    type = "string",
                    description = "SQL语句，支持SELECT/UPDATE/DELETE/SHOW TABLES/SHOW CREATE TABLE。表名为工作表名"
                }
            },
            required = new[] { "sql" }
        };

        public override Task<object> CallAsync(JObject arguments)
        {
            var sql = arguments?[("sql")]?.ToString();
            if (string.IsNullOrEmpty(sql))
            {
                throw new ArgumentException("SQL查询语句不能为空");
            }

            return Task.Run<object>(() =>
            {
                // 直接委托给ExcelManager的SQLite执行
                return _excelManager.ExecuteSqlRaw(sql);
            });
        }
    }
}
