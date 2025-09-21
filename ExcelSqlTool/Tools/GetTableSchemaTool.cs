using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace ExcelSqlTool.Tools
{
    public class GetTableSchemaTool : ToolBase
    {
        private readonly ExcelManager _excelManager;
        public GetTableSchemaTool(ExcelManager excelManager)
        {
            _excelManager = excelManager;
        }

        public override string name => "excel_get_table_schema";
        public override string description => "获取指定表的结构定义，表名应为工作表名称而非文件名";
        public override object inputSchema => new
        {
            type = "object",
            properties = new
            {
                table_name = new
                {
                    type = "string",
                    description = "表名（应为工作表名称，不是Excel文件名）"
                }
            },
            required = new[] { "table_name" }
        };

        public override Task<object> CallAsync(JObject arguments)
        {
            var tableName = arguments?["table_name"]?.ToString();
            if (string.IsNullOrEmpty(tableName))
            {
                throw new System.ArgumentException("表名不能为空");
            }

            return Task.Run<object>(() =>
            {
                var results = _excelManager.GetCreateTableStatementsByFileName(tableName);
                if (results.Count == 1)
                {
                    return (object)new System.Collections.Generic.Dictionary<string, string>
                    {
                        { "table", results[0]["table"] },
                        { "createTable", results[0]["createTable"] }
                    };
                }
                return (object)results;
            });
        }
    }
}
