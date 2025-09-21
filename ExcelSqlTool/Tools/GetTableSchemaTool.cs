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
        public override string description => "��ȡָ����Ľṹ���壬����ӦΪ���������ƶ����ļ���";
        public override object inputSchema => new
        {
            type = "object",
            properties = new
            {
                table_name = new
                {
                    type = "string",
                    description = "������ӦΪ���������ƣ�����Excel�ļ�����"
                }
            },
            required = new[] { "table_name" }
        };

        public override Task<object> CallAsync(JObject arguments)
        {
            var tableName = arguments?["table_name"]?.ToString();
            if (string.IsNullOrEmpty(tableName))
            {
                throw new System.ArgumentException("��������Ϊ��");
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
