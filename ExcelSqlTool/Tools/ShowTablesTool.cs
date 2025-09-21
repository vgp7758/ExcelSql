using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace ExcelSqlTool.Tools
{
    public class ShowTablesTool : ToolBase
    {
        public override string name => "excel_show_tables";
        public override string description => "显示Excel中所有可用的表名（这些名称在SQL查询中用作表名）";
        public override object inputSchema => new 
        {
            type = "object",
            properties = new { },
            required = new string[] { }
        };
        private readonly ExcelManager _excelManager;
        public ShowTablesTool(ExcelManager excelManager)
        {
            _excelManager = excelManager;
        }
        public override Task<object> CallAsync(JObject arguments)
        {
            var tables = _excelManager.GetTableNames();
            return Task.FromResult<object>(tables);
        }
    }

}