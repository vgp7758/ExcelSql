using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace ExcelSqlTool.Tools
{
    public class ListSheetsTool : ToolBase
    {
        private readonly ExcelManager _excelManager;
        public ListSheetsTool(ExcelManager excelManager)
        {
            _excelManager = excelManager;
        }

        public override string name => "excel_list_sheets";
        public override string description => "列出所有Excel工作表";
        public override object inputSchema => new
        {
            type = "object",
            properties = new { },
            required = new string[0]
        };

        public override Task<object> CallAsync(JObject arguments)
        {
            return Task.Run<object>(() => (object)_excelManager.GetTableNames());
        }
    }
}
