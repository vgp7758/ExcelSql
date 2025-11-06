using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace ExcelSqlTool.Tools
{
    public class GetStatsTool : ToolBase
    {
        private readonly ExcelManager _excelManager;
        public GetStatsTool(ExcelManager excelManager) { _excelManager = excelManager; }
        public override string name => "excel_get_stats";
        public override string description => "获取当前加载的文件/表/行统计信息";
        public override object inputSchema => new { type = "object", properties = new { }, required = new string[0] };
        public override Task<object> CallAsync(JObject arguments)
        {
            var stats = _excelManager.GetModificationStats();
            return Task.FromResult<object>(stats);
        }
    }
}
