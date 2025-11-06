using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace ExcelSqlTool.Tools
{
    public class GetDirectoryTool : ToolBase
    {
        private readonly ExcelManager _excelManager;
        public GetDirectoryTool(ExcelManager excelManager) { _excelManager = excelManager; }
        public override string name => "excel_get_directory";
        public override string description => "获取当前Excel目录路径及其是否存在";
        public override object inputSchema => new { type = "object", properties = new { }, required = new string[0] };
        public override Task<object> CallAsync(JObject arguments)
        {
            var result = new
            {
                directory = _excelManager.DirectoryPath,
                exists = _excelManager.IsDirectoryExists
            };
            return Task.FromResult<object>(result);
        }
    }
}
