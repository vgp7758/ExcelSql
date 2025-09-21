using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace ExcelSqlTool.Tools
{
    public class RefreshCacheTool : ToolBase
    {
        private readonly ExcelManager _excelManager;
        public RefreshCacheTool(ExcelManager excelManager)
        {
            _excelManager = excelManager;
        }

        public override string name => "excel_refresh_cache";
        public override string description => "刷新Excel文件缓存，重新加载所有文件";
        public override object inputSchema => new
        {
            type = "object",
            properties = new { },
            required = new string[0]
        };

        public override Task<object> CallAsync(JObject arguments)
        {
            return Task.Run<object>(() =>
            {
                _excelManager.Refresh();
                return (object)"缓存已刷新";
            });
        }
    }
}
