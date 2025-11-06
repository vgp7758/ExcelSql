using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace ExcelSqlTool.Tools
{
    public class ListFilesTool : ToolBase
    {
        private readonly ExcelManager _excelManager;
        public ListFilesTool(ExcelManager excelManager) { _excelManager = excelManager; }
        public override string name => "excel_list_files";
        public override string description => "列出已加载的Excel文件及其工作表名称";
        public override object inputSchema => new { type = "object", properties = new { }, required = new string[0] };
        public override Task<object> CallAsync(JObject arguments)
        {
            var list = _excelManager.GetLoadedFilesInfo();
            return Task.FromResult<object>(list);
        }
    }
}
