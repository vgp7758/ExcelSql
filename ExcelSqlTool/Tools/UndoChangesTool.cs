using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace ExcelSqlTool.Tools
{
    public class UndoChangesTool : ToolBase
    {
        private readonly ExcelManager _excelManager;
        public UndoChangesTool(ExcelManager excelManager) { _excelManager = excelManager; }
        public override string name => "excel_undo_changes";
        public override string description => "撤销内存中未保存的修改（重新从磁盘加载）";
        public override object inputSchema => new { type = "object", properties = new { }, required = new string[0] };
        public override Task<object> CallAsync(JObject arguments)
        {
            _excelManager.UndoChanges();
            return Task.FromResult<object>(new { message = "已重新加载文件，未保存改动已撤销" });
        }
    }
}
