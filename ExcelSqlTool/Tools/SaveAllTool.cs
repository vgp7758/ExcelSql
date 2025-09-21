using System.Collections.Generic;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace ExcelSqlTool.Tools
{
    public class SaveAllTool : ToolBase
    {
        private readonly ExcelManager _excelManager;
        public SaveAllTool(ExcelManager excelManager)
        {
            _excelManager = excelManager;
        }

        public override string name => "excel_save_all";
        public override string description => "保存所有修改到Excel文件";
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
                var savedFiles = _excelManager.SaveAllChanges();
                return (object)new Dictionary<string, object>
                {
                    { "saved_files", savedFiles },
                    { "message", $"成功保存 {savedFiles} 个Excel文件的修改" }
                };
            });
        }
    }
}
