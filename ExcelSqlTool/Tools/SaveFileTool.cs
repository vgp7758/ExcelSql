using System.Collections.Generic;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace ExcelSqlTool.Tools
{
    public class SaveFileTool : ToolBase
    {
        private readonly ExcelManager _excelManager;
        public SaveFileTool(ExcelManager excelManager)
        {
            _excelManager = excelManager;
        }

        public override string name => "excel_save_file";
        public override string description => "保存指定Excel文件的修改";
        public override object inputSchema => new
        {
            type = "object",
            properties = new
            {
                file_name = new
                {
                    type = "string",
                    description = "Excel文件名（如：data.xlsx）"
                }
            },
            required = new[] { "file_name" }
        };

        public override Task<object> CallAsync(JObject arguments)
        {
            var fileName = arguments?["file_name"]?.ToString();
            if (string.IsNullOrEmpty(fileName))
            {
                throw new System.ArgumentException("文件名不能为空");
            }

            return Task.Run<object>(() =>
            {
                var success = _excelManager.SaveChanges(fileName);
                return (object)new Dictionary<string, object>
                {
                    { "success", success },
                    { "message", success ? $"成功保存文件 {fileName} 的修改" : $"保存文件 {fileName} 失败" }
                };
            });
        }
    }
}
