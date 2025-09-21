using System.Collections.Generic;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace ExcelSqlTool.Tools
{
    public class ChangeDirectoryTool : ToolBase
    {
        private readonly ExcelManager _excelManager;
        public ChangeDirectoryTool(ExcelManager excelManager)
        {
            _excelManager = excelManager;
        }

        public override string name => "excel_change_directory";
        public override string description => "更改Excel文件目录，重新加载指定目录中的所有Excel文件";
        public override object inputSchema => new
        {
            type = "object",
            properties = new
            {
                directory = new
                {
                    type = "string",
                    description = "新的Excel文件目录路径"
                }
            },
            required = new[] { "directory" }
        };

        public override Task<object> CallAsync(JObject arguments)
        {
            var newDirectory = arguments?["directory"]?.ToString();
            if (string.IsNullOrEmpty(newDirectory))
            {
                throw new System.ArgumentException("目录路径不能为空");
            }

            return Task.Run<object>(() =>
            {
                var oldDirectory = _excelManager.DirectoryPath;
                _excelManager.UpdateDirectoryPath(newDirectory);
                return (object)new Dictionary<string, string>
                {
                    { "old_directory", oldDirectory },
                    { "new_directory", newDirectory },
                    { "message", "目录已成功更改" }
                };
            });
        }
    }
}
