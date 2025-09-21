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
        public override string description => "����Excel�ļ�Ŀ¼�����¼���ָ��Ŀ¼�е�����Excel�ļ�";
        public override object inputSchema => new
        {
            type = "object",
            properties = new
            {
                directory = new
                {
                    type = "string",
                    description = "�µ�Excel�ļ�Ŀ¼·��"
                }
            },
            required = new[] { "directory" }
        };

        public override Task<object> CallAsync(JObject arguments)
        {
            var newDirectory = arguments?["directory"]?.ToString();
            if (string.IsNullOrEmpty(newDirectory))
            {
                throw new System.ArgumentException("Ŀ¼·������Ϊ��");
            }

            return Task.Run<object>(() =>
            {
                var oldDirectory = _excelManager.DirectoryPath;
                _excelManager.UpdateDirectoryPath(newDirectory);
                return (object)new Dictionary<string, string>
                {
                    { "old_directory", oldDirectory },
                    { "new_directory", newDirectory },
                    { "message", "Ŀ¼�ѳɹ�����" }
                };
            });
        }
    }
}
