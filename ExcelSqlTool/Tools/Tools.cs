using System;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace ExcelSqlTool.Tools
{
    public abstract class ToolBase
    {
        public object Info {
            get
            {
                return new
                {
                    name,
                    description,
                    inputSchema
                };
            }
        }
        public virtual string name { get; }
        public virtual string description { get; }
        public virtual object inputSchema { get; }
        public abstract Task<object> CallAsync(JObject arguments);
    }

}