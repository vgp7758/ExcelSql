using System;
using System.Collections.Generic;

namespace ExcelSqlTool.Models
{
    /// <summary>
    /// Excel文件模型
    /// </summary>
    public class ExcelFile
    {
        /// <summary>
        /// 文件名
        /// </summary>
        public string FileName { get; set; }
        
        /// <summary>
        /// 文件路径
        /// </summary>
        public string Path { get; set; }
        
        /// <summary>
        /// 最后修改时间
        /// </summary>
        public DateTime LastModified { get; set; }
        
        /// <summary>
        /// 工作表集合
        /// </summary>
        public Dictionary<string, Worksheet> Worksheets { get; set; } = new Dictionary<string, Worksheet>();
    }

    /// <summary>
    /// 工作表模型
    /// </summary>
    public class Worksheet
    {
        /// <summary>
        /// 工作表名称（用作表名）
        /// </summary>
        public string Name { get; set; }
        
        /// <summary>
        /// 列定义
        /// </summary>
        public List<Column> Headers { get; set; } = new List<Column>();
        
        /// <summary>
        /// 数据行
        /// </summary>
        public List<Dictionary<string, object>> DataRows { get; set; } = new List<Dictionary<string, object>>();
        
        /// <summary>
        /// 所属文件路径
        /// </summary>
        public string FilePath { get; set; }
    }

    /// <summary>
    /// 列模型
    /// </summary>
    public class Column
    {
        /// <summary>
        /// 列名
        /// </summary>
        public string Name { get; set; }
        
        /// <summary>
        /// 数据类型
        /// </summary>
        public string DataType { get; set; }
        
        /// <summary>
        /// 列索引
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// 描述（第三行COMMENTS）
        /// </summary>
        public string Comments { get; set; }
    }
}