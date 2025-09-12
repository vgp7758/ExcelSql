using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelSqlTool
{
    /// <summary>
    /// SQL语句解析器
    /// </summary>
    public class SqlParser
    {
        /// <summary>
        /// SQL语句类型
        /// </summary>
        public enum SqlStatementType
        {
            Select,
            Insert,
            Update,
            Delete,
            CreateTable,
            AlterTable,
            ShowTables,
            ShowCreateTable,
            Refresh,
            Unknown
        }

        /// <summary>
        /// 解析SQL语句类型
        /// </summary>
        /// <param name="sql">SQL语句</param>
        /// <returns>SQL语句类型</returns>
        public static SqlStatementType ParseStatementType(string sql)
        {
            if (string.IsNullOrWhiteSpace(sql))
                return SqlStatementType.Unknown;

            var trimmedSql = sql.TrimStart().ToUpper();

            if (trimmedSql.StartsWith("SELECT"))
                return SqlStatementType.Select;
            if (trimmedSql.StartsWith("INSERT"))
                return SqlStatementType.Insert;
            if (trimmedSql.StartsWith("UPDATE"))
                return SqlStatementType.Update;
            if (trimmedSql.StartsWith("DELETE"))
                return SqlStatementType.Delete;
            if (trimmedSql.StartsWith("CREATE TABLE"))
                return SqlStatementType.CreateTable;
            if (trimmedSql.StartsWith("ALTER TABLE"))
                return SqlStatementType.AlterTable;
            if (trimmedSql.StartsWith("SHOW TABLES"))
                return SqlStatementType.ShowTables;
            if (trimmedSql.StartsWith("SHOW CREATE TABLE"))
                return SqlStatementType.ShowCreateTable;
            if (trimmedSql.StartsWith("REFRESH"))
                return SqlStatementType.Refresh;

            return SqlStatementType.Unknown;
        }

        /// <summary>
        /// 解析SELECT语句
        /// </summary>
        /// <param name="sql">SELECT语句</param>
        /// <returns>解析结果</returns>
        public static SelectStatement ParseSelect(string sql)
        {
            var select = new SelectStatement();
            
            // 简单的正则表达式解析（实际项目中需要更复杂的SQL解析器）
            var selectRegex = new Regex(@"SELECT\s+(.*?)\s+FROM\s+(\w+)(?:\s+(.*?))?$", RegexOptions.IgnoreCase);
            var match = selectRegex.Match(sql.Trim());
            
            if (match.Success)
            {
                var columnsPart = match.Groups[1].Value.Trim();
                select.TableName = match.Groups[2].Value.Trim();
                
                // 解析WHERE子句
                if (match.Groups.Count > 3 && !string.IsNullOrEmpty(match.Groups[3].Value))
                {
                    select.WhereClause = match.Groups[3].Value.Trim();
                }
                
                // 解析列
                if (columnsPart == "*")
                {
                    select.Columns = new List<string> { "*" };
                }
                else
                {
                    select.Columns = columnsPart.Split(',')
                        .Select(c => c.Trim())
                        .ToList();
                }
                
                return select;
            }
            
            throw new ArgumentException("无法解析SELECT语句");
        }

        /// <summary>
        /// 解析SHOW CREATE TABLE语句
        /// </summary>
        /// <param name="sql">SHOW CREATE TABLE语句</param>
        /// <returns>表名</returns>
        public static string ParseShowCreateTable(string sql)
        {
            var regex = new Regex(@"SHOW CREATE TABLE\s+(\w+)", RegexOptions.IgnoreCase);
            var match = regex.Match(sql.Trim());
            
            if (match.Success)
            {
                return match.Groups[1].Value.Trim();
            }
            
            throw new ArgumentException("无法解析SHOW CREATE TABLE语句");
        }
    }

    /// <summary>
    /// SELECT语句解析结果
    /// </summary>
    public class SelectStatement
    {
        public List<string> Columns { get; set; } = new List<string>();
        public string TableName { get; set; }
        public string WhereClause { get; set; }
    }
}