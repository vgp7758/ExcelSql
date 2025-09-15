using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.CodeDom;
using System.CodeDom.Compiler;
using Microsoft.CSharp;
using System.Reflection;

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
        /// 处理反引号列名，移除反引号并返回原始列名
        /// </summary>
        /// <param name="columnExpression">列表达式</param>
        /// <returns>处理后的列名</returns>
        private static string ProcessBacktickColumn(string columnExpression)
        {
            if (string.IsNullOrWhiteSpace(columnExpression))
                return columnExpression;

            // 检查是否被反引号包围
            var backtickMatch = Regex.Match(columnExpression, @"`([^`]+)`");
            if (backtickMatch.Success)
            {
                var unquotedColumn = backtickMatch.Groups[1].Value;
                Console.WriteLine($"DEBUG: 处理反引号列名: {columnExpression} -> {unquotedColumn}");
                return unquotedColumn;
            }

            // 如果没有反引号，直接返回
            return columnExpression;
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

            // 标准化SQL语句，移除多余的空格和换行，但保留原始表名大小写
            var normalizedSql = Regex.Replace(sql.Trim(), @"\s+", " ");

            // 解析SELECT部分，使用正则表达式匹配但保留原始大小写
            var selectMatch = Regex.Match(normalizedSql, @"SELECT\s+(.*?)\s+FROM\s+([a-zA-Z_][a-zA-Z0-9_]*)(.*)", RegexOptions.IgnoreCase);
            if (!selectMatch.Success)
                throw new ArgumentException("无法解析SELECT语句");

            // 解析列
            var columnsPart = selectMatch.Groups[1].Value.Trim();
            if (columnsPart == "*")
            {
                select.Columns = new List<string> { "*" };
            }
            else
            {
                select.Columns = columnsPart.Split(',')
                    .Select(c => ProcessBacktickColumn(c.Trim()))
                    .ToList();
            }

            select.TableName = selectMatch.Groups[2].Value.Trim();

            // 解析剩余部分
            var remainingClause = selectMatch.Groups[3].Value.Trim();

            // 分离不同的子句
            ParseClauses(remainingClause, select);

            return select;
        }

        /// <summary>
        /// 解析SQL子句（WHERE, JOIN, LIMIT等）
        /// </summary>
        /// <param name="clause">子句字符串</param>
        /// <param name="select">SELECT语句对象</param>
        private static void ParseClauses(string clause, SelectStatement select)
        {
            var tokens = TokenizeClause(clause);
            int index = 0;

            while (index < tokens.Count)
            {
                var token = tokens[index].ToUpper();

                switch (token)
                {
                    case "WHERE":
                        index = ParseWhereClause(tokens, index + 1, select);
                        break;
                    case "INNER":
                    case "LEFT":
                    case "RIGHT":
                    case "FULL":
                    case "JOIN":
                        index = ParseJoinClause(tokens, index, select);
                        break;
                    case "LIMIT":
                        if (index + 1 < tokens.Count && int.TryParse(tokens[index + 1], out int limit))
                        {
                            select.Limit = limit;
                            index += 2;
                        }
                        else
                        {
                            index += 1;
                        }
                        break;
                    case "ORDER":
                    case "GROUP":
                        // 跳过ORDER BY和GROUP BY子句
                        if (index + 1 < tokens.Count && tokens[index + 1].ToUpper() == "BY")
                        {
                            index += 2;
                            // 跳过直到下一个子句
                            while (index < tokens.Count && !IsClauseKeyword(tokens[index]))
                                index++;
                        }
                        else
                        {
                            index += 1;
                        }
                        break;
                    default:
                        index += 1;
                        break;
                }
            }
        }

        /// <summary>
        /// 解析WHERE子句
        /// </summary>
        /// <param name="tokens">token列表</param>
        /// <param name="startIndex">开始位置</param>
        /// <param name="select">SELECT语句对象</param>
        /// <returns>下一个位置</returns>
        private static int ParseWhereClause(List<string> tokens, int startIndex, SelectStatement select)
        {
            var whereClause = new System.Text.StringBuilder();
            int index = startIndex;

            while (index < tokens.Count && !IsClauseKeyword(tokens[index]))
            {
                if (whereClause.Length > 0)
                    whereClause.Append(" ");
                whereClause.Append(tokens[index]);
                index++;
            }

            select.WhereClause = whereClause.ToString().Trim();
            return index;
        }

        /// <summary>
        /// 解析JOIN子句
        /// </summary>
        /// <param name="tokens">token列表</param>
        /// <param name="startIndex">开始位置</param>
        /// <param name="select">SELECT语句对象</param>
        /// <returns>下一个位置</returns>
        private static int ParseJoinClause(List<string> tokens, int startIndex, SelectStatement select)
        {
            var join = new JoinClause();
            int index = startIndex;

            // 解析JOIN类型
            if (tokens[index].ToUpper() == "INNER" || tokens[index].ToUpper() == "LEFT" ||
                tokens[index].ToUpper() == "RIGHT" || tokens[index].ToUpper() == "FULL")
            {
                join.JoinType = tokens[index].ToUpper();
                index++;
                // 跳过JOIN关键字
                if (index < tokens.Count && tokens[index].ToUpper() == "JOIN")
                    index++;
            }
            else
            {
                join.JoinType = "INNER"; // 默认INNER JOIN
            }

            // 解析表名
            if (index < tokens.Count)
            {
                join.TableName = tokens[index];
                index++;
            }

            // 解析ON子句
            if (index < tokens.Count && tokens[index].ToUpper() == "ON")
            {
                index++; // 跳过ON
                var onClause = new System.Text.StringBuilder();

                while (index < tokens.Count && !IsClauseKeyword(tokens[index]))
                {
                    if (onClause.Length > 0)
                        onClause.Append(" ");
                    onClause.Append(tokens[index]);
                    index++;
                }

                join.OnClause = onClause.ToString().Trim();
            }

            select.Joins.Add(join);
            return index;
        }

        /// <summary>
        /// 将子句字符串转换为token列表
        /// </summary>
        /// <param name="clause">子句字符串</param>
        /// <returns>token列表</returns>
        private static List<string> TokenizeClause(string clause)
        {
            var tokens = new List<string>();
            var token = new System.Text.StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < clause.Length; i++)
            {
                char c = clause[i];

                if (c == '\'' || c == '"')
                {
                    inQuotes = !inQuotes;
                    token.Append(c);
                }
                else if (char.IsWhiteSpace(c) && !inQuotes)
                {
                    if (token.Length > 0)
                    {
                        tokens.Add(token.ToString());
                        token.Clear();
                    }
                }
                else if (IsOperator(c) && !inQuotes)
                {
                    if (token.Length > 0)
                    {
                        tokens.Add(token.ToString());
                        token.Clear();
                    }
                    tokens.Add(c.ToString());
                }
                else
                {
                    token.Append(c);
                }
            }

            if (token.Length > 0)
            {
                tokens.Add(token.ToString());
            }

            return tokens;
        }

        /// <summary>
        /// 判断是否为运算符
        /// </summary>
        /// <param name="c">字符</param>
        /// <returns>是否为运算符</returns>
        private static bool IsOperator(char c)
        {
            return c == '=' || c == '<' || c == '>' || c == '!' || c == '+' || c == '-' || c == '*' || c == '/' || c == '(' || c == ')';
        }

        /// <summary>
        /// 判断是否为子句关键字
        /// </summary>
        /// <param name="token">token</param>
        /// <returns>是否为关键字</returns>
        private static bool IsClauseKeyword(string token)
        {
            var upperToken = token.ToUpper();
            return upperToken == "WHERE" || upperToken == "INNER" || upperToken == "LEFT" ||
                   upperToken == "RIGHT" || upperToken == "FULL" || upperToken == "JOIN" ||
                   upperToken == "LIMIT" || upperToken == "ORDER" || upperToken == "GROUP";
        }

        /// <summary>
        /// 数学表达式求值器
        /// </summary>
        public static class MathExpressionEvaluator
        {
            /// <summary>
            /// 计算数学表达式的值
            /// </summary>
            /// <param name="expression">数学表达式</param>
            /// <param name="variables">变量字典</param>
            /// <returns>计算结果</returns>
            public static object Evaluate(string expression, Dictionary<string, object> variables = null)
            {
                if (string.IsNullOrWhiteSpace(expression))
                    return null;

                try
                {
                    // 替换变量
                    var processedExpression = ReplaceVariables(expression, variables);

                    // 编译和执行表达式
                    var result = CompileAndEvaluate(processedExpression);
                    return result;
                }
                catch (Exception ex)
                {
                    throw new ArgumentException($"无法计算表达式 '{expression}': {ex.Message}");
                }
            }

            /// <summary>
            /// 替换表达式中的变量
            /// </summary>
            /// <param name="expression">表达式</param>
            /// <param name="variables">变量字典</param>
            /// <returns>替换后的表达式</returns>
            private static string ReplaceVariables(string expression, Dictionary<string, object> variables)
            {
                if (variables == null || variables.Count == 0)
                    return expression;

                var result = expression;
                foreach (var variable in variables)
                {
                    result = result.Replace(variable.Key, ConvertToString(variable.Value));
                }
                return result;
            }

            /// <summary>
            /// 将值转换为字符串表示
            /// </summary>
            /// <param name="value">值</param>
            /// <returns>字符串表示</returns>
            public static string ConvertToString(object value)
            {
                if (value == null)
                    return "null";

                if (value is string str)
                    return $"\"{str}\"";

                if (value is DateTime dt)
                    return $"\"{dt:yyyy-MM-dd HH:mm:ss}\"";

                if (value is bool b)
                    return b ? "true" : "false";

                return value.ToString();
            }

            /// <summary>
            /// 编译并执行表达式
            /// </summary>
            /// <param name="expression">表达式</param>
            /// <returns>计算结果</returns>
            private static object CompileAndEvaluate(string expression)
            {
                // 将SQL的比较操作符转换为C#语法
                var processedExpression = expression;

                // 将 = 转换为 == （但避免转换 == 为 ===）
                processedExpression = Regex.Replace(processedExpression, @"(?<!=)=(?!=)", "==");

                // 创建C#代码
                var code = $@"
using System;
public class ExpressionEvaluator
{{
    public static object Evaluate()
    {{
        return ({processedExpression});
    }}
}}";

                // 编译代码
                var compilerParams = new CompilerParameters
                {
                    GenerateExecutable = false,
                    GenerateInMemory = true
                };

                var compiler = new CSharpCodeProvider();
                var results = compiler.CompileAssemblyFromSource(compilerParams, code);

                if (results.Errors.HasErrors)
                {
                    var errors = string.Join("\n", results.Errors.Cast<CompilerError>());
                    throw new ArgumentException($"表达式编译错误:\n{errors}");
                }

                // 执行代码
                var assembly = results.CompiledAssembly;
                var evaluatorType = assembly.GetType("ExpressionEvaluator");
                var method = evaluatorType.GetMethod("Evaluate");

                return method.Invoke(null, null);
            }
        }

        /// <summary>
        /// 解析和计算WHERE条件中的数学表达式
        /// </summary>
        /// <param name="whereClause">WHERE子句</param>
        /// <param name="row">数据行</param>
        /// <returns>是否满足条件</returns>
        public static bool EvaluateWhereClause(string whereClause, Dictionary<string, object> row)
        {
            if (string.IsNullOrWhiteSpace(whereClause))
                return true;

            try
            {
                // 替换列名
                var expression = ReplaceColumnNames(whereClause, row);

                // 计算表达式
                var result = MathExpressionEvaluator.Evaluate(expression);

                // 转换为布尔值
                if (result is bool boolResult)
                    return boolResult;

                // 如果结果是数值，转换为布尔值
                if (result is int || result is double || result is decimal)
                {
                    var numericResult = Convert.ToDouble(result);
                    return numericResult != 0;
                }

                // 如果是字符串，检查是否为空
                if (result is string strResult)
                {
                    return !string.IsNullOrEmpty(strResult);
                }

                return result != null;
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"WHERE条件计算失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 替换WHERE子句中的列名
        /// </summary>
        /// <param name="whereClause">WHERE子句</param>
        /// <param name="row">数据行</param>
        /// <returns>替换后的表达式</returns>
        private static string ReplaceColumnNames(string whereClause, Dictionary<string, object> row)
        {
            var result = whereClause;

            // 处理带表名前缀的列名（如：table.column）
            var columnRegex = new Regex(@"\b(\w+)\.(\w+)\b");
            result = columnRegex.Replace(result, match =>
            {
                var columnName = match.Groups[2].Value;
                if (row.ContainsKey(columnName))
                {
                    return MathExpressionEvaluator.ConvertToString(row[columnName]);
                }
                return match.Value;
            });

            // 处理简单列名
            foreach (var columnName in row.Keys)
            {
                // 使用单词边界确保只匹配完整的列名
                var columnPattern = $@"\b{Regex.Escape(columnName)}\b";
                result = Regex.Replace(result, columnPattern, match =>
                {
                    return MathExpressionEvaluator.ConvertToString(row[columnName]);
                });
            }

            return result;
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

        /// <summary>
        /// 解析UPDATE语句
        /// </summary>
        /// <param name="sql">UPDATE语句</param>
        /// <returns>解析结果</returns>
        public static UpdateStatement ParseUpdate(string sql)
        {
            var updateRegex = new Regex(@"UPDATE\s+(\w+)\s+SET\s+(.*?)\s+(WHERE\s+(.*))?$", RegexOptions.IgnoreCase);
            var match = updateRegex.Match(sql.Trim());

            if (match.Success)
            {
                var update = new UpdateStatement
                {
                    TableName = match.Groups[1].Value.Trim()
                };

                // 解析SET子句
                var setClause = match.Groups[2].Value.Trim();
                var setPairs = setClause.Split(',');
                foreach (var pair in setPairs)
                {
                    var keyValue = pair.Split('=');
                    if (keyValue.Length == 2)
                    {
                        update.SetValues[keyValue[0].Trim()] = keyValue[1].Trim().Trim('\'', '"');
                    }
                }

                // 解析WHERE子句
                if (match.Groups.Count > 4 && !string.IsNullOrEmpty(match.Groups[4].Value))
                {
                    update.WhereClause = match.Groups[4].Value.Trim();
                }

                return update;
            }

            throw new ArgumentException("无法解析UPDATE语句");
        }

        /// <summary>
        /// 解析DELETE语句
        /// </summary>
        /// <param name="sql">DELETE语句</param>
        /// <returns>解析结果</returns>
        public static DeleteStatement ParseDelete(string sql)
        {
            var deleteRegex = new Regex(@"DELETE\s+FROM\s+(\w+)(?:\s+(.*))?$", RegexOptions.IgnoreCase);
            var match = deleteRegex.Match(sql.Trim());

            if (match.Success)
            {
                var delete = new DeleteStatement
                {
                    TableName = match.Groups[1].Value.Trim()
                };

                // 解析WHERE子句
                if (match.Groups.Count > 2 && !string.IsNullOrEmpty(match.Groups[2].Value))
                {
                    var remainingClause = match.Groups[2].Value.Trim();

                    // 分离WHERE子句
                    var whereMatch = Regex.Match(remainingClause, @"WHERE\s+(.*)", RegexOptions.IgnoreCase);
                    if (whereMatch.Success)
                    {
                        delete.WhereClause = whereMatch.Groups[1].Value.Trim();
                    }
                }

                return delete;
            }

            throw new ArgumentException("无法解析DELETE语句");
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
        public int? Limit { get; set; }
        public List<JoinClause> Joins { get; set; } = new List<JoinClause>();
    }

    /// <summary>
    /// JOIN子句
    /// </summary>
    public class JoinClause
    {
        public string JoinType { get; set; } // INNER, LEFT, RIGHT, FULL
        public string TableName { get; set; }
        public string OnClause { get; set; }
    }

    /// <summary>
    /// UPDATE语句解析结果
    /// </summary>
    public class UpdateStatement
    {
        public string TableName { get; set; }
        public Dictionary<string, string> SetValues { get; set; } = new Dictionary<string, string>();
        public string WhereClause { get; set; }
    }

    /// <summary>
    /// DELETE语句解析结果
    /// </summary>
    public class DeleteStatement
    {
        public string TableName { get; set; }
        public string WhereClause { get; set; }
    }
}