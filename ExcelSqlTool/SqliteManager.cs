using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;
using System.Globalization;
using ExcelSqlTool.Models;

namespace ExcelSqlTool
{
    /// <summary>
    /// 轻量级SQLite管理器：负责将Excel模型导入SQLite，并对SQLite执行SQL
    /// </summary>
    public class SqliteManager : IDisposable
    {
        private readonly string _dbPath;
        private readonly SQLiteConnection _conn;

        public SqliteManager(string dbPath = null)
        {
            if (string.IsNullOrWhiteSpace(dbPath))
            {
                var baseDir = Path.Combine(Path.GetTempPath(), "ExcelSqlTool");
                Directory.CreateDirectory(baseDir);
                _dbPath = Path.Combine(baseDir, Guid.NewGuid().ToString("N") + ".db");
            }
            else
            {
                Directory.CreateDirectory(Path.GetDirectoryName(dbPath));
                _dbPath = dbPath;
            }

            var cs = new SQLiteConnectionStringBuilder
            {
                DataSource = _dbPath,
                JournalMode = SQLiteJournalModeEnum.Wal,
                SyncMode = SynchronizationModes.Normal,
                FailIfMissing = false
            }.ToString();

            _conn = new SQLiteConnection(cs);
            _conn.Open();
            EnsureMetaTables();
        }

        public void Dispose()
        {
            try { _conn?.Close(); } catch { }
            try { _conn?.Dispose(); } catch { }
            try { if (File.Exists(_dbPath)) File.Delete(_dbPath); } catch { }
        }

        private static string QuoteIdent(string name)
        {
            if (name == null) return null;
            return "\"" + name.Replace("\"", "\"\"") + "\"";
        }

        private static string MapType(string excelType)
        {
            if (string.IsNullOrEmpty(excelType)) return "TEXT";
            switch (excelType.ToUpperInvariant())
            {
                case "INT": return "INTEGER";
                case "FLOAT":
                case "DOUBLE": return "REAL";
                case "DATE": return "TEXT"; // ISO-8601存储
                case "BOOLEAN": return "INTEGER"; // 0/1
                default: return "TEXT";
            }
        }

        private static DbType MapDbType(string excelType)
        {
            if (string.IsNullOrEmpty(excelType)) return DbType.String;
            switch (excelType.ToUpperInvariant())
            {
                case "INT": return DbType.Int64;
                case "FLOAT":
                case "DOUBLE": return DbType.Double;
                case "DATE": return DbType.String;
                case "BOOLEAN": return DbType.Int32;
                default: return DbType.String;
            }
        }

        private void EnsureMetaTables()
        {
            using (var cmd = _conn.CreateCommand())
            {
                cmd.CommandText = @"CREATE TABLE IF NOT EXISTS __excel_column_comments (
    table_name TEXT NOT NULL,
    column_name TEXT NOT NULL,
    comment TEXT,
    PRIMARY KEY(table_name, column_name)
)";
                cmd.ExecuteNonQuery();
            }
        }

        public bool HasTable(string tableName)
        {
            using (var cmd = _conn.CreateCommand())
            {
                cmd.CommandText = "SELECT 1 FROM sqlite_master WHERE type='table' AND name=@n LIMIT 1";
                cmd.Parameters.AddWithValue("@n", tableName);
                var o = cmd.ExecuteScalar();
                return o != null && o != DBNull.Value;
            }
        }

        public List<string> GetTables()
        {
            var list = new List<string>();
            using (var cmd = _conn.CreateCommand())
            {
                cmd.CommandText = "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' AND name != '__excel_column_comments' ORDER BY name";
                using (var rd = cmd.ExecuteReader())
                {
                    while (rd.Read()) list.Add(rd.GetString(0));
                }
            }
            return list;
        }

        private Dictionary<string, string> LoadComments(string tableName)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            using (var cmd = _conn.CreateCommand())
            {
                cmd.CommandText = "SELECT column_name, comment FROM __excel_column_comments WHERE table_name=@t";
                cmd.Parameters.AddWithValue("@t", tableName);
                using (var rd = cmd.ExecuteReader())
                {
                    while (rd.Read())
                    {
                        var col = rd.IsDBNull(0) ? null : rd.GetString(0);
                        var cmt = rd.IsDBNull(1) ? null : rd.GetString(1);
                        if (!string.IsNullOrEmpty(col)) dict[col] = cmt;
                    }
                }
            }
            return dict;
        }

        public string GetCreateTable(string tableName)
        {
            // 如存在列注释，用PRAGMA + 注释生成建表语句
            var commentMap = LoadComments(tableName);
            if (commentMap.Count > 0)
            {
                using (var cmd = _conn.CreateCommand())
                {
                    cmd.CommandText = $"PRAGMA table_info({QuoteIdent(tableName)})";
                    var cols = new List<string>();
                    using (var rd = cmd.ExecuteReader())
                    {
                        while (rd.Read())
                        {
                            var name = rd[1]?.ToString();
                            var type = rd[2]?.ToString();
                            commentMap.TryGetValue(name, out var cmt);
                            var line = new StringBuilder();
                            line.Append("    ").Append(QuoteIdent(name)).Append(' ').Append(type);
                            if (!string.IsNullOrEmpty(cmt))
                            {
                                line.Append(" -- ").Append(cmt.Replace("\n", " "));
                            }
                            cols.Add(line.ToString());
                        }
                    }
                    var sb = new StringBuilder();
                    sb.AppendLine($"CREATE TABLE {QuoteIdent(tableName)} (");
                    for (int i = 0; i < cols.Count; i++)
                    {
                        sb.Append(cols[i]);
                        sb.AppendLine(i < cols.Count - 1 ? "," : "");
                    }
                    sb.AppendLine(");");
                    return sb.ToString();
                }
            }

            // 无注释时，尝试返回sqlite_master中的建表语句
            using (var cmd = _conn.CreateCommand())
            {
                cmd.CommandText = "SELECT sql FROM sqlite_master WHERE type='table' AND name=@n";
                cmd.Parameters.AddWithValue("@n", tableName);
                var sql = cmd.ExecuteScalar() as string;
                if (!string.IsNullOrEmpty(sql)) return sql + ";";
            }
            // fallback: 使用PRAGMA拼装
            using (var cmd = _conn.CreateCommand())
            {
                cmd.CommandText = $"PRAGMA table_info({QuoteIdent(tableName)})";
                var cols = new List<string>();
                using (var rd = cmd.ExecuteReader())
                {
                    while (rd.Read())
                    {
                        var name = rd[1]?.ToString();
                        var type = rd[2]?.ToString();
                        cols.Add($"    {QuoteIdent(name)} {type}");
                    }
                }
                var sb = new StringBuilder();
                sb.AppendLine($"CREATE TABLE {QuoteIdent(tableName)} (");
                for (int i = 0; i < cols.Count; i++)
                {
                    sb.Append(cols[i]);
                    sb.AppendLine(i < cols.Count - 1 ? "," : "");
                }
                sb.AppendLine(");");
                return sb.ToString();
            }
        }

        public void RebuildFromExcel(Dictionary<string, ExcelFile> files)
        {
            using (var tx = _conn.BeginTransaction())
            {
                DropAllUserTables();
                EnsureMetaTables();

                foreach (var file in files.Values)
                {
                    foreach (var ws in file.Worksheets.Values)
                    {
                        CreateTable(ws);
                        BulkInsert(ws);
                        InsertComments(ws);
                    }
                }

                tx.Commit();
            }
        }

        private void DropAllUserTables()
        {
            var tables = new List<string>();
            using (var cmd = _conn.CreateCommand())
            {
                cmd.CommandText = "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'";
                using (var rd = cmd.ExecuteReader())
                {
                    while (rd.Read()) tables.Add(rd.GetString(0));
                }
            }
            foreach (var t in tables)
            {
                using (var cmd = _conn.CreateCommand())
                {
                    cmd.CommandText = $"DROP TABLE IF EXISTS {QuoteIdent(t)}";
                    cmd.ExecuteNonQuery();
                }
            }
        }

        private void CreateTable(Worksheet ws)
        {
            var cols = ws.Headers.Select(h => $"{QuoteIdent(h.Name)} {MapType(h.DataType)}");
            var sql = $"CREATE TABLE {QuoteIdent(ws.Name)} (\n    {string.Join(",\n    ", cols)}\n)";
            using (var cmd = _conn.CreateCommand())
            {
                cmd.CommandText = sql;
                cmd.ExecuteNonQuery();
            }
        }

        private void InsertComments(Worksheet ws)
        {
            using (var cmd = _conn.CreateCommand())
            using (var tx = _conn.BeginTransaction())
            {
                cmd.CommandText = "INSERT OR REPLACE INTO __excel_column_comments(table_name, column_name, comment) VALUES (@t,@c,@m)";
                var pTable = new SQLiteParameter("@t");
                var pCol = new SQLiteParameter("@c");
                var pMsg = new SQLiteParameter("@m");
                cmd.Parameters.AddRange(new[] { pTable, pCol, pMsg });

                foreach (var h in ws.Headers)
                {
                    if (!string.IsNullOrEmpty(h.Comments))
                    {
                        pTable.Value = ws.Name;
                        pCol.Value = h.Name;
                        pMsg.Value = h.Comments;
                        cmd.ExecuteNonQuery();
                    }
                }
                tx.Commit();
            }
        }

        private void BulkInsert(Worksheet ws)
        {
            if (ws.DataRows == null || ws.DataRows.Count == 0) return;

            var colNames = ws.Headers.Select(h => h.Name).ToList();
            var colParams = colNames.Select((n, i) => "@p" + i).ToList();
            var insertSql = $"INSERT INTO {QuoteIdent(ws.Name)} (" + string.Join(",", colNames.Select(QuoteIdent)) + ") VALUES (" + string.Join(",", colParams) + ")";
            using (var cmd = _conn.CreateCommand())
            using (var tx = _conn.BeginTransaction())
            {
                cmd.CommandText = insertSql;

                // 预创建参数并指定类型，降低驱动推断失败的概率
                for (int i = 0; i < colParams.Count; i++)
                {
                    var p = new SQLiteParameter(colParams[i]) { DbType = MapDbType(ws.Headers[i].DataType) };
                    cmd.Parameters.Add(p);
                }

                int rowNo = 0;
                foreach (var row in ws.DataRows)
                {
                    rowNo++;
                    try
                    {
                        for (int i = 0; i < colNames.Count; i++)
                        {
                            var expectedType = ws.Headers[i].DataType;
                            row.TryGetValue(colNames[i], out var v);
                            cmd.Parameters[i].Value = ConvertForDb(v, expectedType);
                        }
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            var data = Newtonsoft.Json.JsonConvert.SerializeObject(row);
                            Console.Error.WriteLine($"[BulkInsert] 表={ws.Name} 行={rowNo} 失败: {ex.Message}\nSQL={insertSql}\n数据={data}");
                            for (int i = 0; i < colNames.Count; i++)
                            {
                                var val = row.ContainsKey(colNames[i]) ? row[colNames[i]] : null;
                                Console.Error.WriteLine($"  列[{i}] {colNames[i]}({ws.Headers[i].DataType}) = '{val}' 类型={val?.GetType().FullName ?? "null"}");
                            }
                        }
                        catch { }
                    }
                }
                tx.Commit();
            }
        }

        private object ConvertForDb(object v, string expectedType)
        {
            if (v == null || v == DBNull.Value) return DBNull.Value;
            var t = (expectedType ?? "").ToUpperInvariant();
            if (t == "INT")
            {
                if (v is bool b) return b ? 1 : 0;
                if (v is double d) return Convert.ToInt64(Math.Truncate(d));
                if (v is float f) return Convert.ToInt64(Math.Truncate(f));
                if (v is long || v is int || v is short || v is byte) return Convert.ToInt64(v);
                var s = v.ToString().Trim();
                if (string.IsNullOrEmpty(s)) return DBNull.Value;
                if (long.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var l)) return l;
                if (double.TryParse(NormalizeNumberString(s), NumberStyles.Any, CultureInfo.InvariantCulture, out var dd)) return Convert.ToInt64(Math.Truncate(dd));
                throw new FormatException($"无法将值 '{s}' 转为 INT");
            }
            if (t == "DOUBLE" || t == "FLOAT" || t == "REAL")
            {
                if (v is double || v is float || v is decimal) return Convert.ToDouble(v);
                var s = v.ToString().Trim();
                if (string.IsNullOrEmpty(s)) return DBNull.Value;
                s = TrimPercent(s);
                if (double.TryParse(NormalizeNumberString(s), NumberStyles.Any, CultureInfo.InvariantCulture, out var d)) return d;
                throw new FormatException($"无法将值 '{s}' 转为 DOUBLE");
            }
            if (t == "DATE" || t == "DATETIME")
            {
                if (v is DateTime dt) return dt.ToString("yyyy-MM-dd HH:mm:ss");
                var s = v.ToString().Trim();
                if (string.IsNullOrEmpty(s)) return DBNull.Value;
                if (DateTime.TryParse(s, out var parsed)) return parsed.ToString("yyyy-MM-dd HH:mm:ss");
                throw new FormatException($"无法将值 '{s}' 转为 DATE");
            }
            if (t == "BOOLEAN" || t == "BIT")
            {
                if (v is bool b) return b ? 1 : 0;
                var s = v.ToString().Trim().ToLowerInvariant();
                if (string.IsNullOrEmpty(s)) return DBNull.Value;
                if (s == "1" || s == "true" || s == "y" || s == "yes" || s == "是") return 1;
                if (s == "0" || s == "false" || s == "n" || s == "no" || s == "否") return 0;
                throw new FormatException($"无法将值 '{v}' 转为 BOOLEAN");
            }
            // 默认TEXT
            return v.ToString();
        }

        private static string NormalizeNumberString(string s)
        {
            // 处理常见的千分位与逗号小数
            s = s.Replace("%", "");
            var hasDot = s.Contains(".");
            var hasComma = s.Contains(",");
            if (hasComma && !hasDot)
            {
                // 用逗号作为小数点的情况
                s = s.Replace(',', '.');
            }
            // 去除千分位分隔符
            if (hasDot && hasComma)
            {
                // 假设逗号为千分位
                s = s.Replace(",", "");
            }
            return s;
        }

        private static string TrimPercent(string s)
        {
            s = s.Trim();
            if (s.EndsWith("%"))
            {
                var core = s.Substring(0, s.Length - 1).Trim();
                if (double.TryParse(NormalizeNumberString(core), NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
                {
                    return (d / 100.0).ToString(CultureInfo.InvariantCulture);
                }
            }
            return s;
        }

        public List<Dictionary<string, object>> ExecuteQuery(string sql)
        {
            using (var cmd = _conn.CreateCommand())
            {
                cmd.CommandText = sql;
                using (var rd = cmd.ExecuteReader())
                {
                    var result = new List<Dictionary<string, object>>();
                    while (rd.Read())
                    {
                        var obj = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                        for (int i = 0; i < rd.FieldCount; i++)
                        {
                            var name = rd.GetName(i);
                            var val = rd.IsDBNull(i) ? null : rd.GetValue(i);
                            obj[name] = val;
                        }
                        result.Add(obj);
                    }
                    return result;
                }
            }
        }

        public int ExecuteNonQuery(string sql)
        {
            using (var cmd = _conn.CreateCommand())
            {
                cmd.CommandText = sql;
                return cmd.ExecuteNonQuery();
            }
        }

        public IDataReader QueryTableData(string tableName, IEnumerable<string> columns)
        {
            var cols = columns?.Any() == true ? string.Join(",", columns.Select(QuoteIdent)) : "*";
            var sql = $"SELECT {cols} FROM {QuoteIdent(tableName)}";
            var cmd = _conn.CreateCommand();
            cmd.CommandText = sql;
            // 返回给调用者管理IDataReader/Command生命周期
            return cmd.ExecuteReader(CommandBehavior.CloseConnection);
        }
    }
}
