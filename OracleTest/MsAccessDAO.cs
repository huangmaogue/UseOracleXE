using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using DAO = Microsoft.Office.Interop.Access.Dao;

namespace OracleTest
{
    public class MsAccessDAO
    {
        private DAO.DBEngine dbEngine = new DAO.DBEngine();
        private DAO.Database db = null;
        private DAO.Recordset rs = null;

        /// <summary>
        /// 設定MsAccess的建構式.
        /// </summary>
        public MsAccessDAO(string dbName)
        {
            try
            {
                db = dbEngine.OpenDatabase(dbName);
            }
            catch (Exception e) { throw e; }
        }

        /// <summary>
		/// 關閉連線.
		/// </summary>
        public void ConnClose()
        {
            try
            {
                if (rs != null) 
                {
                    rs.Close();
                    rs = null;
                }
                if (db != null)
                {
                    db.Close();
                    db = null;
                }
            }
            catch (Exception e) { throw e; }
        }

        /// <summary>
        /// 查詢資料庫內容,並存入指定的 ComboBox 中
        /// </summary>
        public void ToComboBox(bool clear, System.Windows.Forms.ComboBox combobox, string sqlcmd)
        {
            try
            {
                rs = db.OpenRecordset(sqlcmd, DAO.RecordsetTypeEnum.dbOpenForwardOnly, DAO.RecordsetOptionEnum.dbReadOnly);
                while (!rs.EOF)
                {
                    combobox.Items.Add(rs.Fields[0].Value);
                    rs.MoveNext();
                }
            }
            catch (Exception e) { throw e; }
        }

        /// <summary>
        /// 查詢資料庫內容,並存入指定的 ListBox 中
        /// </summary>
        public void ToListBox(bool clear, System.Windows.Forms.ListBox listbox, string sqlcmd)
        {
            try
            {
                rs = db.OpenRecordset(sqlcmd, DAO.RecordsetTypeEnum.dbOpenForwardOnly, DAO.RecordsetOptionEnum.dbReadOnly);
                while (!rs.EOF)
                {
                    listbox.Items.Add(rs.Fields[0].Value);
                    rs.MoveNext();
                }
            }
            catch (Exception e) { throw e; }
        }

        /// <summary>
        /// 查詢資料庫內容,並存入指定的 List 中
        /// </summary>
        public void ToList(bool clear, List<string> list, string sqlcmd)
        {
            try
            {
                rs = db.OpenRecordset(sqlcmd, DAO.RecordsetTypeEnum.dbOpenForwardOnly, DAO.RecordsetOptionEnum.dbReadOnly);
                while (!rs.EOF)
                {
                    list.Add(rs.Fields[0].Value);
                    rs.MoveNext();
                }
            }
            catch (Exception e) { throw e; }
        }

        /// <summary>
        /// 查詢資料庫內容,並存入指定的 List 中
        /// </summary>
        public void ToList(bool clear, List<string>[] lists, string sqlcmd)
        {
            try
            {
                rs = db.OpenRecordset(sqlcmd, DAO.RecordsetTypeEnum.dbOpenForwardOnly, DAO.RecordsetOptionEnum.dbReadOnly);
                while (!rs.EOF)
                {
                    for (int j = 0; j < lists.Length; j++) lists[j].Add(rs.Fields[j].Value);
                    rs.MoveNext();
                }
            }
            catch (Exception e) { throw e; }
        }

        /// <summary>
        /// 查詢資料庫內容,並存入指定的 Dictionary 中
        /// </summary>
        public void ToDictionary(bool clear, Dictionary<string, string> dic, string sqlcmd)
        {
            try
            {
                rs = db.OpenRecordset(sqlcmd, DAO.RecordsetTypeEnum.dbOpenForwardOnly, DAO.RecordsetOptionEnum.dbReadOnly);
                while (!rs.EOF)
                {
                    dic.Add(rs.Fields[0].Value, rs.Fields[1].Value);
                    rs.MoveNext();
                }
            }
            catch (Exception e) { throw e; }
        }

        /// <summary>
        /// 查詢資料庫內容,並存入指定的 Dictionary 中
        /// </summary>
        public Dictionary<string, string> GetDictionary(string sqlcmd)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            try
            {
                rs = db.OpenRecordset(sqlcmd, DAO.RecordsetTypeEnum.dbOpenForwardOnly, DAO.RecordsetOptionEnum.dbReadOnly);
                while (!rs.EOF)
                {
                    dic.Add(rs.Fields[0].Value, rs.Fields[1].Value);
                    rs.MoveNext();
                }
            }
            catch (Exception e) { throw e; }
            return dic;
        }

        /// <summary>
        /// 查詢資料庫內容,並存入指定的 Dictionary 中
        /// </summary>
        public void ToDictionary(bool clear, Dictionary<string, Tuple<string, string>> dic, string sqlcmd)
        {
            try
            {
                if (clear) dic.Clear();
                rs = db.OpenRecordset(sqlcmd, DAO.RecordsetTypeEnum.dbOpenForwardOnly, DAO.RecordsetOptionEnum.dbReadOnly);
                while (!rs.EOF)
                {
                    dic.Add(rs.Fields[0].Value, new Tuple<string, string>(rs.Fields[1].Value, rs.Fields[2].Value));
                    rs.MoveNext();
                }
            }
            catch (Exception e) { throw e; }
        }

        /// <summary>
        /// 查詢資料庫內容,並存入指定的 DataTable 中
        /// </summary>
        public void ToDataTable(bool clear, DataTable table, string sqlcmd)
        {
            try
            {
                if (clear == true) table.Clear();

                rs = db.OpenRecordset(sqlcmd, DAO.RecordsetTypeEnum.dbOpenForwardOnly, DAO.RecordsetOptionEnum.dbReadOnly);
                if (!rs.EOF)
                {
                    int colSize = rs.Fields.Count;
                    string[] data = new string[colSize];
                    for (int i = 0; i < colSize; i++)
                    {
                        if (!table.Columns.Contains(rs.Fields[i].Name)) table.Columns.Add(new DataColumn(rs.Fields[i].Name, typeof(string)));
                    }

                    while (!rs.EOF)
                    {
                        for (int i = 0; i < colSize; i++)
                        {
                            data[i] = string.Empty;
                            data[i] = rs.Fields[i].Value.ToString();
                        }
                        table.Rows.Add(data);
                        rs.MoveNext();
                    }
                }
            }
            catch (Exception e) { throw e; }
        }

        /// <summary>
        /// 查詢資料庫內容,並存入指定的 DataTable 中
        /// </summary>
        public DataTable GetDataTable(string sqlcmd)
        {
            var table = new DataTable();
            try
            {
                rs = db.OpenRecordset(sqlcmd, DAO.RecordsetTypeEnum.dbOpenForwardOnly, DAO.RecordsetOptionEnum.dbReadOnly);
                int colSize = rs.Fields.Count;
                string[] data = new string[colSize];
                for (int i = 0; i < colSize; i++) table.Columns.Add(new DataColumn(rs.Fields[i].Name, typeof(string)));
                while (!rs.EOF)
                {
                    for (int i = 0; i < colSize; i++)
                    {
                        data[i] = string.Empty;
                        data[i] = rs.Fields[i].Value.ToString();
                    }
                    table.Rows.Add(data);
                    rs.MoveNext();
                }
            }
            catch (Exception e) { throw e; }
            return table;
        }

        /// <summary>
        /// 查詢資料庫內容,並存入指定的 DataTable 中
        /// </summary>
        public DataTable GetDataTable(string sqlcmd, string[] columnsExclude, int pageNumber, int rowCount, ref int allCount)
        {
            var table = new DataTable();
            try
            {
                rs = db.OpenRecordset(sqlcmd, DAO.RecordsetTypeEnum.dbOpenForwardOnly, DAO.RecordsetOptionEnum.dbReadOnly);
                int colSize = rs.Fields.Count;
                for (int i = 0; i < colSize; i++)
                {
                    if (!columnsExclude.Contains(rs.Fields[i].Name)) table.Columns.Add(new DataColumn(rs.Fields[i].Name, typeof(string)));
                }
                string[] data = new string[table.Columns.Count];
                int idx = 0;
                int idxStart = (pageNumber * rowCount) - rowCount;
                int idxEnd = pageNumber * rowCount - 1;
                while (!rs.EOF)
                {
                    if (idx >= idxStart && idx <= idxEnd)
                    {
                        for (int i = 0; i < table.Columns.Count; i++)
                        {
                            data[i] = string.Empty;
                            data[i] = rs.Fields[table.Columns[i].ColumnName].Value.ToString();
                        }
                        table.Rows.Add(data);
                    }
                    idx++;
                    //if (idx++ >= idxEnd) break;
                    rs.MoveNext();
                }
                allCount = idx;
            }
            catch (Exception e) { throw e; }
            return table;
        }

        /// <summary>
        /// 將資料庫表格 Schema, 設定至 DataTable 中
        /// </summary>
        public DataTable InitDataTable(string tableName)
        {
            var table = new DataTable();
            try
            {
                List<string> cols = GetTableColumns(tableName);
                foreach (string col in cols) table.Columns.Add(new DataColumn(col, typeof(string)));
            }
            catch (Exception e) { throw e; }
            return table;
        }

        /// <summary>
        /// 傳回 SQL Select 指令第一個欄位所有資料的陣列
        /// </summary>
        public string[] GetValues(string sqlcmd)
        {
            string[] data = null;
            try
            {
                List<string> list = new List<string>();
                ToList(false, list, sqlcmd);
                data = list.ToArray();
            }
            catch (Exception e) { throw e; }
            return data;
        }

        /// <summary>
        /// 傳回 SQL Select 指令第一個欄位所有資料的陣列
        /// </summary>
        public List<string> GetList(string sqlcmd)
        {
            List<string> list = new List<string>();
            try
            {
                rs = db.OpenRecordset(sqlcmd, DAO.RecordsetTypeEnum.dbOpenForwardOnly, DAO.RecordsetOptionEnum.dbReadOnly);
                while (!rs.EOF)
                {
                    list.Add(rs.Fields[0].Value.ToString());
                    rs.MoveNext();
                }
            }
            catch (Exception e) { throw e; }
            return list;
        }

        /// <summary>
        /// 傳回 SQL Select 指令第一個欄位及第一筆資料的值
        /// </summary>
        public string GetValue(string sqlcmd)
        {
            string data = string.Empty;

            try
            {
                rs = db.OpenRecordset(sqlcmd, DAO.RecordsetTypeEnum.dbOpenForwardOnly, DAO.RecordsetOptionEnum.dbReadOnly);
                if (!rs.EOF) data = rs.Fields[0].Value;
            }
            catch (Exception e) { throw e; }
            return data;
        }

        /// <summary>
        /// 取回 mdb 內所有的 table 名稱列表
        /// </summary>
        public List<string> GetTableNames()
        {
            List<string> list = new List<string>();
            try
            {
                rs = db.ListTables();
                while (!rs.EOF)
                {
                    list.Add(rs.Fields[0].Value);
                    rs.MoveNext();
                }
            }
            catch (Exception e) { throw e; }
            return list;
        }

        /// <summary>
        /// 取回 mdb 內 table 所有的 columns 名稱列表
        /// </summary>
        public List<string> GetTableColumns(string _TableName)
        {
            List<string> list = new List<string>();
            try
            {
                rs = db.ListFields(_TableName);
                while (!rs.EOF)
                {
                    list.Add(rs.Fields[0].Value);
                    rs.MoveNext();
                }
            }
            catch (Exception e) { throw e; }
            return list;
        }

        /// <summary>
        /// 取回 mdb 內 table 有設 Index 的 columns
        /// </summary>
        public List<string> GetTableIndexes(string _TableName)
        {
            List<string> list = new List<string>();
            try
            {
                rs = db.OpenTable(_TableName).ListIndexes();
                while (!rs.EOF)
                {
                    list.Add(rs.Fields[0].Value);
                    rs.MoveNext();
                }
            }
            catch (Exception e) { throw e; }
            return list;
        }

        /// <summary>
        /// 取回 mdb 內 table 有設 Index 的 columns
        /// </summary>
        public List<string> GetTableIndexes(string _TableName, string[] _Exclude)
        {
            List<string> list = new List<string>();
            try
            {
                List<string> _ExcludeList = new List<string>(_Exclude);
                rs = db.OpenTable(_TableName).ListIndexes();
                while (!rs.EOF)
                {
                    string idxName = rs.Fields[0].Value.ToString();
                    if (!_ExcludeList.Contains(idxName)) list.Add(idxName);
                    rs.MoveNext();
                }
            }
            catch (Exception e) { throw e; }
            return list;
        }

        /// <summary>
        /// 執行 SQL Insert, Update, Delete 指令
        /// </summary>
        public string NonQuery(string sqlcmd)
        {
            try
            {
                db.BeginTrans();
                //db.Execute(sqlcmd, DAO.RecordsetOptionEnum.dbDenyWrite);
                db.Execute(sqlcmd);
                db.CommitTrans();
                return "Succeed";
            }
            catch(Exception e)
            {
                db.Rollback();
                Debug.WriteLine(e.Message);
                FileRW.Write(true, Environment.CurrentDirectory + @"\NxESL_Log.txt", $"\r\n{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")} : NonQuery Error : {sqlcmd}");
                return "Failed";
            }
        }

        /// <summary>
        /// 執行 SQL Insert, Update, Delete 指令
        /// </summary>
        public string NonQuery(string[] sqlcmds)
        {
            try
            {
                db.BeginTrans();
                foreach (string sql in sqlcmds) if (!string.IsNullOrWhiteSpace(sql)) db.Execute(sql);
                db.CommitTrans();
            }
            catch
            {
                db.Rollback();
                return "Failed";
            }
            return "Succeed";
        }

        //https://msdn.microsoft.com/zh-tw/library/93ehy0z8(v=vs.110).aspx
        public void ExecuteTransaction(string[] sqlcmds)
        {
            using (OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Z:\20170814_165922.accdb;Mode=16"))
            {
                OleDbCommand command = new OleDbCommand();
                OleDbTransaction transaction = null;

                // Set the Connection to the new OleDbConnection.
                command.Connection = connection;

                // Open the connection and execute the transaction.
                try
                {
                    connection.Open();

                    // Start a local transaction with ReadCommitted isolation level.
                    transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

                    // Assign transaction object for a pending local transaction.
                    command.Connection = connection;
                    command.Transaction = transaction;


                    foreach (string sql in sqlcmds)
                    {
                        if (!string.IsNullOrWhiteSpace(sql))
                        {
                            command.CommandText = sql;
                            command.ExecuteNonQuery();
                        }
                    }

                    // Commit the transaction.
                    transaction.Commit();
                    Console.WriteLine("Both records are written to database.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    try
                    {
                        // Attempt to roll back the transaction.
                        transaction.Rollback();
                    }
                    catch
                    {
                        // Do nothing here; transaction is not active.
                    }
                }
                // The connection is automatically closed when the
                // code exits the using block.
            }
        }

        /*
        string updatequery = @"UPDATE [table] SET [Last10Attempts] = ?, [Last10AttemptsSum] = ?, [total-question-attempts] = ? WHERE id = ? ";
        using(OleDbConnection con = new OleDbConnection(Provider = Microsoft.ACE.OLEDB.12.0; Data Source = C:\\Users\\Kalpesh\\Desktop\\Entry.accdb))
        using(OleDbCommand cmd = new OleDbCommand(updatequery, con))
        {
            con.Open();
            cmd.Parameters.AddWithValue("Last10Attempts", last10attempts);
            cmd.Parameters.AddWithValue("Last10AttemptsSum", counter);
            cmd.Parameters.AddWithValue("total-question-attempts", questionattempts + 1);
            cmd.Parameters.AddWithValue("ID", currentid + 1);
            cmd.ExecuteNonQuery();
        }
        */

        /// <summary>
        /// 傳回以現在時間減去間隔時間後的時間
        /// </summary>
        public DateTime DateTimeAgo(int days, int hours, int minutes, int seconds)
        {
            DateTime A = DateTime.Now;
            DateTime B = A.Subtract(new TimeSpan(days, hours, minutes, seconds));
            return B;
        }

        /// <summary>
        /// 傳回以指定時間減去間隔時間後的時間
        /// </summary>
        public DateTime DateTimeAgo(DateTime goal, int days, int hours, int minutes, int seconds)
        {
            DateTime B = goal.Subtract(new TimeSpan(days, hours, minutes, seconds));
            return B;
        }
    }
}
