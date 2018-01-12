using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using DAO = Microsoft.Office.Interop.Access.Dao;

//Queries performed from within the Microsoft Access application itself normally use * and ? as wildcard characters for the LIKE operator. 
//OleDb connections to an Access database from an external application should use the % and _ wildcard characters instead.
namespace OracleTest
{   /// <summary>
    /// 針對 Office Access 所使用的常用函式庫.
    /// </summary>
    public class MsAccess
    {
        private string dbName = string.Empty;
        private string mdbPswd = null;
        private OleDbConnection conn = new OleDbConnection();
        //private OleDbCommand cmd = new OleDbCommand();
        //private OleDbDataAdapter adp = new OleDbDataAdapter();
        //private OleDbDataReader sqlreader;

        /// <summary>
        /// 設定MsAccess的建構式.
        /// </summary>
        //public MsAccess()
        //{
        //}

        /// <summary>
        /// 設定MsAccess的建構式.
        /// </summary>
        public MsAccess(string dbName, string mdbPswd, bool isExclusive)
        {
            try
            {
                this.dbName = dbName;
                this.mdbPswd = mdbPswd;
                Conn(dbName, mdbPswd, isExclusive);
            }
            catch (Exception e) { throw e; }
        }

        /// <summary>
        /// 連線至指定的伺服器及資料庫.
        /// </summary>
        public bool Conn(string dbName, string mdbPswd, bool isExclusive)
        {
            //// Share Mode=16 - use if multiple users must have simultaneous access to the db
            //string constr = @"Provider=Microsoft.ACE.OLEDB.12.0;Mode=16;Data Source=C:\...\RSLogixDB.accdb;user id=;password=;";

            //// Share Mode=12 - exclusive mode if only your app needs access to the db
            //string constr = @"Provider=Microsoft.ACE.OLEDB.12.0;Mode=12;Data Source=C:\...\RSLogixDB.accdb;user id=;password=;";

            int mode;
            this.dbName = dbName;
            this.mdbPswd = mdbPswd;
            bool check = false;
            if (isExclusive) mode = 12; else mode = 16;
            try
            {
                if (conn.State.ToString() == "Open")
                {
                    if (conn.DataSource == dbName) check = true;
                    else conn.Close();
                }
                if (conn.State.ToString() == "Closed")
                {
                    if (dbName.IndexOf(".accdb") > -1)
                    {
                        if (string.IsNullOrEmpty(mdbPswd)) conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbName + ";Mode=" + mode;
                        else conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbName + ";Mode=" + mode + ";Jet OLEDB:Database Password=" + mdbPswd;
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(mdbPswd)) conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dbName + ";Mode=" + mode;
                        else conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dbName + ";Mode=" + mode + ";Jet OLEDB:Database Password=" + mdbPswd;
                    }
                    //cmd.Connection = conn;
                    conn.Open();
                    check = true;
                }
            }
            catch (Exception e) { throw e; }
            return check;
        }

        /// <summary>
		/// 關閉連線.
		/// </summary>
        public void ConnClose()
        {
            try
            {
                if (conn.State.ToString() == "Open")
                {
                    conn.Close();
                    //cmd.Dispose();
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
                //using (OleDbConnection connection = new OleDbConnection(connectionString))
                //using (OleDbCommand command = new OleDbCommand(query, connection))
                using (OleDbCommand cmd = new OleDbCommand(sqlcmd, conn))
                using (OleDbDataAdapter adp = new OleDbDataAdapter(cmd))
                {
                    DataTable table = new DataTable();
                    combobox.Items.Clear();
                    combobox.Text = "";
                    //cmd.CommandText = sqlcmd;
                    //adp.SelectCommand = cmd;
                    adp.Fill(table);
                    for (int i = 0; i < table.Rows.Count; i++) combobox.Items.Add(table.Rows[i][0]);
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
                using (OleDbCommand cmd = new OleDbCommand(sqlcmd, conn))
                using (OleDbDataAdapter adp = new OleDbDataAdapter(cmd))
                {
                    DataTable table = new DataTable();
                    listbox.Items.Clear();
                    listbox.Text = "";
                    //cmd.CommandText = sqlcmd;
                    //adp.SelectCommand = cmd;
                    adp.Fill(table);
                    for (int i = 0; i < table.Rows.Count; i++) listbox.Items.Add(table.Rows[i][0]);
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
                using (OleDbCommand cmd = new OleDbCommand(sqlcmd, conn))
                using (OleDbDataAdapter adp = new OleDbDataAdapter(cmd))
                {
                    DataTable table = new DataTable();
                    list.Clear();
                    //cmd.CommandText = sqlcmd;
                    //adp.SelectCommand = cmd;
                    adp.Fill(table);
                    for (int i = 0; i < table.Rows.Count; i++) list.Add(table.Rows[i][0].ToString());
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
                using (OleDbCommand cmd = new OleDbCommand(sqlcmd, conn))
                using (OleDbDataAdapter adp = new OleDbDataAdapter(cmd))
                {
                    DataTable table = new DataTable();
                    if (clear) foreach (var list in lists) list.Clear();
                    //cmd.CommandText = sqlcmd;
                    //adp.SelectCommand = cmd;
                    adp.Fill(table);
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        for (int j=0;j<lists.Length;j++) lists[j].Add(table.Rows[i][j].ToString());
                    }
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
                using (OleDbCommand cmd = new OleDbCommand(sqlcmd, conn))
                using (OleDbDataAdapter adp = new OleDbDataAdapter(cmd))
                {
                    DataTable table = new DataTable();
                    if (clear) dic.Clear();
                    //cmd.CommandText = sqlcmd;
                    //adp.SelectCommand = cmd;
                    adp.Fill(table);
                    for (int i = 0; i < table.Rows.Count; i++) dic.Add(table.Rows[i][0].ToString(), table.Rows[i][1].ToString());
                }
            }
            catch (Exception e) { throw e; }
        }

        /// <summary>
        /// 查詢資料庫內容,並存入指定的 Dictionary 中
        /// </summary>
        public void ToDictionary(bool clear, Dictionary<string, Tuple<string, string>> dic, string sqlcmd)
        {
            try
            {
                using (OleDbCommand cmd = new OleDbCommand(sqlcmd, conn))
                using (OleDbDataAdapter adp = new OleDbDataAdapter(cmd))
                {
                    DataTable table = new DataTable();
                    if (clear) dic.Clear();
                    //cmd.CommandText = sqlcmd;
                    //adp.SelectCommand = cmd;
                    adp.Fill(table);
                    for (int i = 0; i < table.Rows.Count; i++) dic.Add(table.Rows[i][0].ToString(), new Tuple<string, string>(table.Rows[i][1].ToString(), table.Rows[i][2].ToString()));
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
                using (OleDbCommand cmd = new OleDbCommand(sqlcmd, conn))
                using (OleDbDataAdapter adp = new OleDbDataAdapter(cmd))
                {
                    if (clear == true) table.Clear();
                    //cmd.CommandText = sqlcmd;
                    //adp.SelectCommand = cmd;
                    adp.Fill(table);
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
                using (OleDbCommand cmd = new OleDbCommand(sqlcmd, conn))
                using (OleDbDataAdapter adp = new OleDbDataAdapter(cmd))
                {
                    //cmd.CommandText = sqlcmd;
                    //adp.SelectCommand = cmd;
                    adp.Fill(table);
                }
            }
            catch (Exception e) { throw e; }
            return table;
        }

        /// <summary>
        /// 傳回 SQL Select 指令第一個欄位所有資料的陣列
        /// </summary>
        public object[] GetValues(string sqlcmd)
        {
            object[] data = null;
            try
            {
                using (OleDbCommand cmd = new OleDbCommand(sqlcmd, conn))
                using (OleDbDataAdapter adp = new OleDbDataAdapter(cmd))
                {
                    DataTable table = new DataTable();
                    //cmd.CommandText = sqlcmd;
                    //adp.SelectCommand = cmd;
                    adp.Fill(table);
                    data = new object[table.Rows.Count];
                    for (int i = 0; i < table.Rows.Count; i++) data[i] = table.Rows[i][0];
                }
            }
            catch (Exception e) { throw e; }
            return data;
        }

        /// <summary>
        /// 傳回 SQL Select 指令第一個欄位及第一筆資料的值
        /// </summary>
        public string GetValue(string sqlcmd)
        {
            string data = string.Empty;
               
            try
            {
                using (OleDbCommand cmd = new OleDbCommand(sqlcmd, conn))
                using (OleDbDataAdapter adp = new OleDbDataAdapter(cmd))
                {
                    //cmd.CommandText = sqlcmd;
                    OleDbDataReader sqlreader = cmd.ExecuteReader(CommandBehavior.SingleRow);
                    if (sqlreader.Read() == true) data = sqlreader[0].ToString();
                    sqlreader.Close();
                    sqlreader = null;
                }
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
                DataTable dt = conn.GetSchema("Tables");
                //DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                foreach (DataRow row in dt.Rows) if (row["TABLE_TYPE"].ToString() == "TABLE") list.Add(row["TABLE_NAME"].ToString());

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
                DataTable dt = conn.GetSchema("Columns", new string[] { null, null, _TableName }).AsEnumerable().OrderBy(row => row["ORDINAL_POSITION"]).CopyToDataTable();
                for (int i = 0; i < dt.Rows.Count; i++) list.Add(dt.Rows[i]["COLUMN_NAME"].ToString());
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
                DataTable dt = conn.GetSchema("Indexes").AsEnumerable().Where(w => w.Field<string>("TABLE_NAME").Equals(_TableName)).CopyToDataTable();
                for (int i = 0; i < dt.Rows.Count; i++) list.Add(dt.Rows[i]["COLUMN_NAME"].ToString());
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
                DataTable dt = conn.GetSchema("Indexes").AsEnumerable().Where(w => w.Field<string>("TABLE_NAME").Equals(_TableName)).CopyToDataTable();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (!_ExcludeList.Contains(dt.Rows[i]["COLUMN_NAME"].ToString())) list.Add(dt.Rows[i]["COLUMN_NAME"].ToString());
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
                using (OleDbCommand cmd = new OleDbCommand(sqlcmd, conn))
                {
                    //cmd.CommandText = sqlcmd;
                    cmd.ExecuteNonQuery();
                    return "OK";
                }

            }
            catch
            {
                FileRW.Write(true, Environment.CurrentDirectory + @"\NxESL_Log.txt", $"\r\n{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")} : NonQuery Error : {sqlcmd}");
                return "FAIL";
            }
        }

        /// <summary>
        /// 執行 SQL Insert, Update, Delete 指令
        /// </summary>
        public string NonQuery(string[] sqlcmds)
        {
            using (var trans = conn.BeginTransaction())
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.Transaction = trans;
                try
                {
                    foreach (string sql in sqlcmds)
                    {
                        if (!string.IsNullOrWhiteSpace(sql))
                        {
                            cmd.Connection = conn;
                            cmd.CommandText = sql;
                            cmd.ExecuteNonQuery();
                        }
                    }
                    trans.Commit();
                }
                catch(Exception e)
                {
                    FileRW.Write(true, Environment.CurrentDirectory + @"\NxESL_Log.txt", $"\r\n{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")} : NonQuery Array Error.");
                    trans.Rollback();
                    return e.Message;
                }
            }
            return "OK";
        }

        /// <summary>
        /// 執行 SQL Insert, Update, Delete 指令
        /// </summary>
        public string NonQueryDAO(string dbName, string[] sqlcmds)
        {
            DAO.DBEngine dbe = new DAO.DBEngine();
            DAO.Database db = dbe.OpenDatabase(dbName);

            try
            {
                db.BeginTrans();
                foreach (string sql in sqlcmds) if (!string.IsNullOrWhiteSpace(sql)) db.Execute(sql);
                db.CommitTrans();
            }
            catch
            {
                db.Rollback();
                return "FAIL";
            }
            return "OK";
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

        /// <summary>
        /// 傳回連線的 Server 名稱
        /// </summary>
        public string Server() { return conn.DataSource; }

        /// <summary>
        /// 傳回連線的 Database 名稱
        /// </summary>
        public string Database() { return conn.Database; }
    }
}