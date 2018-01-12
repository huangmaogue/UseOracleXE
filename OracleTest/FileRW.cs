using System;
using System.IO;
using System.Collections;
using System.ComponentModel;
using System.Net;
using System.Configuration;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;
using LumenWorks.Framework.IO.Csv;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace OracleTest
{
    /// <summary>
    /// 針對檔案處理所使用的常用工具.
    /// </summary>
    public class FileRW
    {
        private static DirectoryInfo myDir;
        private static FileInfo myFile;
        private static FileStream myFileStream;
        private static StreamWriter myStreamWriter;
        //private static StreamReader myStreamReader;
        //private FileWebRequest fileWebReq;
        //private FileWebResponse fileWebRespon;
        //private static HttpWebRequest httpReq;
        //private static HttpWebResponse httpRespon;

        //public FileRW()	{}

        /// <summary>
        /// This function is used to check specified file being used or not
        /// </summary>
        /// <param name="file">FileInfo of required file</param>
        /// <returns>If that specified file is being processed 
        /// or not found is return true</returns>
        public static Boolean IsFileLocked(string filename)
        {
            FileInfo file = new FileInfo(filename);
            FileStream stream = null;
            try
            {
                //Don't change FileAccess to ReadWrite, 
                //because if a file is in readOnly, it fails.
                stream = file.Open
                (
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.None
                );

            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                Debug.WriteLine("File Locked.");
                return true;
            }
            finally
            {
                if (stream != null) stream.Close();
                file = null;
            }
            //file is not locked
            return false;
        }

        /// <summary>
        /// 檢查目錄是否存在.
        /// </summary>
        public static bool DirCheck(string full_path_dirname)
        {
            if (Directory.Exists(full_path_dirname)) return true;
            else return false;
            //			myDir = new DirectoryInfo(full_path_dirname);
            //			if (myDir.Exists == false) return false;
            //			else return true;
        }

        /// <summary>
        /// 建立一個目錄, 如果目錄已經存在則動作取消.
        /// </summary>
        public static void DirCreate(string full_path_dirname)
        {
            if (DirCheck(full_path_dirname) == false)
            {
                myDir = Directory.CreateDirectory(full_path_dirname);
                myDir.Refresh();
            }
        }

        /// <summary>
        /// 檢查檔案是否存在.
        /// </summary>
        public static bool Check(string full_path_filename)
        {
            myFile = new FileInfo(full_path_filename);
            if (myFile.Exists == false) return false;
            else return true;
        }

        /// <summary>
        /// 清空檔案內容
        /// </summary>
        public static void Clear(string full_path_filename)
        {
            myStreamWriter = null;

            try
            {
                myStreamWriter = new StreamWriter(full_path_filename, false, System.Text.Encoding.Default);
                myStreamWriter.Write("");

            }
            catch (Exception fileExp)
            {
                throw new Exception("Clear failure : " + fileExp.Message);
            }
            finally
            {
                myStreamWriter.Flush();
                myStreamWriter.Close();
            }
        }

        /// <summary>
        /// 建立一個指定檔名的空檔案, 如果檔案已經存在則動作取消.
        /// </summary>
        public static void Create(string full_path_filename)
        {
            myFile = new FileInfo(full_path_filename);
            if (myFile.Exists == false)
            {
                myFileStream = myFile.Create();
                myFileStream.Close();
            }
            myFile.Refresh();
        }

        /// <summary>
        /// 刪除一個指定檔名的空檔案.
        /// </summary>
        public static void Delete(string full_path_filename)
        {
            myFile = new FileInfo(full_path_filename);
            if (myFile.Exists == true) { myFile.Delete(); }
            myFile.Refresh();
        }

        /// <summary>
        /// 將字串寫入指定的檔案中, 如果該檔案已存在且有資料, 則字串在最後一行寫入. 
        /// 文字檔的換行符號可用 "line 1\r\nline 2" or "line 1" + Environment.NewLine + "line 2".
        /// </summary>
        public static void Write(bool isNewLine, string full_path_filename, string text)
        {
            Create(full_path_filename);
            //myStreamReader = new StreamReader(full_path_filename);
            //ArrayList str_array = new ArrayList();
            //string str_read;
            //do
            //{
            //	str_read = myStreamReader.ReadLine();
            //	if (str_read != null) str_array.Add(str_read);
            //}
            //while (str_read != null);
            //myStreamReader.Close();
            //myStreamWriter = new StreamWriter(full_path_filename);
            //foreach (object obj in str_array) myStreamWriter.WriteLine(obj.ToString());

            myFile = new FileInfo(full_path_filename); //可以縮成一行來執行
            myStreamWriter = myFile.AppendText();      //myStreamWriter = File.AppendText(full_path_filename);
            if (isNewLine) myStreamWriter.WriteLine(text);
            else myStreamWriter.Write(text);
            myStreamWriter.Flush();
            myStreamWriter.Close();
        }

        /// <summary>
        /// 搬移檔案.
        /// </summary>
        public static void Move_File(string source_filename, string destination_filename)
        {
            myFile = new FileInfo(source_filename);
            if (myFile.Exists == true) myFile.MoveTo(destination_filename);
        }

        /// <summary>
        /// 複製檔案.
        /// </summary>
        public static void Copy_File(bool overwrite, string source_filename, string destination_filename)
        {
            File.Copy(source_filename, destination_filename, overwrite);
            //			myFile = new FileInfo(source_filename);
            //			if (myFile.Exists == true && overwrite == true)
            //			{
            //				myFile.CopyTo(destination_filename,overwrite);
            //				myFile.Refresh();
            //			}
        }

        public static string[] File2Array(string file)
        {
            List<string> list = new List<string>();
            using (StreamReader f = new StreamReader(file, Encoding.UTF8))
            {
                string line;
                while ((line = f.ReadLine()) != null)
                {
                    list.Add(line);
                }
            }
            return list.ToArray();
        }

        /// <summary>
        /// Return excel sheets names.
        /// </summary>
        public static string[] GetExcelSheets(string excelFile)
        {
            OleDbConnection oleConn = new OleDbConnection();
            OleDbDataAdapter oleAdp = new OleDbDataAdapter();
            OleDbCommand oleCmd = new OleDbCommand();

            //oleConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+excelFile+";Extended Properties=\"Excel 8.0\"";
            oleConn.ConnectionString = ExcelConnectionString(excelFile);
            oleCmd.Connection = oleConn;
            oleConn.Open();

            List<string> _SheetList = new List<string>();
            DataTable excelTB = new DataTable();
            string sheetName = string.Empty;
            excelTB = oleConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new Object[] { null, null, null, "TABLE" });
            for (int i = 0; i < excelTB.Rows.Count; i++)
            {
                sheetName = excelTB.Rows[i]["TABLE_NAME"].ToString();
                //if (sheetName.IndexOf('$') > -1) _SheetList.Add(sheetName.Substring(0, sheetName.Length - 1));
                if (sheetName.IndexOf('$') > -1) _SheetList.Add(sheetName);
            }
            oleConn.Close();
            return _SheetList.ToArray();
        }

        /// <summary>
		/// Return excel sheet column names.
		/// </summary>
		public static string[] GetExcelColumns(string excelFile, string excelSheet)
        {
            OleDbConnection oleConn = new OleDbConnection();
            OleDbDataAdapter oleAdp = new OleDbDataAdapter();
            OleDbCommand oleCmd = new OleDbCommand();

            //oleConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+excelFile+";Extended Properties=\"Excel 8.0\"";
            //oleConn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFile + ";Extended Properties=\"Excel 12.0 Xml;HDR=Yes;IMEX=1\"";
            oleConn.ConnectionString = ExcelConnectionString(excelFile);
            oleCmd.Connection = oleConn;
            oleConn.Open();

            DataTable excelTB = new DataTable();
            string columnName = string.Empty;
            excelTB = oleConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Columns, new Object[] { null, null, excelSheet, null });
            //for (int i = 0; i < excelTB.Columns.Count; i++)
            //{
            //    Console.WriteLine(excelTB.Columns[i].ToString());
            //}
            string[] _Columns = new string[excelTB.Rows.Count];
            for (int i = 0; i < excelTB.Rows.Count; i++)
            {
                _Columns[i] = excelTB.Rows[i]["COLUMN_NAME"].ToString();
            }
            oleConn.Close();
            return _Columns;
        }

        /// <summary>
        /// Return CSV column names.
        /// </summary>
        public static string[] GetCSVColumns(string _FileName, Encoding _Encoding)
        {
            string[] _Columns;
            using (CsvReader csv = new CsvReader(new StreamReader(@"d:\pos.csv", _Encoding), true))
            {

                _Columns = csv.GetFieldHeaders();

                //int fieldCount = csv.FieldCount;
                //while (csv.ReadNextRecord())
                //{
                //    for (int i = 0; i < fieldCount; i++)
                //        Console.Write(string.Format("{0} = {1};", headers[i], csv[i]));
                //    Console.WriteLine();
                //}
            }
            return _Columns;
        }

        /// <summary>
        /// 將 CSV 檔案 Import 到指定的 DataTable 中
        /// </summary>
        public static void CSV2DataTable(bool clear, DataTable _DataTable, string _FileName, Encoding _Encoding)
        {
            if (clear) _DataTable.Clear();
            using (CachedCsvReader csv = new CachedCsvReader(new StreamReader(_FileName, _Encoding), true))
            {
                _DataTable.Load(csv);
            }
        }

        /// <summary>
        /// 將 Excel 檔案 Import 到指定的 DataTable 中
        /// </summary>
        public static void Excel2DataTable(bool clear, DataTable _DataTable, string excelFile, string excelSheet)
        {
            OleDbConnection oleConn = new OleDbConnection();
            OleDbDataAdapter oleAdp = new OleDbDataAdapter();
            OleDbCommand oleCmd = new OleDbCommand();

            //Syntax: Provider=Microsoft.Jet.OLEDB.4.0;Data Source=<Full Path of Excel File>; Extended Properties="Excel 8.0; HDR=No; IMEX=1".
            //
            //Definition of Extended Properties: 
            //Excel = <No> 
            //One should specify the version of Excel Sheet here. For Excel 2000 and above, it is set it to Excel 8.0 and for all others, it is Excel 5.0.
            //
            //HDR= <Yes/No> 
            //This property will be used to specify the definition of header for each column. If the value is ‘Yes’, the first row will be treated as heading. Otherwise, the heading will be generated by the system like F1, F2 and so on.
            //
            //IMEX= <0/1/2> 
            //IMEX refers to IMport EXport mode. This can take three possible values.
            //
            //IMEX=0 and IMEX=2 will result in ImportMixedTypes being ignored and the default value of ‘Majority Types’ is used. In this case, it will take the first 8 rows and then the data type for each column will be decided. 
            //IMEX=1 is the only way to set the value of ImportMixedTypes as Text. Here, everything will be treated as text. 
            //補充
            //"HDR=Yes;" indicates that the first row contains columnnames, not data
            //"IMEX=1;" tells the driver to always read "intermixed" data columns as text
            //"IMEX=2;" seems can update excel data
            //發生過即使 IMEX=1, Import 的資料, 如果第一行某欄位的值為數值, 則以下其他行也會被當成數值對待, 這時如果有類似 A1, AAA 之類的資料, Import 進來會變成 <Null> , 這時可將非數值的資料行搬移到第一行, 就可解決.
            //oleConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + excelFile + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1\"";
            oleConn.ConnectionString = ExcelConnectionString(excelFile);
            oleCmd.Connection = oleConn;
            oleConn.Open();

            //SQL syntax: "SELECT * FROM [sheet1$]" - i.e. worksheet name followed by a "$" and wrapped in "[" "]" brackets.
            oleCmd.CommandText = "select * from [" + excelSheet + "]";
            oleAdp.SelectCommand = oleCmd;
            if (clear) _DataTable.Clear();
            try { oleAdp.Fill(_DataTable); } catch (Exception e) { System.Windows.Forms.MessageBox.Show(e.Message); }
            oleConn.Close();
        }
        //		另一簡單版本
        //		public static void ExcelToDataTable(DataTable table, string excelFile, string excelSheet)
        //		{
        //			OleDbConnection oleConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+excelFile+";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1\"");
        //			oleConn.Open();
        //			OleDbDataAdapter oleAdp = new OleDbDataAdapter("select * from ["+excelSheet+"$]",oleConn);
        //			oleAdp.Fill(table);	
        //		}

        /// <summary>
        /// 將 Excel 檔案 Import 到指定的 DataTable 中, SQL syntax: "SELECT * FROM [sheet1$]"
        /// </summary>
        public static void ExcelSelect(bool clear, DataTable table, string excelFile, string sqlcmd)
        {
            OleDbConnection oleConn = new OleDbConnection();
            OleDbDataAdapter oleAdp = new OleDbDataAdapter();
            OleDbCommand oleCmd = new OleDbCommand();

            //"HDR=Yes;" indicates that the first row contains columnnames, not data
            //"IMEX=1;" tells the driver to always read "intermixed" data columns as text
            //oleConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + excelFile + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1\"";
            oleConn.ConnectionString = ExcelConnectionString(excelFile);
            oleCmd.Connection = oleConn;
            oleConn.Open();

            //SQL syntax: "SELECT * FROM [sheet1$]" - i.e. worksheet name followed by a "$" and wrapped in "[" "]" brackets.
            oleCmd.CommandText = sqlcmd;
            oleAdp.SelectCommand = oleCmd;
            if (clear) table.Clear();
            oleAdp.Fill(table);
            oleConn.Close();
        }

        /// <summary>
        /// 將 Excel 檔案 Import 到指定的 ArrayList 中, SQL syntax: "SELECT * FROM [sheet1$]"
        /// </summary>
        public static void ExcelSelect(bool clear, ArrayList arraylist, string excelFile, string sqlcmd)
        {
            OleDbConnection oleConn = new OleDbConnection();
            OleDbDataAdapter oleAdp = new OleDbDataAdapter();
            OleDbCommand oleCmd = new OleDbCommand();

            //"HDR=Yes;" indicates that the first row contains columnnames, not data
            //"IMEX=1;" tells the driver to always read "intermixed" data columns as text
            //oleConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + excelFile + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1\"";
            oleConn.ConnectionString = ExcelConnectionString(excelFile);
            oleCmd.Connection = oleConn;
            oleConn.Open();

            //SQL syntax: "SELECT * FROM [sheet1$]" - i.e. worksheet name followed by a "$" and wrapped in "[" "]" brackets.
            oleCmd.CommandText = sqlcmd;
            oleAdp.SelectCommand = oleCmd;
            DataTable table = new DataTable();
            oleAdp.Fill(table);

            if (clear) arraylist.Clear();
            for (int i = 0; i < table.Rows.Count; i++) { arraylist.Add(table.Rows[i][0].ToString()); }
            oleConn.Close();
        }

        /// <summary>
        /// syntax: command.CommandText = "INSERT INTO [excelSheet$] ([Column1], [Column2]) VALUES(4,\"Tampa\")";
        /// </summary>
        public static void DataToExcel(string excelFile, string command)
        {

            OleDbConnection oleConn = new OleDbConnection();
            OleDbDataAdapter oleAdp = new OleDbDataAdapter();
            OleDbCommand oleCmd = new OleDbCommand();

            //"HDR=Yes;" indicates that the first row contains columnnames, not data
            //"IMEX=1;" tells the driver to always read "intermixed" data columns as text
            //oleConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + excelFile + ";Extended Properties=\"Excel 8.0\"";
            //oleConn.ConnectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={excelFile};Mode=ReadWrite;Extended Properties=\"Excel 12.0;HDR=Yes\"";
            oleConn.ConnectionString = ExcelConnectionString(excelFile);
            oleCmd.Connection = oleConn;
            oleConn.Open();

            oleCmd = oleConn.CreateCommand();
            oleCmd.CommandText = command;
            oleCmd.ExecuteNonQuery();
            oleConn.Close();
        }

        private static string ExcelConnectionString(string _FileName)
        {
            Dictionary<string, string> props = new Dictionary<string, string>();
            if (_FileName.ToLower().IndexOf(".xlsx") > -1)
            {
                // XLSX - Excel 2007, 2010, 2012, 2013
                props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
                //props["Extended Properties"] = "Excel 12.0 Xml;";
                props["Extended Properties"] = "\"Excel 12.0;HDR=Yes\"";
            }
            else
            {
                // XLS - Excel 2003 and Older
                props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
                //props["Extended Properties"] = "Excel 8.0;";
                props["Extended Properties"] = "\"Excel 8.0;HDR=Yes\"";
            }
            props["Data Source"] = _FileName;

            StringBuilder sb = new StringBuilder();
            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }
            return sb.ToString();
        }

        public static void CreateExcel(string file, string nameSheet, string[] columns)
        {
            try
            {
                Excel.Application excel;
                Excel.Workbook book;
                Excel.Worksheet sheet;

                //Start Excel and get Application object.
                excel = new Excel.Application();
                excel.Visible = false;
                excel.UserControl = false;
                excel.DisplayAlerts = false;

                //Get a new workbook.
                book = excel.Workbooks.Add();
                sheet = book.ActiveSheet;
                sheet.Name = nameSheet;

                //Add table headers going cell by cell.
                for (int i = 0; i < columns.Length; i++)
                {
                    sheet.Cells[1, i + 1] = columns[i];
                    sheet.Columns[i + 1].ColumnWidth = 18;
                }

                //Format A1:D1 as bold, vertical alignment = center.
                sheet.get_Range("A1", "Z1").Font.Bold = true;
                //sheet.get_Range("A1", "Z1").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                book.SaveAs(file, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                book.Close();
            }
            catch(Exception e)
            {
                Debug.WriteLine(e.Message);
            }
        }

        private void ExcelDEMO()
        {
            Excel.Application excel;
            Excel.Workbook book;
            Excel.Worksheet sheet;
            Excel.Range range;

            //Start Excel and get Application object.
            excel = new Excel.Application();
            excel.Visible = false;
            excel.UserControl = false;
            excel.DisplayAlerts = false;

            //Get a new workbook.
            book = excel.Workbooks.Add();
            sheet = book.ActiveSheet;
            sheet.Name = "TEST";

            //Add table headers going cell by cell.
            sheet.Cells[1, 1] = "First Name";
            sheet.Cells[1, 2] = "Last Name";
            sheet.Cells[1, 3] = "Full Name";
            sheet.Cells[1, 4] = "Salary";

            //Format A1:D1 as bold, vertical alignment = center.
            sheet.get_Range("A1", "D1").Font.Bold = true;
            sheet.get_Range("A1", "D1").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

            // Create an array to multiple values at once.
            string[,] saNames = new string[5, 2];

            saNames[0, 0] = "John";
            saNames[0, 1] = "Smith";
            saNames[1, 0] = "Tom";

            saNames[4, 1] = "Johnson";

            //Fill A2:B6 with an array of values (First and Last Names).
            sheet.get_Range("A2", "B6").Value2 = saNames;

            //Fill C2:C6 with a relative formula (=A2 & " " & B2).
            range = sheet.get_Range("C2", "C6");
            range.Formula = "=A2 & \" \" & B2";

            //Fill D2:D6 with a formula(=RAND()*100000) and apply format.
            range = sheet.get_Range("D2", "D6");
            range.Formula = "=RAND()*100000";
            range.NumberFormat = "$0.00";

            //AutoFit columns A:D.
            range = sheet.get_Range("A1", "D1");
            range.EntireColumn.AutoFit();

            book.SaveAs(@"z:\test5.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            book.Close();
        }

        //		public void Updater()
        //		{
        //			Uri uri = new Uri(System.Configuration.ConfigurationSettings.AppSettings["Updater_Path"]+@"/"+System.Configuration.ConfigurationSettings.AppSettings["File_1"]);
        //			//Uri uri = new Uri(@"http://192.168.28.67/murray_updater/SPCFDS_Monitor.exe");
        //			//CredentialCache myCred = new CredentialCache();
        //			//myCred.Add(uri,"Basic",new NetworkCredential("APCEDT","APCREPORT"));
        //			//myCred.Add(uri,"Digest", new NetworkCredential("APCEDT","APCREPORT"));
        //			//httpReq.Credentials = myCred;
        //			httpReq = (HttpWebRequest)WebRequest.Create(uri);			
        //			httpRespon = (HttpWebResponse)httpReq.GetResponse();
        //			Stream stream = httpRespon.GetResponseStream();
        //			BufferedStream buffRead = new BufferedStream(stream);
        //			BufferedStream buffWrite = new BufferedStream(new FileStream(@"d:\APTest\SPCFDS_Monitor.exe",FileMode.Create,FileAccess.Write));
        //			byte[] byteBuff = new byte[256];
        //			int count = 0;
        //			do
        //			{
        //				count = buffRead.Read(byteBuff,0,256);
        //				buffWrite.Write(byteBuff,0,count);
        //				buffWrite.Flush();
        //				buffRead.Flush();
        //			}
        //			while(count>0);
        //			httpRespon.Close();
        //			stream.Close();
        //			buffWrite.Close();
        //			buffRead.Close();
        //		}

        //		public void CopyFromNet(string filename)
        //		{
        //			Uri uri = new Uri(ConfigurationSettings.AppSettings["Updater_Path"]+filename);
        //			httpReq = (HttpWebRequest)WebRequest.Create(uri);			
        //			httpRespon = (HttpWebResponse)httpReq.GetResponse();
        //			Stream stream = httpRespon.GetResponseStream();
        //			BufferedStream buffRead = new BufferedStream(stream);
        //			BufferedStream buffWrite = new BufferedStream(new FileStream(filename,FileMode.Create,FileAccess.Write));
        //			byte[] byteBuff = new byte[256];
        //			int count = 0;
        //			do
        //			{
        //				count = buffRead.Read(byteBuff,0,256);
        //				buffWrite.Write(byteBuff,0,count);
        //				buffWrite.Flush();
        //				buffRead.Flush();
        //			}
        //			while(count>0);
        //			httpRespon.Close();
        //			stream.Close();
        //			buffWrite.Close();
        //			buffRead.Close();
        //		}
    }
}
