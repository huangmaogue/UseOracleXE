using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DAO = Microsoft.Office.Interop.Access.Dao;

namespace OracleTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MsAccessDAO msa = new MsAccessDAO(@"z:\blank.accdb");
            msa.NonQuery("update Article set JSON=''");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string path = "z:\\blank.accdb";
            DAO.DBEngine dbe = new DAO.DBEngine();
            DAO.Database db = dbe.OpenDatabase(path);
            DAO.TableDef tbdef = db.CreateTableDef();
            tbdef.Name = "linkedtable2";//Whatever you want the linked table to be named
            tbdef.Connect = "ODBC;Driver={Oracle in XE};dsn={Oracle - NXESL};dbq=XE;Uid=nxesl;Pwd=admin;Database=NXESL;Trusted_Connection=YES";
            tbdef.SourceTableName = "NXESL.Article";//Whatever the SQL Server Table Name is
            db.TableDefs.Append(tbdef);
        }
    }
}
