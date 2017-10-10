using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsApplication
{
    class ExcelResolver
    {
        string currentPathname = System.AppDomain.CurrentDomain.BaseDirectory;
        private Dictionary<string, SingleDataRecord> rawDataMap = new Dictionary<string, SingleDataRecord>();

        public ExcelResolver()
        {

        }

        // filename contains absolute path
        public ExcelResolver(string[] filenameArray)
        {
            int fileCount = filenameArray.Length;
            for (int i = 0; i < fileCount; i++)
            {
                DataTable tempTable = ReadExcelToTable(currentPathname + "/" + filenameArray[i]);
                Console.WriteLine(tempTable.Rows.Count);
                Console.WriteLine();
            }
        }
        //public DataSet ExcelToDS(string Path)
        //{
        //    string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + Path + ";" + "Extended Properties=Excel 8.0;";
        //    OleDbConnection conn = new OleDbConnection(strConn);
        //    conn.Open();
        //    string strExcel = "";
        //    OleDbDataAdapter myCommand = null;
        //    DataSet ds = null;
        //    strExcel = "select * from [sheet1$]";
        //    myCommand = new OleDbDataAdapter(strExcel, strConn);
        //    ds = new DataSet();
        //    myCommand.Fill(ds, "Sheet1");
        //    return ds;
        //}

        //根据excle的路径把第一个sheel中的内容放入datatable
        private DataTable ReadExcelToTable(string path)//excel存放的路径
        {
            try
            {
                //连接字符串
                string connstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"; // Office 07及以上版本 不能出现多余的空格 而且分号注意
                //string connstring = Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"; //Office 07以下版本 
                using (OleDbConnection conn = new OleDbConnection(connstring))
                {
                    conn.Open();
                    DataTable sheetsName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" }); //得到所有sheet的名字
                    string firstSheetName = sheetsName.Rows[0][2].ToString(); //得到第一个sheet的名字
                    string sql = string.Format("SELECT * FROM [{0}]", firstSheetName); //查询字符串

                    OleDbDataAdapter ada = new OleDbDataAdapter(sql, connstring);
                    DataSet set = new DataSet();
                    ada.Fill(set);
                    return set.Tables[0];
                }
            }
            catch (Exception)
            {
                return null;
            }

        }
    }
}
