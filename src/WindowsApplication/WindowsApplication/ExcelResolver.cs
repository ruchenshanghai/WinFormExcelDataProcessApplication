using System;
using System.Collections;
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
        SingleFileContainer[] fileContainers;
        double rangeDelta = 0;
        private string primaryKeyName = "Chromatogram";
        private string secondKeyName = "RT [min]";
        private bool inputResult = false;


        public ExcelResolver()
        {

        }

        // filename contains absolute path
        public ExcelResolver(string[] filenameArray, double rangeParam)
        {
            rangeDelta = rangeParam;
            // get data from files
            int fileCount = filenameArray.Length;
            fileContainers = new SingleFileContainer[fileCount];
            for (int fileIndex = 0; fileIndex < fileCount; fileIndex++)
            {
                DataTable tempTable = ReadExcelToTable(currentPathname + "/" + filenameArray[fileIndex]);
                SingleFileContainer tempContainer = new SingleFileContainer();

                // construct the SingleDataRecord: ID, time, other...
                int tempWidth = tempTable.Columns.Count - 1;
                int tempHeight = tempTable.Rows.Count - 1;
                int primaryIndex = 0;
                int secondIndex = 0;
                string tempTypeName;

                if (tempWidth > 0 && tempHeight > 0)
                {
                    tempTypeName = tempTable.Rows[0][0].ToString();
                    string[] headerArray = new string[tempWidth];
                    // get headerName
                    for (int headerIndex = 0; headerIndex < tempWidth; headerIndex++)
                    {
                        headerArray[headerIndex] = tempTable.Rows[0][headerIndex + 1].ToString();
                        //Console.WriteLine(headerArray[headerIndex]);
                        if (headerArray[headerIndex] == primaryKeyName)
                        {
                            primaryIndex = headerIndex;
                        }
                        else if (headerArray[headerIndex] == secondKeyName)
                        {
                            secondIndex = headerIndex;
                        }
                    }
                    if ((primaryIndex == 0) && (secondIndex == 0))
                    {
                        // not found time column
                        return;
                    }
                    // get data
                    for (int heightIndex = 0; heightIndex < tempHeight; heightIndex++)
                    {
                        SingleDataRecord tempRecord = new SingleDataRecord();
                        for (int widthIndex = 0; widthIndex < tempWidth; widthIndex++)
                        {
                            string tempValue = tempTable.Rows[heightIndex + 1][widthIndex + 1].ToString();
                            if (widthIndex == primaryIndex)
                            {
                                tempRecord.ID = tempValue;
                            } else if (widthIndex == secondIndex)
                            {
                                tempRecord.time = double.Parse(tempValue);
                            }
                            else
                            {
                                tempRecord.AddProperty(headerArray[widthIndex], tempValue, tempTypeName);
                            }
                        }
                        tempContainer.InsertRecord(tempRecord);
                    }

                }
                fileContainers[fileIndex] = tempContainer;
            }
            inputResult = true;

            // merge all fileContainer
            fileContainers[0].MergeFileContainer(fileContainers[1]);
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

        //根据excle的路径把第一个sheet: Sheet1中的内容放入datatable
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
