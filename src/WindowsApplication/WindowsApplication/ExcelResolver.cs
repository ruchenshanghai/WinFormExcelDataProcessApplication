using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsApplication
{
    class ExcelResolver
    {
        string currentPathname = System.AppDomain.CurrentDomain.BaseDirectory;
        private SingleFileContainer[] fileContainers;
        private SingleFileContainer resultFileContainer;
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
            inputResult = ReadByObjectLibrary(filenameArray);
            if (!inputResult)
            {
                // input error
                MessageBox.Show("source Excel file has format error!!");
                return;
            }

            // merge all fileContainer
            resultFileContainer = fileContainers[0];
            for (int fileIndex = 1; fileIndex < filenameArray.Length; fileIndex++)
            {
                resultFileContainer = resultFileContainer.MergeFileContainer(fileContainers[fileIndex], rangeDelta);
            }

            // need to output the result in two style: detail and simple
            int resultCount = resultFileContainer.recordList.Count;
            Console.WriteLine(resultCount);
            SaveByObjectLibrary();

        }



        private SingleFileContainer ReadByObjectLibrary(string path)
        {
            string tempPathname = path;
            SingleFileContainer tempFileContainer = new SingleFileContainer();
            string tempTypename = "";
            string[] headerArray;
            int primaryIndex = -1;
            int secondIndex = -1;

            string tempValue;
            int tempRowIndex;
            int tempColumnIndex;
            int tempRowCount = 0;
            int tempColumnCount = 0;
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Range range;
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(tempPathname, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            range = xlWorkSheet.UsedRange;
            tempRowCount = range.Rows.Count;
            tempColumnCount = range.Columns.Count;
            if (tempRowCount == 0 || tempColumnCount == 0)
            {
                // no data
                return null;
            }

            // get header row first
            tempTypename = (range.Cells[1, 1] as Microsoft.Office.Interop.Excel.Range).Value2.ToString();
            headerArray = new string[tempColumnCount - 1];
            for (tempColumnIndex = 2; tempColumnIndex <= tempColumnCount; tempColumnIndex++)
            {
                if ((range.Cells[1, tempColumnIndex] as Microsoft.Office.Interop.Excel.Range).Value2 == null)
                {
                    tempValue = SingleFileContainer.DEFAULT_VALUE;
                }
                else
                {
                    tempValue = (range.Cells[1, tempColumnIndex] as Microsoft.Office.Interop.Excel.Range).Value2.ToString();
                }
                //tempValue = (range.Cells[1, tempColumnIndex] as Microsoft.Office.Interop.Excel.Range).Value2.ToString();
                headerArray[tempColumnIndex - 2] = tempValue;
                //Console.WriteLine(headerArray[headerIndex]);
                if (tempValue == primaryKeyName)
                {
                    primaryIndex = tempColumnIndex - 2;
                }
                else if (tempValue == secondKeyName)
                {
                    secondIndex = tempColumnIndex - 2;
                }
            }
            if ((primaryIndex == -1) || (secondIndex == -1))
            {
                // not found time column
                return null;
            }

            // get data
            for (tempRowIndex = 2; tempRowIndex <= tempRowCount; tempRowIndex++)
            {
                SingleDataRecord tempRecord = new SingleDataRecord();
                for (tempColumnIndex = 2; tempColumnIndex <= tempColumnCount; tempColumnIndex++)
                {
                    if ((range.Cells[tempRowIndex, tempColumnIndex] as Microsoft.Office.Interop.Excel.Range).Value2 == null)
                    {
                        tempValue = SingleFileContainer.DEFAULT_VALUE;
                    }
                    else
                    {
                        tempValue = (range.Cells[tempRowIndex, tempColumnIndex] as Microsoft.Office.Interop.Excel.Range).Value2.ToString();
                    }
                    if ((tempColumnIndex - 2) == primaryIndex)
                    {
                        tempRecord.ID = tempValue;
                    }
                    else if ((tempColumnIndex - 2) == secondIndex)
                    {
                        tempRecord.time = double.Parse(tempValue);
                    }
                    else
                    {
                        if (!tempRecord.AddProperty(headerArray[tempColumnIndex - 2], tempValue, tempTypename))
                        {
                            // add property failed
                            return null;
                        }
                    }
                }
                tempFileContainer.InsertRecord(tempRecord);
            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            return tempFileContainer;
        }

        private bool ReadByObjectLibrary(string[] filenameArray)
        {
            int fileCount = filenameArray.Length;
            fileContainers = new SingleFileContainer[fileCount];
            for (int fileIndex = 0; fileIndex < fileCount; fileIndex++)
            {
                string tempPathname = currentPathname + "/" + filenameArray[fileIndex];
                fileContainers[fileIndex] = ReadByObjectLibrary(tempPathname);
                if (fileContainers[fileIndex] == null)
                {
                    // read error
                    return false;
                }
            }

            //Console.WriteLine(fileContainers.Length);
            return true;
        }

        private void SaveByObjectLibrary()
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }
            else
            {
                MessageBox.Show("Welcome!!");

                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                // construct header
                xlWorkSheet.Cells[1, 1] = primaryKeyName;
                xlWorkSheet.Cells[1, 2] = secondKeyName;
                int columnIndex = 3;
                ArrayList tempKeyArray = resultFileContainer.keyArray;
                for (int headerIndex = 0; headerIndex < tempKeyArray.Count; headerIndex++)
                {
                    xlWorkSheet.Cells[1, columnIndex] = tempKeyArray[headerIndex];
                    columnIndex++;
                }

                // construct the content
                for (int recordIndex = 0; recordIndex < resultFileContainer.recordList.Count; recordIndex++)
                {
                    SingleDataRecord tempRecord = (SingleDataRecord)(resultFileContainer.recordList[recordIndex]);
                    xlWorkSheet.Cells[recordIndex + 2, 1] = tempRecord.ID;
                    xlWorkSheet.Cells[recordIndex + 2, 2] = tempRecord.time;
                    for (int keyIndex = 0; keyIndex < tempKeyArray.Count; keyIndex++)
                    {
                        xlWorkSheet.Cells[recordIndex + 2, keyIndex + 3] = tempRecord.propertyMap[(string)(tempKeyArray[keyIndex])];
                    }
                }


                //xlWorkSheet.Cells[1, 1] = "ID";
                //xlWorkSheet.Cells[1, 2] = "Name";
                //xlWorkSheet.Cells[2, 1] = "1";
                //xlWorkSheet.Cells[2, 2] = "One";
                //xlWorkSheet.Cells[3, 1] = "2";
                //xlWorkSheet.Cells[3, 2] = "Two";



                xlWorkBook.SaveAs("E:\\Project\\School\\test-Excel.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);

                MessageBox.Show("Excel file created!!");

            }
        }

        //public void DSToExcel(string Path, DataSet oldds)
        //{
        //    //先得到汇总EXCEL的DataSet 主要目的是获得EXCEL在DataSet中的结构 
        //    string strCon = " Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source =" + Path + ";Extended Properties=Excel 8.0";
        //    OleDbConnection myConn = new OleDbConnection(strCon);
        //    string strCom = "select * from [Sheet1$]";
        //    myConn.Open();
        //    OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, myConn);
        //    ystem.Data.OleDb.OleDbCommandBuilder builder = new OleDbCommandBuilder(myCommand);
        //    //QuotePrefix和QuoteSuffix主要是对builder生成InsertComment命令时使用。 
        //    builder.QuotePrefix = "[";     //获取insert语句中保留字符（起始位置） 
        //    builder.QuoteSuffix = "]"; //获取insert语句中保留字符（结束位置） 
        //    DataSet newds = new DataSet();
        //    myCommand.Fill(newds, "Table1");
        //    for (int i = 0; i < oldds.Tables[0].Rows.Count; i++)
        //    {
        //        //在这里不能使用ImportRow方法将一行导入到news中，因为ImportRow将保留原来DataRow的所有设置(DataRowState状态不变)。
        //        在使用ImportRow后newds内有值，但不能更新到Excel中因为所有导入行的DataRowState != Added
        //    DataRow nrow = aDataSet.Tables["Table1"].NewRow();
        //        for (int j = 0; j < newds.Tables[0].Columns.Count; j++)
        //        {
        //            nrow[j] = oldds.Tables[0].Rows[i][j];
        //        }
        //        newds.Tables["Table1"].Rows.Add(nrow);
        //    }
        //    myCommand.Update(newds, "Table1");
        //    myConn.Close();
        //}

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

        ////根据excle的路径把第一个sheet: Sheet1中的内容放入datatable
        //private DataTable ReadExcelToTable(string path)//excel存放的路径
        //{
        //    try
        //    {
        //        //连接字符串
        //        string connstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"; // Office 07及以上版本 不能出现多余的空格 而且分号注意
        //        //string connstring = Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"; //Office 07以下版本 
        //        using (OleDbConnection conn = new OleDbConnection(connstring))
        //        {
        //            conn.Open();
        //            DataTable sheetsName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" }); //得到所有sheet的名字
        //            string firstSheetName = sheetsName.Rows[0][2].ToString(); //得到第一个sheet的名字
        //            string sql = string.Format("SELECT * FROM [{0}]", firstSheetName); //查询字符串

        //            OleDbDataAdapter ada = new OleDbDataAdapter(sql, connstring);
        //            DataSet set = new DataSet();
        //            ada.Fill(set);
        //            return set.Tables[0];
        //        }
        //    }
        //    catch (Exception)
        //    {
        //        return null;
        //    }

        //}
    }
}
