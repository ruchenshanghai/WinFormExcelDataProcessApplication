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
        private string targetKeyName = "Area";
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
            for (long fileIndex = 1; fileIndex < filenameArray.Length; fileIndex++)
            {
                resultFileContainer = resultFileContainer.MergeFileContainer(fileContainers[fileIndex], rangeDelta);
                fileContainers[fileIndex] = null;
                GC.Collect();
            }

            // need to output the result in two style: detail and simple
            long resultCount = resultFileContainer.recordList.Count;
            Console.WriteLine(resultCount);
            SaveByObjectLibrary();

            MessageBox.Show("Excel file created!!");
        }



        private SingleFileContainer ReadByObjectLibrary(string path)
        {
            string tempPathname = path;
            SingleFileContainer tempFileContainer = new SingleFileContainer();
            string tempTypename = "";
            string[] headerArray;
            long primaryIndex = -3;
            long secondIndex = -3;
            long targetIndex = -3;

            string tempValue;
            long tempRowIndex;
            long tempColumnIndex;
            long tempRowCount = 0;
            long tempColumnCount = 0;
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
                } else if (tempValue == targetKeyName)
                {
                    targetIndex = tempColumnIndex - 2;
                }
            }
            if ((primaryIndex == -3) || (secondIndex == -3))
            {
                // not found time column
                return null;
            }



            // get data
            for (tempRowIndex = 2; tempRowIndex <= tempRowCount; tempRowIndex++)
            {
                SingleDataRecord tempRecord = new SingleDataRecord();
                if (targetIndex == -3)
                {
                    // merge all column

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
                            if (!tempRecord.AddProperty(headerArray[tempColumnIndex - 2], tempValue))
                            {
                                // add property failed
                                return null;
                            }
                        }
                    }
                }
                else
                {
                    // only add Column: primary, second, Aera 
                    tempRecord.ID = (range.Cells[tempRowIndex, primaryIndex + 2] as Microsoft.Office.Interop.Excel.Range).Value2.ToString();
                    tempRecord.time = double.Parse((range.Cells[tempRowIndex, secondIndex + 2] as Microsoft.Office.Interop.Excel.Range).Value2.ToString());
                    if (!tempRecord.AddProperty(targetKeyName, (range.Cells[tempRowIndex, targetIndex + 2] as Microsoft.Office.Interop.Excel.Range).Value2.ToString(), tempTypename))
                    {
                        // add property failed
                        return null;
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
            Microsoft.Office.Interop.Excel.Workbooks xlWorkBooks = xlApp.Workbooks;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlWorkBooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = xlWorkBook.Worksheets[1];
            object misValue = System.Reflection.Missing.Value;
            if (xlApp == null || xlWorkBooks == null || xlWorkBook == null || xlWorkSheet == null)
            {
                return;
            }
            xlWorkSheet.Name = "result";
            xlWorkSheet.Cells[1, 1] = "result";

            // construct header
            xlWorkSheet.Cells[1, 2] = primaryKeyName;
            xlWorkSheet.Cells[1, 3] = secondKeyName;
            long columnIndex = 4;
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
                xlWorkSheet.Cells[recordIndex + 2, 2] = tempRecord.ID;
                xlWorkSheet.Cells[recordIndex + 2, 3] = tempRecord.time;
                for (int keyIndex = 0; keyIndex < tempKeyArray.Count; keyIndex++)
                {
                    xlWorkSheet.Cells[recordIndex + 2, keyIndex + 4] = tempRecord.propertyMap[(string)(tempKeyArray[keyIndex])];
                }
            }

            xlWorkBook.SaveAs(currentPathname + "result.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution, misValue, misValue, misValue, misValue);
            xlWorkBooks.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBooks);
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
