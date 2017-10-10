using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsApplication
{
    public partial class MainForm : Form
    {
        string[] sourceFilenameArray;
        double rangeDealta = 0;
        string outputResultFilename = "result.xlsx";
        string outputDetailFilename = "detail.xlsx";

        public MainForm()
        {
            InitializeComponent();

            //ExcelResolver testResolver = new ExcelResolver();
            //DataTable resultTable = testResolver.ReadExcelToTable(@"C:\Users\wenja\Desktop\Project\School\Data\QC.xlsx");
            //Console.WriteLine(resultTable.Rows.Count);
            //Console.WriteLine(resultTable.Select().ToString());
        }


        private void MainSourceButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx|所有文件|*.*";
            openFileDialog.ValidateNames = true;
            openFileDialog.CheckPathExists = true;
            openFileDialog.CheckFileExists = true;
            openFileDialog.Multiselect = true;



            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // get filename array
                Console.WriteLine("Range Dealt: " + rangeDealta);
                sourceFilenameArray = openFileDialog.SafeFileNames;
                //for (int i = 0; i < sourceFilenameArray.Length; i++)
                //{
                //    Console.WriteLine(sourceFilenameArray[i]);

                //}
                ExcelResolver excelResolver = new ExcelResolver(sourceFilenameArray);
            }

        }

        public static bool IsNumeric(string value)
        {
            return Regex.IsMatch(value, @"^[+-]?\d*[.]?\d*$");
        }

        private void RangeInput_TextChanged(object sender, EventArgs e)
        {
            if (!IsNumeric(RangeInput.Text))
            {
                RangeInput.Text = "";
                return;
            }
            else
            {
                rangeDealta = Double.Parse(RangeInput.Text);
            }
        }


    }

}
