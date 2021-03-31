using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Pi.Core
{
    /// <summary>
    /// File này chứa các hàm liên quan tới file excel
    /// </summary>
    public static class CommonOfficeExcel
    {
        public static String openFile()
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
            //cbbSheetName.Items.Clear();
            //txtSheet.Clear();
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //txtfilename.Text = file.FileName;

                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook excelBook = xlApp.Workbooks.Open(file.FileName);
                foreach (Microsoft.Office.Interop.Excel.Worksheet wSheet in excelBook.Worksheets)
                {
                    //cbbSheetName.Items.Add(wSheet.Name);
                }
            }
            return file.FileName;
        }
        public static DataTable ReadExcel(string fileName, string fileExt, String sheetName)
        {

            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    //OleDbDataAdapter oleAdpt = new OleDbDataAdapter($"select * from [Sheet{sheetnumber}$]", con); //here we read data from sheet1
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter($"select * from [{sheetName}$]", con); //here we read data from sheet1
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch { }
            }
            return dtexcel;
        }

        public static List<ExcelSheetModel> LoadExcel(string pathFile)
        {
            try
            {
                int defaultColumn = 21;

                var list = new List<ExcelSheetModel>();

                Excel.Application xlap = new Excel.Application();
                Excel.Workbook xbook = xlap.Workbooks.Open(pathFile);

                for (int i = 0; i < xbook.Worksheets.Count; i++)
                {
                    Excel.Worksheet xsheet = xbook.Sheets[i + 1];

                    //int columnCount = xsheet.UsedRange.Columns.Count;
                    string columnName = xsheet.Cells[2, defaultColumn].Value;

                    var dto = new ExcelSheetModel()
                    {
                        SheetName = xsheet.Name,
                        ColumnIndex = defaultColumn,
                        ColumnName = columnName,
                    };

                    list.Add(dto);
                }

                xbook.Close(true, null, null);
                xlap.Quit();

                Marshal.ReleaseComObject(xbook);
                Marshal.ReleaseComObject(xlap);

                return list;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static ExcelCalculatorModel LoadExcelSheet(string pathFile, string sheetName, int columnIndex)
        {
            try
            {
                Excel.Application xlap = new Excel.Application();
                Excel.Workbook xbook = xlap.Workbooks.Open(pathFile);
                Excel.Worksheet xsheet = xbook.Sheets[sheetName];

                //int columnCount = xsheet.UsedRange.Columns.Count;
                string columnName = xsheet.Cells[2, columnIndex].Value;

                var dblMin = xlap.WorksheetFunction.Min(xsheet.Range[xsheet.Cells[3, columnIndex], xsheet.Cells[xsheet.UsedRange.Rows.Count, columnIndex]]);
                var dblMax = xlap.WorksheetFunction.Max(xsheet.Range[xsheet.Cells[3, columnIndex], xsheet.Cells[xsheet.UsedRange.Rows.Count, columnIndex]]);

                var dblMaxFinal = xlap.WorksheetFunction.ImAbs(dblMin) > xlap.WorksheetFunction.ImAbs(dblMax)
                                  ? xlap.WorksheetFunction.Round(xlap.WorksheetFunction.ImAbs(dblMin), 0)
                                  : xlap.WorksheetFunction.Round(xlap.WorksheetFunction.ImAbs(dblMax), 0);

                var dto = new ExcelCalculatorModel()
                {
                    SheetName = xsheet.Name,
                    ColumnIndex = columnIndex,
                    ColumnName = columnName,
                    MinValue = dblMin,
                    MaxValue = dblMax,
                    MaxValueFinal = dblMaxFinal
                };

                xbook.Close(true, null, null);
                xlap.Quit();

                Marshal.ReleaseComObject(xsheet);
                Marshal.ReleaseComObject(xbook);
                Marshal.ReleaseComObject(xlap);

                return dto;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static string OpenFileExcel()
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog()
            {
                ValidateNames = true,
                Multiselect = false,
                Filter = "Excel Doucment|*.xlsx|Excel 97 - 2003 Document|*.xls"
            };

            if (file.ShowDialog() == DialogResult.OK)
                return file.FileName;
            else
                return "";

        }

        public static List<ExcelSheetModel> LoadExcelExt(Excel.Workbook xbook)
        {
            try
            {
                int defaultColumn = 21;

                var list = new List<ExcelSheetModel>();

                //Excel.Application xlap = new Excel.Application();
                //Excel.Workbook xbook = xlap.Workbooks.Open(pathFile);

                for (int i = 0; i < xbook.Worksheets.Count; i++)
                {
                    Excel.Worksheet xsheet = xbook.Sheets[i + 1];

                    string columnName = xsheet.Cells[2, defaultColumn].Value;

                    var dto = new ExcelSheetModel()
                    {
                        SheetName = xsheet.Name,
                        ColumnIndex = defaultColumn,
                        ColumnName = columnName,
                    };

                    list.Add(dto);
                }

                return list;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static ExcelCalculatorModel LoadExcelSheetExt(Excel.Application xlap, Excel.Worksheet xsheet, int columnIndex)
        {
            try
            {
                //Excel.Application xlap = new Excel.Application();
                //Excel.Workbook xbook = xlap.Workbooks.Open(pathFile);
                //Excel.Worksheet xsheet = xbook.Sheets[sheetName];

                //int columnCount = xsheet.UsedRange.Columns.Count;
                string columnName = xsheet.Cells[2, columnIndex].Value;

                var dblMin = xlap.WorksheetFunction.Min(xsheet.Range[xsheet.Cells[3, columnIndex], xsheet.Cells[xsheet.UsedRange.Rows.Count, columnIndex]]);
                var dblMax = xlap.WorksheetFunction.Max(xsheet.Range[xsheet.Cells[3, columnIndex], xsheet.Cells[xsheet.UsedRange.Rows.Count, columnIndex]]);

                //var dblMaxFinal1 = xlap.WorksheetFunction.ImAbs(dblMin) > xlap.WorksheetFunction.ImAbs(dblMax)
                //                  ? xlap.WorksheetFunction.Round(xlap.WorksheetFunction.ImAbs(dblMin), 0)
                //                  : xlap.WorksheetFunction.Round(xlap.WorksheetFunction.ImAbs(dblMax), 0);
                var dblMaxFinal = Math.Abs(dblMin) > Math.Abs(dblMax) ? Math.Abs(dblMin) : Math.Abs(dblMax);

                var dto = new ExcelCalculatorModel()
                {
                    SheetName = xsheet.Name,
                    ColumnIndex = columnIndex,
                    ColumnName = columnName,
                    MinValue = dblMin,
                    MaxValue = dblMax,
                    MaxValueFinal = dblMaxFinal
                };

                //xbook.Close(true, null, null);
                //xlap.Quit();

                //Marshal.ReleaseComObject(xsheet);
                //Marshal.ReleaseComObject(xbook);
                //Marshal.ReleaseComObject(xlap);

                return dto;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
