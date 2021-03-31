using Pi.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Pi.ExportTool.ReadWriteFile
{
    public partial class frmReadFileExcelTemplate : Form
    {
        List<ExcelSheetModel> listSheet;
        List<ExcelCalculatorModel> listCalculator;

        public frmReadFileExcelTemplate()
        {
            InitializeComponent();
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            txtPath.Text = CommonOfficeExcel.OpenFileExcel();
        }

        private void btnReadFileExcel_Click(object sender, EventArgs e)
        {
            try
            {
                Excel.Application xlap = new Excel.Application();
                Excel.Workbook xbook = xlap.Workbooks.Open(txtPath.Text);

                //listSheet = CommonOfficeExcel.LoadExcel(filePath);
                listSheet = CommonOfficeExcel.LoadExcelExt(xbook);

                dgvListSheet.DataSource = listSheet;

                xbook.Close(true, null, null);
                xlap.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void btnTinhGiaTriLonNhat_Click(object sender, EventArgs e)
        {
            try
            {
                listSheet = dgvListSheet.DataSource as List<ExcelSheetModel>;
                listCalculator = new List<ExcelCalculatorModel>();

                if (backgroundProcessing.IsBusy != true)
                {
                    backgroundProcessing.RunWorkerAsync(listSheet);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void backgroundProcessing_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                BackgroundWorker worker = sender as BackgroundWorker;
                e.Result = ComputeData((List<ExcelSheetModel>)e.Argument, worker, e);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void backgroundProcessing_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.progressBar1.Value = e.ProgressPercentage;
        }

        List<ExcelCalculatorModel> ComputeData(List<ExcelSheetModel> list, BackgroundWorker worker, DoWorkEventArgs e)
        {
            Excel.Application xlap = new Excel.Application();
            Excel.Workbook xbook = xlap.Workbooks.Open(txtPath.Text);            

            int i = 1;

            foreach (var dto in list)
            {
                if (worker.CancellationPending == true)
                {
                    e.Cancel = true;
                    break;
                }
                else
                {
                    Excel.Worksheet xsheet = xbook.Sheets[dto.SheetName];

                    //var dtoCalculator = CommonOfficeExcel.LoadExcelSheet(txtPath.Text, dto.SheetName, dto.ColumnIndex);
                    var dtoCalculator = CommonOfficeExcel.LoadExcelSheetExt(xlap, xsheet, dto.ColumnIndex);

                    listCalculator.Add(dtoCalculator);

                    int percentComplete = (int)((float)i / (float)list.Count * 100);

                    worker.ReportProgress(percentComplete);
                    i++;
                }
            }

            xbook.Close(true, null, null);
            xlap.Quit();

            return listCalculator;
        }

        private void backgroundProcessing_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            dgvListMaxValue.DataSource = listCalculator;
        }
    }
}
