using Pi.Core;
using Pi.Core.model;
using Pi.ExportTool.Net.ReadWriteFile;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pi.ExportTool.ReadWriteFile
{
    public partial class test2 : Form
    {
        public test2()
        {
            InitializeComponent();
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            txtPathFile.Text = CommonOfficeExcel.openFile();
        }

        private void btnReadFile_Click(object sender, EventArgs e)
        {
            String sheetName = txtSheetName.Text;
            String filePath = txtPathFile.Text; //get the path of the file  
            String fileExt = Path.GetExtension(filePath); //get the file extension 
            if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
            {
                try
                {
                    DataTable dtExcel = new DataTable();
                    dtExcel = CommonOfficeExcel.ReadExcel(filePath, fileExt, sheetName); //read excel file  
                    dataGrid.Visible = true;
                    dataGrid.DataSource = dtExcel;
                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
            else
            {
                MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
            }
        }
        private void dataGrid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Tab && dataGrid.CurrentCell.ColumnIndex == 1)
            {
                e.Handled = true;
                DataGridViewCell cell = dataGrid.Rows[0].Cells[0];
                dataGrid.CurrentCell = cell;
                dataGrid.BeginEdit(true);
            }
        }

        private void btnSend_Click(object sender, EventArgs e)
        {
           if(dataGrid.DataSource != null)
            {
                UserModel user = new UserModel()
                {
                    name = dataGrid.Rows[0].Cells[0].Value.ToString() ?? String.Empty,
                    BirthDay = dataGrid.Rows[1].Cells[0].Value.ToString() ?? String.Empty,
                    Address = dataGrid.Rows[2].Cells[0].Value.ToString() ?? String.Empty
                    //name="thien",
                    //BirthDay="12/12/1212",
                    //Address="hue"
                };
                var f = new frmTest1(user);
                f.ShowDialog();
                //this.Close();
            }
            else
            {
                MessageBox.Show("chua co data");
            }

        }
    }
}
