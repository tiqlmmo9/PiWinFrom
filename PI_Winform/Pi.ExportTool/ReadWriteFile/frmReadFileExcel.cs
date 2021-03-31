using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pi.ExportTool.Net.ReadWriteFile
{
    public partial class frmReadFileExcel : Form
    {
        public frmReadFileExcel()
        {
            InitializeComponent();
        }

        private DataTable ReadExcel(string fileName, string fileExt, string sheet)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //cho Ex < 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //cho Ex > 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter($"select * from [{sheet}$]", con); //đọc sheet  
                    oleAdpt.Fill(dtexcel); //ghi dữ liệu Ex vào DataTable  
                }
                catch { }
            }
            return dtexcel;
        }

        

        private void btnOpen_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog(); //mở hộp thoại chọn file
            if (file.ShowDialog() == DialogResult.OK) //Nếu chọn file  
            {
                txtPath.Text = file.FileName; //lấy, ghi đường dẫn vào txtPath
            }
        }

        private void btnRead_Click_1(object sender, EventArgs e)
        {
            string sheet = txtSheet.Text;
            string filePath = txtPath.Text; //Lấy đường dẫn của file 
            string fileExt = Path.GetExtension(filePath); //lấy phần mở rộng file  
            if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
            {
                try
                {
                    DataTable dtExcel = new DataTable();
                    dtExcel = ReadExcel(filePath, fileExt, sheet); //đọc file  
                    dataGridView1.Visible = true;
                    dataGridView1.DataSource = dtExcel;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
            else
            {
                MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtPath_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
