using ExcelDataReader;
using Microsoft.Office.Interop.Word;
using Pi.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Word.Application;
using DataTable = System.Data.DataTable;
using Word = Microsoft.Office.Interop.Word;

namespace Pi.ExportTool.ReadWriteFile
{
    public partial class frmPi : Form
    {
        private const int DEFAULT_VALUE = 0;
        public static string paramResult = null;
        //public static string param2 = null;
        //public static string param3 = null;
        public static string paramSelected = null;
        DataSet ds;
        public frmPi()
        {
            InitializeComponent();
            cbTemplate.SelectedIndex = DEFAULT_VALUE;
            cbFunc.SelectedIndex = DEFAULT_VALUE;
            //cbParam.SelectedIndex = DEFAULT_VALUE;
            cbParam.Text = "--Select Param Index --";
        }

        //Select a template
        private void cbTemplate_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                ComboBox cb = sender as ComboBox;
                var template = cb.SelectedItem.ToString();
                object FileName = @"C:\Users\ich\Desktop\PI_Winform\Pi.ExportTool\TemplateFile\" + template + ".docx";
                object readOnly = true;
                object visible = true;
                object save = false;
                object missing = Type.Missing;
                object newTemplate = false;
                object docType = 0;
                Word._Document oDoc = null;
                Word._Application oWord = new Word.Application()
                {
                    Visible = false
                };
                oDoc = oWord.Documents.Open(
                        ref FileName, ref missing, ref readOnly, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref visible,
                        ref missing, ref missing, ref missing, ref missing);
                oDoc.ActiveWindow.Selection.WholeStory();
                oDoc.ActiveWindow.Selection.Copy();
                IDataObject data = Clipboard.GetDataObject();
                richTextBox1.Rtf = data.GetData(DataFormats.Rtf).ToString();
                oWord.Quit(ref missing, ref missing, ref missing);

                GetParams();

                Cursor.Current = Cursors.Default;
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void GetParams()
        {
            var paramCount = FindParamCount(richTextBox1, "PI_PARAM");
            var listParam = new List<string>();

            for (int i = 1; i <= paramCount; i++)
            {
                var item = "[PI_PARAM" + i + "]";
                listParam.Add(item);
            }

            cbParam.DataSource = listParam;
        }

        //Reset Template
        private void btnResetTemplate_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
            cbTemplate.Text = "--Select a template--";
        }


        //Browse a file
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel Workbook 97-2003|*.xls", ValidateNames = true })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    Cursor.Current = Cursors.WaitCursor;

                    using (var stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        IExcelDataReader reader;
                        if (ofd.FilterIndex == 2)
                        {
                            reader = ExcelReaderFactory.CreateBinaryReader(stream);
                        }
                        else
                        {
                            reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                        }

                        ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        });

                        cbSheet.Items.Clear();
                        foreach (DataTable dt in ds.Tables)
                        {
                            cbSheet.Items.Add(dt.TableName);
                        }
                        reader.Close();

                    }
                }
                txtPath.Text = ofd.FileName;
                cbSheet.SelectedIndex = DEFAULT_VALUE;
            }

            
            Cursor.Current = Cursors.Default;

        }


        //Select a sheet
        private void cbSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.DataSource = ds.Tables[cbSheet.SelectedIndex];
        }

        //Select a function
        private void cbFunc_SelectedIndexChanged(object sender, EventArgs e)
        {
            lbFunc.Text = cbFunc.Text;
        }


        //Select parameter index
        //private void cbParamIndex_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    var listParam = new List<string>();
        //    var paramIndex = Convert.ToInt32(cbParamIndex.SelectedItem);

        //    for (int i = 1; i <= paramIndex; i++)
        //    {
        //        var item = "[PI_PARAM" + i + "]";
        //        listParam.Add(item);
        //    }

        //    cbParam.DataSource = listParam;
        //}



        //Double click to execute function
        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                var currentRow = dataGridView1.CurrentCell.ColumnIndex;
                if (cbFunc.SelectedItem.ToString() == "MAX")
                {
                    paramResult = dataGridView1.Rows.Cast<DataGridViewRow>()
                             .Max(r => Math.Abs(Convert.ToDouble(r.Cells[currentRow].Value))).ToString();
                }

                if (cbFunc.SelectedItem.ToString() == "SUM")
                {
                    paramResult = dataGridView1.Rows.Cast<DataGridViewRow>()
                             .Sum(r => Math.Abs(Convert.ToDouble(r.Cells[currentRow].Value))).ToString();
                }

                if (cbFunc.SelectedItem.ToString() == "AVG")
                {
                    paramResult = dataGridView1.Rows.Cast<DataGridViewRow>()
                             .Average(r => Math.Abs(Convert.ToDouble(r.Cells[currentRow].Value))).ToString();
                }



                lbFuncValue.Text = paramResult;
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            //MessageBox.Show(MaxID.ToString());


        }

        //Replace parameter
        public static void QuickReplace(RichTextBox rtb, String oldValue, String newValue)
        {
            rtb.Rtf = rtb.Rtf.Replace(oldValue, newValue);
        }


        //Execute
        private void btnExcute_Click(object sender, EventArgs e)
        {

            paramSelected = cbParam.SelectedItem.ToString();
            if (cbParam.SelectedItem == null)
            {
                MessageBox.Show("Please select a param index.");
                return;
            }
            if (paramResult == null)
            {
                MessageBox.Show("param1 is null. Please double click column.");
                return;
            }
            if (cbTemplate.SelectedItem == null)
            {
                MessageBox.Show("Please select a template.");
                return;
            }
            QuickReplace(richTextBox1, paramSelected, paramResult);

        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtColumnIndex.Text = (dataGridView1.CurrentCell.ColumnIndex + 1).ToString();

            //Find2(richTextBox1, "[PI_PARAM");


            //var x1 = richTextBox1.Find("[PI_PARAM1]");
            //var x2 = richTextBox1.Find("[PI_PARAM2]");
            //var x3 = richTextBox1.Find("[PI_PARAM3]");
            //var x4 = richTextBox1.Find("[PI_PARAM4]");
            //var x5 = richTextBox1.Find("[PI_PARAM5]");
            //var x6 = richTextBox1.Find("[PI_PARAM6]");

        }
        //private void dataGridView1_DragDrop(object sender, DragEventArgs e)
        //{
        //    int v = richTextBox1.GetCharIndexFromPosition(Point.);
        //}












        //private void btnReplace_Click(object sender, EventArgs e)
        //{

        //    Find(richTextBox1, textBox2.Text, Color.Red);
        //}

        //public static void Find(RichTextBox rtb, String word, Color color)
        //{
        //    if (word == "")
        //    {
        //        return;
        //    }
        //    int s_start = rtb.SelectionStart, startIndex = 0, index;
        //    while ((index = rtb.Text.IndexOf(word, startIndex)) != -1)
        //    {
        //        rtb.Select(index, word.Length);
        //        rtb.SelectionColor = color;
        //        startIndex = index + word.Length;
        //    }
        //    rtb.SelectionStart = 0;
        //    rtb.SelectionLength = rtb.TextLength;
        //    //rtb.SelectionColor = Color.Black;
        //}

        public static int FindParamCount(RichTextBox rtb, string word)
        {

            int startIndex = 0, count = 0;
            while (startIndex < rtb.TextLength)
            {
                int index = rtb.Find(word, startIndex, RichTextBoxFinds.None);
                if (index != -1)
                {
                    count++;
                }
                else break;
                startIndex = index + word.Length;
            }

            return count;
            //rtb.SelectionColor = Color.Black;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            SaveMyFile();
        }

        public void SaveMyFile()
        {
            // Create temp rtf file
            var fileName = @"C:\Users\ich\Desktop\PI_Winform\Pi.ExportTool\TemplateFile\temp.rtf";

            richTextBox1.SaveFile(fileName, RichTextBoxStreamType.RichText);

            //Create docx file
            SaveFileDialog saveFile1 = new SaveFileDialog();
            saveFile1.DefaultExt = "*.docx";
            saveFile1.Filter = "Word Files|*.docx";

            if (saveFile1.ShowDialog() == System.Windows.Forms.DialogResult.OK &&
               saveFile1.FileName.Length > 0)
            {
                var wordApp = new Microsoft.Office.Interop.Word.Application();
                var currentDoc = wordApp.Documents.Open(@"C:\Users\ich\Desktop\PI_Winform\Pi.ExportTool\TemplateFile\temp.rtf");
                currentDoc.SaveAs(saveFile1.FileName, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault);
                currentDoc.Close();
                wordApp.Quit();
                MessageBox.Show("Created file successful!");
            }
        }

    }
}
