using Pi.Core;
using Pi.Core.model;
using Pi.ExportTool.ReadWriteFile;
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
using Word = Microsoft.Office.Interop.Word;

namespace Pi.ExportTool.Net.ReadWriteFile
{
    public partial class frmTest1 : Form
    {
        public frmTest1()
        {
            InitializeComponent();
        }
        UserModel user;
        public frmTest1(UserModel user)
        {
            InitializeComponent();
            //this.user = user;
            this.txtParam1.Text = user.name;
            this.txtParam2.Text = user.BirthDay;
            this.txtParam3.Text = user.Address;
            
        }

        private void frmTest1_Load(object sender, EventArgs e)
        {
            //for test
            //CommonOfficeWord.WriteFunctionHere();
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog()
            {
                ValidateNames = true,
                Multiselect = false,
                Filter = "Word Document|*.docx|Word 97 - 2003 Document|*.doc"
            })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    txtPath.Text = ofd.FileName;
                    object readOnly = true;
                    object visible = true;
                    object save = false;
                    object fileName = ofd.FileName;
                    object missing = Type.Missing;
                    object newTemplate = false;
                    object docType = 0;
                    Word._Document oDoc = null;
                    Word._Application oWord = new Word.Application()
                    {
                        Visible = false
                    };
                    oDoc = oWord.Documents.Open(
                            ref fileName, ref missing, ref readOnly, ref missing,
                            ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref visible,
                            ref missing, ref missing, ref missing, ref missing);
                    oDoc.ActiveWindow.Selection.WholeStory();
                    oDoc.ActiveWindow.Selection.Copy();
                    IDataObject data = Clipboard.GetDataObject();
                    rtbWord.Rtf = data.GetData(DataFormats.Rtf).ToString();
                    oWord.Quit(ref missing, ref missing, ref missing);
                }
            }
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                CommonOfficeWord.CreateWordDocument(txtPath.Text, sfd.FileName, txtParam1.Text, txtParam2.Text, txtParam3.Text);
            }
        }
    }
}
