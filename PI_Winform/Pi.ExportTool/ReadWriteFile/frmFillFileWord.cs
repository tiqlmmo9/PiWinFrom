using Pi.Core;
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
    public partial class frmFillFileWord : Form
    {
        public frmFillFileWord()
        {
            InitializeComponent();
        }

        private void frmFillFileWord_Load(object sender, EventArgs e)
        {
            //for test
            //CommonOfficeWord.WriteFunctionHere();
        }
    }
}
