using Pi.Core;
using Pi.ExportTool.Net.ReadWriteFile;
using Pi.ExportTool.ReadWriteFile;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pi.Export.Net
{
    public partial class frmMenu : Form
    {
        public frmMenu()
        {
            InitializeComponent();
        }

        private void btnDemoFillFileWord_Click(object sender, EventArgs e)
        {
            frmFillFileWord frm = new frmFillFileWord();

            // Show testDialog as a modal dialog and determine if DialogResult = OK.
            if (frm.ShowDialog(this) == DialogResult.OK)
            {
            }
            frm.Dispose();
        }

        private void fileFileWordDùngMicrosoftOfficeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //frmFillFileWordOfficeMicrosoft frm = new frmFillFileWordOfficeMicrosoft();

            //// Show testDialog as a modal dialog and determine if DialogResult = OK.
            //if (frm.ShowDialog(this) == DialogResult.OK)
            //{
            //}
            //frm.Dispose();
        }

        private void thienToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void btnDocObjectWord_Click(object sender, EventArgs e)
        {
            var w = new frmFillFileWord();
            w.ShowDialog();
        }

        private void btnTest1_Click(object sender, EventArgs e)
        {
            var w = new frmTest1();
            w.ShowDialog();
        }

        private void btnExecl_Click(object sender, EventArgs e)
        {
            var f = new test2();
            f.ShowDialog();
        }

        private void testExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var ex = new frmReadFileExcel();
            ex.ShowDialog();
        }

        private void btnLoadDataExcel_Click(object sender, EventArgs e)
        {
            var frm = new frmReadFileExcelTemplate();
            frm.ShowDialog();
        }

        private void piToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var pi = new frmPi();
            pi.ShowDialog();
        }
    }
}
