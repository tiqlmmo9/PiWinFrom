namespace Pi.ConvertTool.Net.ReadWriteFile
{
    partial class frmFillFileWordOfficeMicrosoft
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmFillFileWordOfficeMicrosoft));
            this.printPreviewDialog1 = new System.Windows.Forms.PrintPreviewDialog();
            this.documentViewer1 = new DevExpress.XtraPrinting.Preview.DocumentViewer();
            this.documentViewer2 = new DevExpress.XtraPrinting.Preview.DocumentViewer();
            this.documentViewer3 = new DevExpress.XtraPrinting.Preview.DocumentViewer();
            this.remoteDocumentSource1 = new DevExpress.ReportServer.Printing.RemoteDocumentSource();
            this.SuspendLayout();
            // 
            // printPreviewDialog1
            // 
            this.printPreviewDialog1.AutoScrollMargin = new System.Drawing.Size(0, 0);
            this.printPreviewDialog1.AutoScrollMinSize = new System.Drawing.Size(0, 0);
            this.printPreviewDialog1.ClientSize = new System.Drawing.Size(400, 300);
            this.printPreviewDialog1.Enabled = true;
            this.printPreviewDialog1.Icon = ((System.Drawing.Icon)(resources.GetObject("printPreviewDialog1.Icon")));
            this.printPreviewDialog1.Name = "printPreviewDialog1";
            this.printPreviewDialog1.Visible = false;
            // 
            // documentViewer1
            // 
            this.documentViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.documentViewer1.IsMetric = false;
            this.documentViewer1.Location = new System.Drawing.Point(0, 0);
            this.documentViewer1.Name = "documentViewer1";
            this.documentViewer1.Size = new System.Drawing.Size(800, 450);
            this.documentViewer1.TabIndex = 0;
            // 
            // documentViewer2
            // 
            this.documentViewer2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.documentViewer2.IsMetric = false;
            this.documentViewer2.Location = new System.Drawing.Point(0, 0);
            this.documentViewer2.Name = "documentViewer2";
            this.documentViewer2.Size = new System.Drawing.Size(800, 450);
            this.documentViewer2.TabIndex = 1;
            // 
            // documentViewer3
            // 
            this.documentViewer3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.documentViewer3.DocumentSource = this.remoteDocumentSource1;
            this.documentViewer3.IsMetric = false;
            this.documentViewer3.Location = new System.Drawing.Point(0, 0);
            this.documentViewer3.Name = "documentViewer3";
            this.documentViewer3.Size = new System.Drawing.Size(800, 450);
            this.documentViewer3.TabIndex = 2;
            // 
            // frmFillFileWordOfficeMicrosoft
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.documentViewer3);
            this.Controls.Add(this.documentViewer2);
            this.Controls.Add(this.documentViewer1);
            this.Name = "frmFillFileWordOfficeMicrosoft";
            this.Text = "frmFillFileWordOfficeMicrosoft";
            this.Load += new System.EventHandler(this.frmFillFileWordOfficeMicrosoft_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PrintPreviewDialog printPreviewDialog1;
        private DevExpress.XtraPrinting.Preview.DocumentViewer documentViewer1;
        private DevExpress.XtraPrinting.Preview.DocumentViewer documentViewer2;
        private DevExpress.XtraPrinting.Preview.DocumentViewer documentViewer3;
        private DevExpress.ReportServer.Printing.RemoteDocumentSource remoteDocumentSource1;
    }
}