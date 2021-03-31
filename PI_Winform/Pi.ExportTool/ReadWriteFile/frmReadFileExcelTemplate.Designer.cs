namespace Pi.ExportTool.ReadWriteFile
{
    partial class frmReadFileExcelTemplate
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
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.btnReadFileExcel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtPath = new System.Windows.Forms.TextBox();
            this.btnOpen = new System.Windows.Forms.Button();
            this.tabControl = new System.Windows.Forms.TabControl();
            this.tabLoadSheet = new System.Windows.Forms.TabPage();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.btnTinhGiaTriLonNhat = new System.Windows.Forms.Button();
            this.dgvListMaxValue = new System.Windows.Forms.DataGridView();
            this.dgvListSheet = new System.Windows.Forms.DataGridView();
            this.backgroundProcessing = new System.ComponentModel.BackgroundWorker();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.tabControl.SuspendLayout();
            this.tabLoadSheet.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvListMaxValue)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvListSheet)).BeginInit();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.btnReadFileExcel);
            this.splitContainer1.Panel1.Controls.Add(this.label1);
            this.splitContainer1.Panel1.Controls.Add(this.txtPath);
            this.splitContainer1.Panel1.Controls.Add(this.btnOpen);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.tabControl);
            this.splitContainer1.Size = new System.Drawing.Size(1307, 657);
            this.splitContainer1.SplitterDistance = 74;
            this.splitContainer1.TabIndex = 0;
            // 
            // btnReadFileExcel
            // 
            this.btnReadFileExcel.Location = new System.Drawing.Point(1146, 35);
            this.btnReadFileExcel.Name = "btnReadFileExcel";
            this.btnReadFileExcel.Size = new System.Drawing.Size(151, 28);
            this.btnReadFileExcel.TabIndex = 6;
            this.btnReadFileExcel.Text = "Đọc file";
            this.btnReadFileExcel.UseVisualStyleBackColor = true;
            this.btnReadFileExcel.Click += new System.EventHandler(this.btnReadFileExcel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 38);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(136, 17);
            this.label1.TabIndex = 5;
            this.label1.Text = "Đường dẫn file excel";
            // 
            // txtPath
            // 
            this.txtPath.Location = new System.Drawing.Point(155, 38);
            this.txtPath.Name = "txtPath";
            this.txtPath.Size = new System.Drawing.Size(805, 22);
            this.txtPath.TabIndex = 4;
            // 
            // btnOpen
            // 
            this.btnOpen.Location = new System.Drawing.Point(966, 35);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(165, 28);
            this.btnOpen.TabIndex = 3;
            this.btnOpen.Text = "Chọn đường dẫn file";
            this.btnOpen.UseVisualStyleBackColor = true;
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // tabControl
            // 
            this.tabControl.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl.Controls.Add(this.tabLoadSheet);
            this.tabControl.Location = new System.Drawing.Point(0, 0);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(1307, 576);
            this.tabControl.TabIndex = 0;
            // 
            // tabLoadSheet
            // 
            this.tabLoadSheet.Controls.Add(this.progressBar1);
            this.tabLoadSheet.Controls.Add(this.btnTinhGiaTriLonNhat);
            this.tabLoadSheet.Controls.Add(this.dgvListMaxValue);
            this.tabLoadSheet.Controls.Add(this.dgvListSheet);
            this.tabLoadSheet.Location = new System.Drawing.Point(4, 25);
            this.tabLoadSheet.Name = "tabLoadSheet";
            this.tabLoadSheet.Padding = new System.Windows.Forms.Padding(3);
            this.tabLoadSheet.Size = new System.Drawing.Size(1299, 547);
            this.tabLoadSheet.TabIndex = 0;
            this.tabLoadSheet.Text = "Load Sheets";
            this.tabLoadSheet.UseVisualStyleBackColor = true;
            // 
            // progressBar1
            // 
            this.progressBar1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar1.Location = new System.Drawing.Point(12, 521);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(1279, 23);
            this.progressBar1.TabIndex = 3;
            // 
            // btnTinhGiaTriLonNhat
            // 
            this.btnTinhGiaTriLonNhat.Location = new System.Drawing.Point(699, 6);
            this.btnTinhGiaTriLonNhat.Name = "btnTinhGiaTriLonNhat";
            this.btnTinhGiaTriLonNhat.Size = new System.Drawing.Size(113, 55);
            this.btnTinhGiaTriLonNhat.TabIndex = 2;
            this.btnTinhGiaTriLonNhat.Text = "Lấy giá trị lớn nhất";
            this.btnTinhGiaTriLonNhat.UseVisualStyleBackColor = true;
            this.btnTinhGiaTriLonNhat.Click += new System.EventHandler(this.btnTinhGiaTriLonNhat_Click);
            // 
            // dgvListMaxValue
            // 
            this.dgvListMaxValue.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvListMaxValue.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvListMaxValue.Location = new System.Drawing.Point(818, 6);
            this.dgvListMaxValue.Name = "dgvListMaxValue";
            this.dgvListMaxValue.RowHeadersWidth = 51;
            this.dgvListMaxValue.RowTemplate.Height = 24;
            this.dgvListMaxValue.Size = new System.Drawing.Size(473, 509);
            this.dgvListMaxValue.TabIndex = 1;
            // 
            // dgvListSheet
            // 
            this.dgvListSheet.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.dgvListSheet.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvListSheet.Location = new System.Drawing.Point(12, 6);
            this.dgvListSheet.Name = "dgvListSheet";
            this.dgvListSheet.RowHeadersWidth = 51;
            this.dgvListSheet.RowTemplate.Height = 24;
            this.dgvListSheet.Size = new System.Drawing.Size(681, 509);
            this.dgvListSheet.TabIndex = 0;
            // 
            // backgroundProcessing
            // 
            this.backgroundProcessing.WorkerReportsProgress = true;
            this.backgroundProcessing.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundProcessing_DoWork);
            this.backgroundProcessing.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundProcessing_ProgressChanged);
            this.backgroundProcessing.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundProcessing_RunWorkerCompleted);
            // 
            // frmReadFileExcelTemplate
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1307, 657);
            this.Controls.Add(this.splitContainer1);
            this.Name = "frmReadFileExcelTemplate";
            this.Text = "frmReadFileExcelTemplate";
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel1.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.tabControl.ResumeLayout(false);
            this.tabLoadSheet.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvListMaxValue)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvListSheet)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.TextBox txtPath;
        private System.Windows.Forms.Button btnOpen;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage tabLoadSheet;
        private System.Windows.Forms.Button btnTinhGiaTriLonNhat;
        private System.Windows.Forms.DataGridView dgvListMaxValue;
        private System.Windows.Forms.DataGridView dgvListSheet;
        private System.Windows.Forms.Button btnReadFileExcel;
        private System.ComponentModel.BackgroundWorker backgroundProcessing;
        private System.Windows.Forms.ProgressBar progressBar1;
    }
}