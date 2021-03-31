namespace Pi.Export.Net
{
    partial class frmMenu
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
            this.components = new System.ComponentModel.Container();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.contextMenuStrip2 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.tooltipChucNang = new System.Windows.Forms.ToolStripMenuItem();
            this.btnDocFileExcel = new System.Windows.Forms.ToolStripMenuItem();
            this.btnDemoFillFileWord = new System.Windows.Forms.ToolStripMenuItem();
            this.fileFileWordDùngMicrosoftOfficeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.btnLoadDataExcel = new System.Windows.Forms.ToolStripMenuItem();
            this.thienToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.btnDocObjectWord = new System.Windows.Forms.ToolStripMenuItem();
            this.btnExecl = new System.Windows.Forms.ToolStripMenuItem();
            this.đứcAnhToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.btnTest1 = new System.Windows.Forms.ToolStripMenuItem();
            this.testExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.phuocToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.piToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(61, 4);
            // 
            // contextMenuStrip2
            // 
            this.contextMenuStrip2.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.contextMenuStrip2.Name = "contextMenuStrip2";
            this.contextMenuStrip2.Size = new System.Drawing.Size(61, 4);
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tooltipChucNang,
            this.thienToolStripMenuItem,
            this.đứcAnhToolStripMenuItem,
            this.phuocToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(5, 2, 0, 2);
            this.menuStrip1.Size = new System.Drawing.Size(800, 28);
            this.menuStrip1.TabIndex = 3;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // tooltipChucNang
            // 
            this.tooltipChucNang.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnDocFileExcel,
            this.btnDemoFillFileWord,
            this.fileFileWordDùngMicrosoftOfficeToolStripMenuItem,
            this.btnLoadDataExcel});
            this.tooltipChucNang.Name = "tooltipChucNang";
            this.tooltipChucNang.Size = new System.Drawing.Size(92, 22);
            this.tooltipChucNang.Text = "Chức năng";
            // 
            // btnDocFileExcel
            // 
            this.btnDocFileExcel.Name = "btnDocFileExcel";
            this.btnDocFileExcel.Size = new System.Drawing.Size(312, 26);
            this.btnDocFileExcel.Text = "Đọc file excel";
            // 
            // btnDemoFillFileWord
            // 
            this.btnDemoFillFileWord.Name = "btnDemoFillFileWord";
            this.btnDemoFillFileWord.Size = new System.Drawing.Size(312, 26);
            this.btnDemoFillFileWord.Text = "Fill file word";
            this.btnDemoFillFileWord.Click += new System.EventHandler(this.btnDemoFillFileWord_Click);
            // 
            // fileFileWordDùngMicrosoftOfficeToolStripMenuItem
            // 
            this.fileFileWordDùngMicrosoftOfficeToolStripMenuItem.Name = "fileFileWordDùngMicrosoftOfficeToolStripMenuItem";
            this.fileFileWordDùngMicrosoftOfficeToolStripMenuItem.Size = new System.Drawing.Size(312, 26);
            this.fileFileWordDùngMicrosoftOfficeToolStripMenuItem.Text = "File file word dùng Microsoft Office";
            this.fileFileWordDùngMicrosoftOfficeToolStripMenuItem.Click += new System.EventHandler(this.fileFileWordDùngMicrosoftOfficeToolStripMenuItem_Click);
            // 
            // btnLoadDataExcel
            // 
            this.btnLoadDataExcel.Name = "btnLoadDataExcel";
            this.btnLoadDataExcel.Size = new System.Drawing.Size(312, 26);
            this.btnLoadDataExcel.Text = "Load Data Excel";
            this.btnLoadDataExcel.Click += new System.EventHandler(this.btnLoadDataExcel_Click);
            // 
            // thienToolStripMenuItem
            // 
            this.thienToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnDocObjectWord,
            this.btnExecl});
            this.thienToolStripMenuItem.Name = "thienToolStripMenuItem";
            this.thienToolStripMenuItem.Size = new System.Drawing.Size(58, 22);
            this.thienToolStripMenuItem.Text = "Thiện";
            this.thienToolStripMenuItem.Click += new System.EventHandler(this.thienToolStripMenuItem_Click);
            // 
            // btnDocObjectWord
            // 
            this.btnDocObjectWord.Name = "btnDocObjectWord";
            this.btnDocObjectWord.Size = new System.Drawing.Size(309, 26);
            this.btnDocObjectWord.Text = "Đọc danh sách object từ file word";
            this.btnDocObjectWord.Click += new System.EventHandler(this.btnDocObjectWord_Click);
            // 
            // btnExecl
            // 
            this.btnExecl.Name = "btnExecl";
            this.btnExecl.Size = new System.Drawing.Size(309, 26);
            this.btnExecl.Text = "Đọc file excel";
            this.btnExecl.Click += new System.EventHandler(this.btnExecl_Click);
            // 
            // đứcAnhToolStripMenuItem
            // 
            this.đứcAnhToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnTest1,
            this.testExcelToolStripMenuItem});
            this.đứcAnhToolStripMenuItem.Name = "đứcAnhToolStripMenuItem";
            this.đứcAnhToolStripMenuItem.Size = new System.Drawing.Size(78, 22);
            this.đứcAnhToolStripMenuItem.Text = "Đức Anh";
            // 
            // btnTest1
            // 
            this.btnTest1.Name = "btnTest1";
            this.btnTest1.Size = new System.Drawing.Size(224, 26);
            this.btnTest1.Text = "Test1";
            this.btnTest1.Click += new System.EventHandler(this.btnTest1_Click);
            // 
            // testExcelToolStripMenuItem
            // 
            this.testExcelToolStripMenuItem.Name = "testExcelToolStripMenuItem";
            this.testExcelToolStripMenuItem.Size = new System.Drawing.Size(224, 26);
            this.testExcelToolStripMenuItem.Text = "TestExcel";
            this.testExcelToolStripMenuItem.Click += new System.EventHandler(this.testExcelToolStripMenuItem_Click);
            // 
            // phuocToolStripMenuItem
            // 
            this.phuocToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.piToolStripMenuItem});
            this.phuocToolStripMenuItem.Name = "phuocToolStripMenuItem";
            this.phuocToolStripMenuItem.Size = new System.Drawing.Size(61, 24);
            this.phuocToolStripMenuItem.Text = "Phuoc";
            // 
            // piToolStripMenuItem
            // 
            this.piToolStripMenuItem.Name = "piToolStripMenuItem";
            this.piToolStripMenuItem.Size = new System.Drawing.Size(224, 26);
            this.piToolStripMenuItem.Text = "Pi";
            this.piToolStripMenuItem.Click += new System.EventHandler(this.piToolStripMenuItem_Click);
            // 
            // frmMenu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.menuStrip1);
            this.IsMdiContainer = true;
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "frmMenu";
            this.Text = "Menu";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip2;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem tooltipChucNang;
        private System.Windows.Forms.ToolStripMenuItem btnDocFileExcel;
        private System.Windows.Forms.ToolStripMenuItem btnDemoFillFileWord;
        private System.Windows.Forms.ToolStripMenuItem fileFileWordDùngMicrosoftOfficeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem thienToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem btnDocObjectWord;
        private System.Windows.Forms.ToolStripMenuItem đứcAnhToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem btnTest1;
        private System.Windows.Forms.ToolStripMenuItem btnExecl;
        private System.Windows.Forms.ToolStripMenuItem testExcelToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem btnLoadDataExcel;
        private System.Windows.Forms.ToolStripMenuItem phuocToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem piToolStripMenuItem;
    }
}

