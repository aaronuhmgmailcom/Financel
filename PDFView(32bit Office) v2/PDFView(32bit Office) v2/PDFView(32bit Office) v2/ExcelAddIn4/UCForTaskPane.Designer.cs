namespace ExcelAddIn4
{
    partial class UCForTaskPane
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lbTest = new System.Windows.Forms.Label();
            this.Panel1 = new System.Windows.Forms.Panel();
            this.rbXPDF = new System.Windows.Forms.RadioButton();
            this.rbGS = new System.Windows.Forms.RadioButton();
            this.btOCR = new System.Windows.Forms.Button();
            this.Panel2 = new System.Windows.Forms.Panel();
            this.pdfViewer1 = new PDFView.PDFViewer();
            this.panel3 = new System.Windows.Forms.Panel();
            this.webBrowser1 = new System.Windows.Forms.WebBrowser();
            this.Panel1.SuspendLayout();
            this.Panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // lbTest
            // 
            this.lbTest.AutoSize = true;
            this.lbTest.Location = new System.Drawing.Point(13, 15);
            this.lbTest.Name = "lbTest";
            this.lbTest.Size = new System.Drawing.Size(0, 13);
            this.lbTest.TabIndex = 1;
            // 
            // Panel1
            // 
            this.Panel1.Controls.Add(this.rbXPDF);
            this.Panel1.Controls.Add(this.rbGS);
            this.Panel1.Controls.Add(this.btOCR);
            this.Panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.Panel1.Location = new System.Drawing.Point(0, 0);
            this.Panel1.Name = "Panel1";
            this.Panel1.Size = new System.Drawing.Size(894, 0);
            this.Panel1.TabIndex = 0;
            // 
            // rbXPDF
            // 
            this.rbXPDF.AutoSize = true;
            this.rbXPDF.Checked = true;
            this.rbXPDF.Location = new System.Drawing.Point(783, 10);
            this.rbXPDF.Name = "rbXPDF";
            this.rbXPDF.Size = new System.Drawing.Size(53, 17);
            this.rbXPDF.TabIndex = 0;
            this.rbXPDF.TabStop = true;
            this.rbXPDF.Text = "XPDF";
            this.rbXPDF.UseVisualStyleBackColor = true;
            this.rbXPDF.Visible = false;
            // 
            // rbGS
            // 
            this.rbGS.AutoSize = true;
            this.rbGS.Location = new System.Drawing.Point(756, 7);
            this.rbGS.Name = "rbGS";
            this.rbGS.Size = new System.Drawing.Size(80, 17);
            this.rbGS.TabIndex = 1;
            this.rbGS.Text = "GhostScript";
            this.rbGS.UseVisualStyleBackColor = true;
            this.rbGS.Visible = false;
            // 
            // btOCR
            // 
            this.btOCR.Location = new System.Drawing.Point(842, 4);
            this.btOCR.Name = "btOCR";
            this.btOCR.Size = new System.Drawing.Size(49, 23);
            this.btOCR.TabIndex = 3;
            this.btOCR.Text = "OCR";
            this.btOCR.UseVisualStyleBackColor = true;
            this.btOCR.Visible = false;
            // 
            // Panel2
            // 
            this.Panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Panel2.Location = new System.Drawing.Point(0, 0);
            this.Panel2.Name = "Panel2";
            this.Panel2.Size = new System.Drawing.Size(894, 150);
            this.Panel2.TabIndex = 1;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.webBrowser1);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel3.Location = new System.Drawing.Point(399, 226);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(200, 100);
            this.panel3.TabIndex = 0;
            // 
            // webBrowser1
            // 
            this.webBrowser1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webBrowser1.Location = new System.Drawing.Point(0, 0);
            this.webBrowser1.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.Size = new System.Drawing.Size(200, 100);
            this.webBrowser1.TabIndex = 0;
            // 
            // pdfViewer1
            // 
            this.pdfViewer1.AllowBookmarks = true;
            this.pdfViewer1.AutoSize = true;
            this.pdfViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pdfViewer1.FileName = null;
            this.pdfViewer1.Location = new System.Drawing.Point(0, 0);
            this.pdfViewer1.Name = "pdfViewer1";
            this.pdfViewer1.Size = new System.Drawing.Size(894, 150);
            this.pdfViewer1.TabIndex = 0;
            this.pdfViewer1.UseXPDF = true;
            // 
            // UCForTaskPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.Controls.Add(this.pdfViewer1);
            this.Controls.Add(this.lbTest);
            this.Controls.Add(this.Panel2);
            this.Controls.Add(this.Panel1);
            this.Controls.Add(this.panel3);
            this.Name = "UCForTaskPane";
            this.Size = new System.Drawing.Size(894, 150);
            this.Panel1.ResumeLayout(false);
            this.Panel1.PerformLayout();
            this.Panel2.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbTest;
        public System.Windows.Forms.Panel Panel1;
        public System.Windows.Forms.Panel Panel2;
        private System.Windows.Forms.RadioButton rbGS;
        private System.Windows.Forms.RadioButton rbXPDF;
        public PDFView.PDFViewer pdfViewer1;
        private System.Windows.Forms.Button btOCR;
        public System.Windows.Forms.Panel panel3;
        public System.Windows.Forms.WebBrowser webBrowser1;
    }
}
