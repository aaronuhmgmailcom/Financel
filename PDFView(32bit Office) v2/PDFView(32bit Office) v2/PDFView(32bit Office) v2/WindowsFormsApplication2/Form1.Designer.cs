namespace WindowsFormsApplication2
{
    partial class Form1
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
            this.OpenFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.Panel1 = new System.Windows.Forms.Panel();
            this.btOCR = new System.Windows.Forms.Button();
            this.Panel3 = new System.Windows.Forms.Panel();
            this.rbGS = new System.Windows.Forms.RadioButton();
            this.rbXPDF = new System.Windows.Forms.RadioButton();
            this.TextBox1 = new System.Windows.Forms.TextBox();
            this.Button1 = new System.Windows.Forms.Button();
            this.Panel2 = new System.Windows.Forms.Panel();
            this.PdfViewer1 = new PDFView.PDFViewer();
            this.Panel1.SuspendLayout();
            this.Panel3.SuspendLayout();
            this.Panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // OpenFileDialog1
            // 
            this.OpenFileDialog1.FileName = "OpenFileDialog1";
            // 
            // Panel1
            // 
            this.Panel1.Controls.Add(this.btOCR);
            this.Panel1.Controls.Add(this.Panel3);
            this.Panel1.Controls.Add(this.TextBox1);
            this.Panel1.Controls.Add(this.Button1);
            this.Panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.Panel1.Location = new System.Drawing.Point(0, 0);
            this.Panel1.Name = "Panel1";
            this.Panel1.Size = new System.Drawing.Size(539, 30);
            this.Panel1.TabIndex = 0;
            // 
            // btOCR
            // 
            this.btOCR.Location = new System.Drawing.Point(255, 3);
            this.btOCR.Name = "btOCR";
            this.btOCR.Size = new System.Drawing.Size(49, 23);
            this.btOCR.TabIndex = 3;
            this.btOCR.Text = "OCR";
            this.btOCR.UseVisualStyleBackColor = true;
            this.btOCR.Click += new System.EventHandler(this.btOCR_Click);
            // 
            // Panel3
            // 
            this.Panel3.Controls.Add(this.rbGS);
            this.Panel3.Controls.Add(this.rbXPDF);
            this.Panel3.Location = new System.Drawing.Point(307, 3);
            this.Panel3.Name = "Panel3";
            this.Panel3.Size = new System.Drawing.Size(148, 23);
            this.Panel3.TabIndex = 2;
            this.Panel3.Paint += new System.Windows.Forms.PaintEventHandler(this.Panel3_Paint);
            // 
            // rbGS
            // 
            this.rbGS.AutoSize = true;
            this.rbGS.Location = new System.Drawing.Point(62, 3);
            this.rbGS.Name = "rbGS";
            this.rbGS.Size = new System.Drawing.Size(80, 17);
            this.rbGS.TabIndex = 1;
            this.rbGS.Text = "GhostScript";
            this.rbGS.UseVisualStyleBackColor = true;
            this.rbGS.CheckedChanged += new System.EventHandler(this.rbGS_CheckedChanged);
            // 
            // rbXPDF
            // 
            this.rbXPDF.AutoSize = true;
            this.rbXPDF.Checked = true;
            this.rbXPDF.Location = new System.Drawing.Point(3, 3);
            this.rbXPDF.Name = "rbXPDF";
            this.rbXPDF.Size = new System.Drawing.Size(53, 17);
            this.rbXPDF.TabIndex = 0;
            this.rbXPDF.TabStop = true;
            this.rbXPDF.Text = "XPDF";
            this.rbXPDF.UseVisualStyleBackColor = true;
            // 
            // TextBox1
            // 
            this.TextBox1.Enabled = false;
            this.TextBox1.Location = new System.Drawing.Point(4, 4);
            this.TextBox1.Name = "TextBox1";
            this.TextBox1.Size = new System.Drawing.Size(249, 20);
            this.TextBox1.TabIndex = 1;
            // 
            // Button1
            // 
            this.Button1.Location = new System.Drawing.Point(461, 3);
            this.Button1.Name = "Button1";
            this.Button1.Size = new System.Drawing.Size(75, 23);
            this.Button1.TabIndex = 0;
            this.Button1.Text = "Browse...";
            this.Button1.UseVisualStyleBackColor = true;
            this.Button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // Panel2
            // 
            this.Panel2.Controls.Add(this.PdfViewer1);
            this.Panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Panel2.Location = new System.Drawing.Point(0, 30);
            this.Panel2.Name = "Panel2";
            this.Panel2.Size = new System.Drawing.Size(539, 434);
            this.Panel2.TabIndex = 1;
            // 
            // PdfViewer1
            // 
            this.PdfViewer1.AllowBookmarks = true;
            this.PdfViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            //this.PdfViewer1.FileName = null;
            this.PdfViewer1.Location = new System.Drawing.Point(0, 0);
            this.PdfViewer1.Name = "PdfViewer1";
            this.PdfViewer1.Size = new System.Drawing.Size(539, 434);
            this.PdfViewer1.TabIndex = 0;
            this.PdfViewer1.UseXPDF = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(539, 464);
            this.Controls.Add(this.Panel2);
            this.Controls.Add(this.Panel1);
            this.MinimumSize = new System.Drawing.Size(555, 500);
            this.Name = "Form1";
            this.Text = "Free PDF .NET Viewer";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Panel1.ResumeLayout(false);
            this.Panel1.PerformLayout();
            this.Panel3.ResumeLayout(false);
            this.Panel3.PerformLayout();
            this.Panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        //private PDFView.PDFViewer pdfViewer1;

        private System.Windows.Forms.OpenFileDialog OpenFileDialog1;
        private System.Windows.Forms.Panel Panel1;
        private System.Windows.Forms.TextBox TextBox1;
        private System.Windows.Forms.Button Button1;
        private System.Windows.Forms.Panel Panel2;
        private System.Windows.Forms.Panel Panel3;
        private System.Windows.Forms.RadioButton rbGS;
        private System.Windows.Forms.RadioButton rbXPDF;
        private PDFView.PDFViewer PdfViewer1;
        private System.Windows.Forms.Button btOCR;
    }
}

