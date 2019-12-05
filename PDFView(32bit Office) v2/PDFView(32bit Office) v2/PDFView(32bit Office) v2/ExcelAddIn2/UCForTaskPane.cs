using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelAddIn2
{
    public partial class UCForTaskPane : UserControl
    {
        public UCForTaskPane()
        {
            InitializeComponent();
        }


        private void Button1_Click(object sender, EventArgs e)
        {
            pdfViewer1.SelectFile();
            TextBox1.Text = pdfViewer1.FileName;
        }

        private void Panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void rbGS_CheckedChanged(object sender, EventArgs e)
        {
            if (rbXPDF.Checked)
            {
                pdfViewer1.UseXPDF = true;
            }
            else
            {
                pdfViewer1.UseXPDF = false;
            }

            pdfViewer1.FileName = pdfViewer1.FileName;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            pdfViewer1.Dispose();
        }

        private void btOCR_Click(object sender, EventArgs e)
        {
            //MsgBox(PdfViewer1.OCRCurrentPage);
        }

    }
}
