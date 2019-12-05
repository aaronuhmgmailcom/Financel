using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
             PdfViewer1.SelectFile();
             TextBox1.Text = PdfViewer1.FileName;
        }

        private void Panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void rbGS_CheckedChanged(object sender, EventArgs e)
        {
                if(rbXPDF.Checked)
                {
                    PdfViewer1.UseXPDF = true;
                }
                else
                {
                    PdfViewer1.UseXPDF = false;
                }

                PdfViewer1.FileName = PdfViewer1.FileName;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            PdfViewer1.Dispose();
        }

        private void btOCR_Click(object sender, EventArgs e)
        {
            //MsgBox(PdfViewer1.OCRCurrentPage);
        }


    }
}
