using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using System.Configuration;

namespace ExcelAddIn4
{
    public partial class UCForTaskPane : UserControl
    {
        public UCForTaskPane()
        {
            InitializeComponent();
        }
        //private void dgv_CellMouseDoubleClick(object sender, DataGridViewCellEventArgs e)
        //{
        //DataGridView dgv2 = sender as DataGridView;
        //if (e.ColumnIndex >= 0 & e.RowIndex >= 0)
        //{
        //    this.Controls.Clear();
        //    InitializeComponent();
        //    string name = dgv2[1, e.RowIndex].Value.ToString();
        //    this.pdfViewer1.FileName = Finance_Tools.GetAppConfig("BasePDFDocumentsPath") + name;
        //    this.pdfViewer1.AutoSize = true;
        ////    string name = dgv2[1, e.RowIndex].Value.ToString();
        ////    WebBrowser wb = new WebBrowser();
        ////    wb.Dock = DockStyle.Fill;
        ////    wb.Navigate(@"C:\Rstest\RSTest_Documents\" + name);
        ////    this.Controls.Add(wb);
        ////    dgv2.Visible = false;
        ////    wb.Visible = true;
        //}

        //}
        //private void rbGS_CheckedChanged(object sender, EventArgs e)
        //{
        //    if (rbXPDF.Checked)
        //    {
        //        pdfViewer1.UseXPDF = true;
        //    }
        //    else
        //    {
        //        pdfViewer1.UseXPDF = false;
        //    }

        //    pdfViewer1.FileName = pdfViewer1.FileName;
        //}
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            pdfViewer1.Dispose();
        }
        //private void btOCR_Click(object sender, EventArgs e)
        //{
        //    //MsgBox(PdfViewer1.OCRCurrentPage);
        //}
        //private void btnBack_Click(object sender, EventArgs e)
        //{
        //this.Controls.Clear();
        //DirectoryInfo di = new DirectoryInfo(Finance_Tools.GetAppConfig("BasePDFDocumentsPath"));
        //FileInfo[] myfile = di.GetFiles("*.pdf");
        //DataGridView dgv = new DataGridView();
        //dgv.Columns.Add("Scan_Date", "Scan Date");
        //dgv.Columns.Add("Scan_Name", "Description");
        //dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        //foreach (FileInfo f in myfile)
        //{
        //    dgv.Rows.Add(f.CreationTime, f.Name);
        //}
        //dgv.Dock = DockStyle.Fill;
        //dgv.Visible = true;
        //dgv.CellDoubleClick += new DataGridViewCellEventHandler(dgv_CellMouseDoubleClick);
        //this.Controls.Add(dgv); 
        //}
    }
}
