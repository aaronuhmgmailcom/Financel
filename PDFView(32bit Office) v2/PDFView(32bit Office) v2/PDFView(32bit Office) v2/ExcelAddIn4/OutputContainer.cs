/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<OutputContainer>   
 * Author：Peter.uhm  (yanb@shinetechchina.com)
 * Modify date：2016.04
 * Modify date：2016.09
 * Version : 2.0.0.2
 */
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Diagnostics;
using ExcelAddIn4.Common;
using System.IO;
using System.Xml;

namespace ExcelAddIn4
{
    public partial class OutputContainer : UserControl
    {
        /// <summary>
        /// 
        /// </summary>
        internal DataGridView dgvLD;
        /// <summary>
        /// 
        /// </summary>
        internal DataGridView dgvSaveOptions;
        /// <summary>
        /// 
        /// </summary>
        internal DataGridView dgvCreateTextFile;
        /// <summary>
        /// 
        /// </summary>
        internal static string filter;
        /// <summary>
        /// 
        /// </summary>
        internal List<Specialist> finallist = new List<Specialist>();
        /// <summary>
        /// 
        /// </summary>
        internal List<ExcelAddIn4.Common2.Specialist> TransUpdFinallist = new List<ExcelAddIn4.Common2.Specialist>();
        /// <summary>
        /// 
        /// </summary>
        internal List<RowCreateTextFile> finallistCTF = new List<RowCreateTextFile>();
        /// <summary>
        /// 
        /// </summary>
        internal static bool isTransUpdFlag = false;
        /// <summary>
        /// 
        /// </summary>
        internal static string searchStatus = string.Empty;
        /// <summary>
        /// 
        /// </summary>
        internal static string updateStatus = string.Empty;
        /// <summary>
        /// 
        /// </summary>
        internal static string updateStatusForPost = string.Empty;
        /// <summary>
        /// 
        /// </summary>
        //bool canChange = true;
        /// <summary>
        /// 
        /// </summary>
        internal List<KeyValuePair<int, string>> NameValueCollection;
        /// <summary>
        /// 
        /// </summary>
        internal List<KeyValuePair<int, string>> NameValueUpdate;
        /// <summary>
        /// 
        /// </summary>
        internal bool isFromBindingFlag = false;
        /// <summary>
        /// 
        /// </summary>
        internal static Finance_Tools ft
        {
            get { return new Finance_Tools(); }
        }
        /// <summary>
        /// 
        /// </summary>
        internal static RSFinanceToolsEntities db
        {
            get { return new RSFinanceToolsEntities(); }
        }
        public bool bTab0HasLoad = false;
        public bool bTab1HasLoad = false;
        public bool bTab2HasLoad = false;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.tabControl1.SelectedIndex == 1 && bTab1HasLoad == false)
            {
                //DataTable dt = ft.GetReportsViaTemplatePath();
                //if (dt.Rows.Count > 0) canChange = false;
                //BindSaveOptions();
                BasePage.VerifyWriteButton(SessionInfo.UserInfo.FileName, this.button2);
                bTab1HasLoad = true;
            }
            else if (this.tabControl1.SelectedIndex == 0 && bTab0HasLoad == false)
            {
                BasePage.VerifyWriteButton(SessionInfo.UserInfo.FileName, this.button1);
                BasePage.VerifyWriteButton(SessionInfo.UserInfo.FileName, this.btnTestJournal);
                BasePage.VerifyWriteButton(SessionInfo.UserInfo.FileName, this.btnSetMax);
                BasePage.VerifyWriteButton(SessionInfo.UserInfo.FileName, button3);
                bTab0HasLoad = true;
            }
            else if (this.tabControl1.SelectedIndex == 2 && bTab2HasLoad == false)
            {
                BasePage.VerifyWriteButton(SessionInfo.UserInfo.FileName, this.btnTestCTF);
                BasePage.VerifyWriteButton(SessionInfo.UserInfo.FileName, this.CTF_btnSave);
                bTab2HasLoad = true;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        public OutputContainer()
        {
            InitializeComponent();
            BindLineDetailDGV();
            BindSaveOptions();
            tabControl1_SelectedIndexChanged(null, null);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dt"></param>
        public void BindCreateTextFileDGV(DataTable dt)
        {
            try
            {
                dgvCreateTextFile = new DataGridView();
                dgvCreateTextFile.Columns.Clear();
                dgvCreateTextFile.Columns.Add("ReferenceNumber", "Ref");
                dgvCreateTextFile.Columns["ReferenceNumber"].DataPropertyName = "ReferenceNumber";
                dgvCreateTextFile.Columns.Add("LineIndicator", "Line Indicator");
                dgvCreateTextFile.Columns["LineIndicator"].DataPropertyName = "LineIndicator";
                dgvCreateTextFile.Columns.Add("StartinginCell", "StartinginCell");
                dgvCreateTextFile.Columns["StartinginCell"].DataPropertyName = "StartinginCell";
                if (!string.IsNullOrEmpty(dt.Rows[0]["TextFileName"].ToString()))
                {
                    DataGridViewButtonColumn dgvb = new DataGridViewButtonColumn();
                    dgvb.HeaderText = "SavePath";
                    dgvb.Name = "SavePath";
                    dgvb.Text = "Browse...";
                    dgvb.DataPropertyName = "SavePath";
                    dgvCreateTextFile.Columns.Add(dgvb);
                    dgvCreateTextFile.Columns.Add("SaveName", "SaveName");
                    dgvCreateTextFile.Columns["SaveName"].DataPropertyName = "SaveName";
                    DataGridViewComboBoxColumn combox = new DataGridViewComboBoxColumn();
                    combox.HeaderText = "IncludeHeaderRow";
                    combox.Name = "IncludeHeaderRow";
                    combox.Items.Add("True");
                    combox.Items.Add("False");
                    combox.DataPropertyName = "IncludeHeaderRow";
                    combox.SortMode = DataGridViewColumnSortMode.NotSortable;
                    dgvCreateTextFile.Columns.Add(combox);
                }
                int count = 1;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (bool.Parse(dt.Rows[i]["Visible"].ToString()) == true)
                    {
                        if (dgvCreateTextFile.Columns.Contains(dt.Rows[i]["Field"].ToString()))
                        {
                            dgvCreateTextFile.Columns.Add(dt.Rows[i]["Field"].ToString() + count, dt.Rows[i]["FriendlyName"].ToString());
                            dgvCreateTextFile.Columns[dt.Rows[i]["Field"].ToString() + count].DataPropertyName = "Column" + count;
                            dgvCreateTextFile.Columns[dt.Rows[i]["Field"].ToString() + count].Tag = dt.Rows[i]["DefaultValue"].ToString() + ",,," + dt.Rows[i]["Mandatory"].ToString() + ",,," + dt.Rows[i]["Separator"].ToString() + ",,," + dt.Rows[i]["TextLength"].ToString() + ",,," + dt.Rows[i]["Prefix"].ToString() + ",,," + dt.Rows[i]["Suffix"].ToString() + ",,," + dt.Rows[i]["RemoveCharacters"].ToString() + ",,," + dt.Rows[i]["Parent"].ToString() + ",,,";
                        }
                        else
                        {
                            dgvCreateTextFile.Columns.Add(dt.Rows[i]["Field"].ToString(), dt.Rows[i]["FriendlyName"].ToString());
                            dgvCreateTextFile.Columns[dt.Rows[i]["Field"].ToString()].DataPropertyName = "Column" + count;
                            dgvCreateTextFile.Columns[dt.Rows[i]["Field"].ToString()].Tag = dt.Rows[i]["DefaultValue"].ToString() + ",,," + dt.Rows[i]["Mandatory"].ToString() + ",,," + dt.Rows[i]["Separator"].ToString() + ",,," + dt.Rows[i]["TextLength"].ToString() + ",,," + dt.Rows[i]["Prefix"].ToString() + ",,," + dt.Rows[i]["Suffix"].ToString() + ",,," + dt.Rows[i]["RemoveCharacters"].ToString() + ",,," + dt.Rows[i]["Parent"].ToString() + ",,,";
                        }
                        count++;
                    }
                }
                dgvCreateTextFile.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
                dgvCreateTextFile.AutoGenerateColumns = false;
                dgvCreateTextFile.ColumnHeadersHeight = 40;
                dgvCreateTextFile.Dock = DockStyle.Fill;
                for (int i = 0; i < dgvCreateTextFile.Columns.Count; i++)
                    dgvCreateTextFile.Columns[i].Width = 55;
                dgvCreateTextFile.Columns["ReferenceNumber"].Width = 35;
                dgvCreateTextFile.AllowUserToAddRows = true;
                dgvCreateTextFile.CellDoubleClick += new DataGridViewCellEventHandler(dgvctf_CellMouseDoubleClick);
                dgvCreateTextFile.CellClick += new DataGridViewCellEventHandler(dgvctf_CellMouseClick);
                dgvCreateTextFile.RowPostPaint += new DataGridViewRowPostPaintEventHandler(dgvctf_RowPostPaint);
                dgvCreateTextFile.CellFormatting += new DataGridViewCellFormattingEventHandler(dgvCreateTextFile_CellFormatting);
                dgvCreateTextFile.NotifyCurrentCellDirty(false);
                dgvCreateTextFile.DataError += new DataGridViewDataErrorEventHandler(dgvctf_DataError);
                dgvCreateTextFile.EditMode = DataGridViewEditMode.EditOnKeystroke;
                dgvCreateTextFile.KeyDown += new KeyEventHandler(dgvctf_KeyDown);
                dgvCreateTextFile.CellContentClick += new DataGridViewCellEventHandler(dgvCreateTextFile_CellContentClick);
                ((System.ComponentModel.ISupportInitialize)(this.dgvCreateTextFile)).BeginInit();
                ((System.ComponentModel.ISupportInitialize)(this.dataGridViewColumnHeaderEditor1)).BeginInit();
                this.dataGridViewColumnHeaderEditor1.TargetControl = this.dgvCreateTextFile;
                BindDataCTF();
                this.panel17.AutoSize = true;
                this.panel17.Controls.Clear();
                this.panel17.Controls.Add(dgvCreateTextFile);
                dgvCreateTextFile.RowHeadersWidth = 55;
                ((System.ComponentModel.ISupportInitialize)(this.dgvCreateTextFile)).EndInit();
                ((System.ComponentModel.ISupportInitialize)(this.dataGridViewColumnHeaderEditor1)).EndInit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, " From Create Text File Tab, Output settings Error");
                LogHelper.WriteLog(typeof(OutputContainer), ex.Message + " From Create Text File Tab, Output settings Error");
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvCreateTextFile_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if ((e.ColumnIndex == 3) && (e.RowIndex != -1) && cbXMLOrText.Text != "XML")
            {
                DialogResult drctf = fbdAd_UpdateFolder.ShowDialog();
                if (drctf == DialogResult.OK)
                {
                    dgvCreateTextFile[e.ColumnIndex, e.RowIndex].Value = fbdAd_UpdateFolder.SelectedPath;
                    dgvCreateTextFile[e.ColumnIndex, e.RowIndex].ToolTipText = fbdAd_UpdateFolder.SelectedPath;
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvCreateTextFile_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex == 0)
                dgvCreateTextFile.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = dgvCreateTextFile.RowHeadersDefaultCellStyle.BackColor;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvctf_DataError(object sender, DataGridViewDataErrorEventArgs e) { }
        /// <summary>
        /// 
        /// </summary>
        private void BindDataCTF()
        {
            DataTable dt = new DataTable();
            if (cbXMLOrText.Text == "XML")
                dt = ft.GetCreateTextFileDataFromDB(cbItems.Text.Substring(0, cbItems.Text.LastIndexOf(",")), cbItems.Text.Substring(cbItems.Text.LastIndexOf(",") + 1));
            else
                dt = ft.GetCreateTextFileDataFromDB(cbItems.Text);
            if (dt.Rows.Count > 0)
            {
                List<RowCreateTextFile> list = GetCTFDataList(dt);
                for (int k = 0; k < list.Count; k++)
                    dgvCreateTextFile.Rows.Add(new DataGridViewRow());
                for (int j = 0; j < dt.Rows.Count; j++)
                    for (int i = 0; i < dgvCreateTextFile.Columns.Count; i++)
                        dgvCreateTextFile.Rows[j].Cells[i].Value = DataConversionTools.GetPropertyValue(dgvCreateTextFile.Columns[i].DataPropertyName, list[j]);//initialize operation data in dgvCreateTextFile dgv
            }
            if (cbXMLOrText.Text == "XML")
                for (int i = 0; i < dgvCreateTextFile.Columns.Count; i++)
                {
                    string type = ft.GetSectionFromDB(cbItems.Text.Substring(0, cbItems.Text.LastIndexOf(",")), cbItems.Text.Substring(cbItems.Text.LastIndexOf(",") + 1), dgvCreateTextFile.Columns[i].Name, dgvCreateTextFile.Columns[i].HeaderText);
                    if (type == "Header")
                        dgvCreateTextFile.Columns[i].DefaultCellStyle.BackColor = System.Drawing.Color.Ivory;
                    else if (type == "Footer")
                        dgvCreateTextFile.Columns[i].DefaultCellStyle.BackColor = System.Drawing.Color.Beige;
                }
            this.panel16.AutoSize = true;
            this.panel17.Controls.Add(dgvCreateTextFile);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        private List<RowCreateTextFile> GetCTFDataList(DataTable dt)
        {
            List<RowCreateTextFile> list = new List<RowCreateTextFile>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                RowCreateTextFile ctfRow = new RowCreateTextFile();
                for (int j = 0; j < dt.Columns.Count; j++)
                    DataConversionTools.SetPropertyValue(dt.Columns[j].ColumnName, dt.Rows[i][j].ToString(), ref ctfRow);  //initialize ctfRow

                list.Add(ctfRow);
            }
            return list;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void dgvctf_KeyDown(object sender, KeyEventArgs e)
        {
            ft.EditingControlWantsInputKey(e.KeyCode, dgvCreateTextFile);
        }
        /// <summary>
        /// Generate datarow number before the Rows
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvctf_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            System.Drawing.Rectangle rectangle = new System.Drawing.Rectangle(e.RowBounds.Location.X,
                e.RowBounds.Location.Y,
                dgvCreateTextFile.RowHeadersWidth - 4,
                e.RowBounds.Height);
            TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(),
                dgvCreateTextFile.RowHeadersDefaultCellStyle.Font,
                rectangle,
                dgvCreateTextFile.RowHeadersDefaultCellStyle.ForeColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvctf_CellMouseClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 0 || (((e.ColumnIndex == 5) || (e.ColumnIndex == 3)) && cbXMLOrText.Text != "XML"))
                    return;
                if (dgvCreateTextFile[0, e.RowIndex].Value == "" || dgvCreateTextFile[0, e.RowIndex].Value == null) dgvCreateTextFile[0, e.RowIndex].Value = e.RowIndex + 1;
                dgvCreateTextFile[e.ColumnIndex, e.RowIndex].Value = "";
                var xlRange = Globals.ThisAddIn.Application.ActiveCell.Address;
                if (xlRange != null)
                {
                    dgvCreateTextFile[e.ColumnIndex, e.RowIndex].Value = xlRange.Replace("$", "");
                    dgvCreateTextFile.Focus();
                }
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvctf_CellMouseDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 0 || (((e.ColumnIndex == 5) || (e.ColumnIndex == 3)) && cbXMLOrText.Text != "XML"))
                    return;
                if (dgvCreateTextFile[e.ColumnIndex, e.RowIndex].Value != null)
                {
                    if (dgvCreateTextFile[e.ColumnIndex, e.RowIndex].Value.ToString().Contains("$"))
                    {
                        string KeyWord = dgvCreateTextFile[e.ColumnIndex, e.RowIndex].Value.ToString().Replace("$", "");
                        dgvCreateTextFile[e.ColumnIndex, e.RowIndex].Value = "";
                        dgvCreateTextFile[e.ColumnIndex, e.RowIndex].Value = KeyWord;
                    }
                    else
                    {
                        string KeyWord = dgvCreateTextFile[e.ColumnIndex, e.RowIndex].Value.ToString();
                        string res = Regex.Replace(KeyWord, @"(\d+)|(\s+) ", " $1 $2 ", RegexOptions.Compiled | RegexOptions.IgnoreCase);
                        KeyWord = "$" + res.Trim().Replace(" ", "$");
                        dgvCreateTextFile[e.ColumnIndex, e.RowIndex].Value = "";
                        dgvCreateTextFile[e.ColumnIndex, e.RowIndex].Value = KeyWord;
                    }
                    dgvCreateTextFile.EndEdit();
                }
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        //private void BindPDFViewer()
        //{
        //    try
        //    {
        //        SessionInfo.UserInfo.Containerpath = (from FT_sett in db.rsTemplateContainers
        //                                              where FT_sett.TemplateID == SessionInfo.UserInfo.File_ftid
        //                                              select FT_sett.ft_relatefilepath).First();
        //        string column = (from FT_sett in db.rsTemplateContainers
        //                         where FT_sett.TemplateID == SessionInfo.UserInfo.File_ftid
        //                         select FT_sett.column).First();

        //        bool? viewFromDB = (from FT_sett in db.rsTemplateContainers
        //                            where FT_sett.TemplateID == SessionInfo.UserInfo.File_ftid
        //                            select FT_sett.FromDB).First();
        //    }
        //    catch { }
        //}
        /// <summary>
        /// 
        /// </summary>
        public void BindSaveOptions()
        {
            try
            {
                dgvSaveOptions = ft.IniSaveOptionsGrd();
                dgvSaveOptions.AllowUserToAddRows = true;
                dgvSaveOptions.CellDoubleClick += new DataGridViewCellEventHandler(dgvSaveOptions_CellMouseDoubleClick);
                dgvSaveOptions.CellClick += new DataGridViewCellEventHandler(dgvSaveOptions_CellMouseClick);
                dgvSaveOptions.RowPostPaint += new DataGridViewRowPostPaintEventHandler(dgvSaveOptions_RowPostPaint);
                dgvSaveOptions.CellFormatting += new DataGridViewCellFormattingEventHandler(dgvSaveOptions_CellFormatting);
                dgvSaveOptions.NotifyCurrentCellDirty(false);
                dgvSaveOptions.EditMode = DataGridViewEditMode.EditOnKeystroke;
                dgvSaveOptions.KeyDown += new KeyEventHandler(dgvSaveOptions_KeyDown);
                BindCriterias();
                this.panel7.AutoSize = true;
                this.panel7.Controls.Clear();
                this.panel7.Controls.Add(dgvSaveOptions);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, " From SaveOptions Tab, Output settings Error");
                LogHelper.WriteLog(typeof(OutputContainer), ex.Message + " From SaveOptions Tab, Output settings Error");
            }
        }
        /// <summary>
        /// 
        /// </summary>
        public void BindCriterias()
        {
            DataTable dt = ft.GetReportCriteria(SessionInfo.UserInfo.File_ftid);
            for (int k = 0; k < dt.Rows.Count; k++)
                dgvSaveOptions.Rows.Add(new DataGridViewRow());
            for (int j = 0; j < dt.Rows.Count; j++)
                for (int i = 0; i < dgvSaveOptions.Columns.Count; i++)
                {
                    dgvSaveOptions.Rows[j].Cells[i].Value = dt.Rows[j][dgvSaveOptions.Columns[i].DataPropertyName].ToString();
                    if (!string.IsNullOrEmpty(dgvSaveOptions.Rows[j].Cells[i].Value.ToString()) && (i != 2) && (i != 4) && (i != 6) && (i != 8) && (i != 10))
                    {
                        if ((i != dgvSaveOptions.Columns.Count - 5))//!canChange &&
                        {
                            dgvSaveOptions.Rows[j].Cells[i].ReadOnly = true;
                            dgvSaveOptions.Rows[j].Cells[i].Style.BackColor = System.Drawing.Color.LightGray;
                        }
                    }
                }
            dgvSaveOptions.Columns[0].ReadOnly = true;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvSaveOptions_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex == 0)
                dgvSaveOptions.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = dgvSaveOptions.RowHeadersDefaultCellStyle.BackColor;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void dgvSaveOptions_KeyDown(object sender, KeyEventArgs e)
        {
            if (dgvSaveOptions.CurrentCell.Style.BackColor == System.Drawing.Color.LightGray) return;
            if (dgvSaveOptions.CurrentCell.ColumnIndex == 0) return;
            ft.EditingControlWantsInputKey(e.KeyCode, dgvSaveOptions);
        }
        /// <summary>
        /// Generate datarow number before the Rows
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvSaveOptions_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            System.Drawing.Rectangle rectangle = new System.Drawing.Rectangle(e.RowBounds.Location.X,
                e.RowBounds.Location.Y,
                dgvSaveOptions.RowHeadersWidth - 4,
                e.RowBounds.Height);
            TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(),
                dgvSaveOptions.RowHeadersDefaultCellStyle.Font,
                rectangle,
                dgvSaveOptions.RowHeadersDefaultCellStyle.ForeColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvSaveOptions_CellMouseClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 0 || e.ColumnIndex == 1 || e.ColumnIndex == 3 || e.ColumnIndex == 5 || e.ColumnIndex == 7 || e.ColumnIndex == 9 || e.ColumnIndex == 11)
                    return;
                if (dgvSaveOptions.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor == System.Drawing.Color.LightGray) return;
                if (dgvSaveOptions[0, e.RowIndex].Value == "" || dgvSaveOptions[0, e.RowIndex].Value == null) dgvSaveOptions[0, e.RowIndex].Value = e.RowIndex + 1;
                dgvSaveOptions[e.ColumnIndex, e.RowIndex].Value = "";
                var xlRange = Globals.ThisAddIn.Application.ActiveCell.Address;
                if (xlRange != null)
                {
                    dgvSaveOptions[e.ColumnIndex, e.RowIndex].Value = xlRange.Replace("$", "");
                    dgvSaveOptions.Focus();
                }
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvSaveOptions_CellMouseDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 0 || e.ColumnIndex == 1 || e.ColumnIndex == 3 || e.ColumnIndex == 5 || e.ColumnIndex == 7 || e.ColumnIndex == 9 || e.ColumnIndex == 11)
                    return;
                if (dgvSaveOptions.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor == System.Drawing.Color.LightGray) return;
                if (dgvSaveOptions[e.ColumnIndex, e.RowIndex].Value != null)
                {
                    if (dgvSaveOptions[e.ColumnIndex, e.RowIndex].Value.ToString().Contains("$"))
                    {
                        string KeyWord = dgvSaveOptions[e.ColumnIndex, e.RowIndex].Value.ToString().Replace("$", "");
                        dgvSaveOptions[e.ColumnIndex, e.RowIndex].Value = "";
                        dgvSaveOptions[e.ColumnIndex, e.RowIndex].Value = KeyWord;
                    }
                    else
                    {
                        string KeyWord = dgvSaveOptions[e.ColumnIndex, e.RowIndex].Value.ToString();
                        string res = Regex.Replace(KeyWord, @"(\d+)|(\s+) ", " $1 $2 ", RegexOptions.Compiled | RegexOptions.IgnoreCase);
                        KeyWord = "$" + res.Trim().Replace(" ", "$");
                        dgvSaveOptions[e.ColumnIndex, e.RowIndex].Value = "";
                        dgvSaveOptions[e.ColumnIndex, e.RowIndex].Value = KeyWord;
                    }
                    dgvSaveOptions.EndEdit();
                }
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        private void BindLineDetailDGV()
        {
            try
            {
                dgvLD = ft.IniGrd();
                dgvLD.AllowUserToAddRows = true;
                dgvLD.CellDoubleClick += new DataGridViewCellEventHandler(dgv_CellMouseDoubleClick);
                dgvLD.CellClick += new DataGridViewCellEventHandler(dgv_CellMouseClick);
                dgvLD.CellMouseDown += new DataGridViewCellMouseEventHandler(dgvLD_CellMouseDown);
                dgvLD.RowPostPaint += new DataGridViewRowPostPaintEventHandler(dgv_RowPostPaint);
                dgvLD.CellFormatting += new DataGridViewCellFormattingEventHandler(dgvLD_CellFormatting);
                dgvLD.CellValueChanged += new DataGridViewCellEventHandler(dgvLD_CellValueChanged);
                dgvLD.NotifyCurrentCellDirty(false);
                dgvLD.EditMode = DataGridViewEditMode.EditOnKeystroke;
                dgvLD.KeyDown += new KeyEventHandler(dgvLD_KeyDown);
                dgvLD.DataBindingComplete += new DataGridViewBindingCompleteEventHandler(dgvLD_DataBindingComplete);
                BindData();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, " From Journal Tab, Output settings Error");
                LogHelper.WriteLog(typeof(OutputContainer), ex.Message + " From Journal Tab, Output settings Error");
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvLD_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 1)
            {
                DataTable dt = ft.GetReportCriteriaByRef(SessionInfo.UserInfo.File_ftid, dgvLD.Rows[e.RowIndex].Cells[e.ColumnIndex].FormattedValue.ToString());
                if (dt.Rows.Count == 0)
                    dgvLD.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "This save reference doesn't exist in Save Options!";
                else
                    dgvLD.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "";
            }
            if (e.ColumnIndex == 0)
            {
                if ((dgvLD.Rows[e.RowIndex].Cells[0].Value.ToString().Trim() != "") && ((dgvLD.Rows[e.RowIndex].Cells[1].Value.ToString().Trim() == "" && dgvLD.Rows[e.RowIndex].Cells[2].Value.ToString().Trim() == "" && dgvLD.Rows[e.RowIndex].Cells[3].Value.ToString().Trim() == "" && dgvLD.Rows[e.RowIndex].Cells[4].Value.ToString().Trim() == "" && dgvLD.Rows[e.RowIndex].Cells[5].Value.ToString().Trim() == "" && dgvLD.Rows[e.RowIndex].Cells[6].Value.ToString().Trim() == "" && dgvLD.Rows[e.RowIndex].Cells[7].Value.ToString().Trim() == "" && dgvLD.Rows[e.RowIndex].Cells[8].Value.ToString().Trim() == "") || (e.RowIndex > bindCount - 1)))
                    for (int i = 0; i < dgvLD.Rows.Count; i++)
                        if ((dgvLD.Rows[i].Cells[0].Value != null) && (dgvLD.Rows[i].Cells[0].Value.ToString().Trim() == dgvLD.Rows[e.RowIndex].Cells[0].Value.ToString().Trim()))
                        {
                            dgvLD.Rows[e.RowIndex].Cells[1].Value = dgvLD.Rows[i].Cells[1].Value.ToString();
                            dgvLD.Rows[e.RowIndex].Cells[2].Value = dgvLD.Rows[i].Cells[2].Value.ToString();
                            dgvLD.Rows[e.RowIndex].Cells[3].Value = dgvLD.Rows[i].Cells[3].Value.ToString();
                            dgvLD.Rows[e.RowIndex].Cells[4].Value = dgvLD.Rows[i].Cells[4].Value.ToString();
                            dgvLD.Rows[e.RowIndex].Cells[5].Value = dgvLD.Rows[i].Cells[5].Value.ToString();
                            dgvLD.Rows[e.RowIndex].Cells[6].Value = dgvLD.Rows[i].Cells[6].Value.ToString();
                            dgvLD.Rows[e.RowIndex].Cells[7].Value = dgvLD.Rows[i].Cells[7].Value.ToString();
                            dgvLD.Rows[e.RowIndex].Cells[8].Value = dgvLD.Rows[i].Cells[8].Value.ToString();
                            dgvLD.Rows[e.RowIndex].Cells[9].Value = dgvLD.Rows[i].Cells[9].Value.ToString();
                            dgvLD.Rows[e.RowIndex].Cells[10].Value = dgvLD.Rows[i].Cells[10].Value.ToString();
                            break;
                        }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvLD_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex == 0)
                dgvLD.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = dgvLD.RowHeadersDefaultCellStyle.BackColor;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvLD_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            try
            {
                if (isFromBindingFlag)//set consolidation cell yellow in each row
                {
                    foreach (KeyValuePair<int, string> a in NameValueCollection)
                    {
                        if (!string.IsNullOrEmpty(a.Value) && (a.Key >= 0))
                            this.dgvLD.Rows[a.Key].Cells[a.Value].Style.BackColor = Color.Yellow;
                    }
                    foreach (KeyValuePair<int, string> a in NameValueUpdate)
                    {
                        if (!string.IsNullOrEmpty(a.Value) && (a.Key >= 0))
                            this.dgvLD.Rows[a.Key].Cells[a.Value].Style.BackColor = Color.Aqua;
                    }
                }
                for (int i = 0; i < dgvLD.Rows.Count; i++)
                    if (dgvLD.Rows[i].Cells[0].Value != null && i < bindCount)
                    {
                        DataTable dt = ft.GetTemplateActionByRef(SessionInfo.UserInfo.File_ftid, dgvLD.Rows[i].Cells[0].Value.ToString().Trim());
                        if (dt.Rows.Count > 0)
                        {
                            dgvLD.Rows[i].Cells[0].ToolTipText = "Can't be changed! Some Action(s) are using this process reference.";
                        }
                    }
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        private int currentColumnIndex = 0;
        /// <summary>
        /// 
        /// </summary>
        private int currentRowIndex = 0;
        /// <summary>
        /// SHOW CONSOLIDATION/UNSOLIDATION MENUSTRIP WHen right click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvLD_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                if (e.Button == MouseButtons.Right)
                {
                    currentColumnIndex = e.ColumnIndex;
                    currentRowIndex = e.RowIndex;
                    if (e.ColumnIndex >= 11)
                    {
                        setMenu(false);
                        if (dgvLD.Columns[e.ColumnIndex].Selected == false)
                        {
                            dgvLD.ClearSelection();
                            dgvLD.Columns[e.ColumnIndex].Selected = true;
                            if (DataConversionTools.IsPropertyInClassProperties<Common2.Actions>(dgvLD.Columns[currentColumnIndex].Name))
                                updateToolStripMenuItem.Visible = true;
                            else
                                updateToolStripMenuItem.Visible = false;
                        }
                        this.contextMenuStrip1.Show(MousePosition.X, MousePosition.Y);
                    }
                    else if (e.ColumnIndex == 0)
                    {
                        setMenu(true);
                        if (copyrow == -2) PasteStripMenuItem.Enabled = false;
                        else PasteStripMenuItem.Enabled = true;
                        this.contextMenuStrip1.Show(MousePosition.X, MousePosition.Y);
                    }
                }
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="b"></param>
        private void setMenu(bool b)
        {
            if (b)
            {
                toolStripMenuItem1.Visible = false;
                updateToolStripMenuItem.Visible = false;
                CopyStripMenuItem.Visible = true;
                PasteStripMenuItem.Visible = true;
                InsertStripMenuItem.Visible = true;
                RemoveStripMenuItem.Visible = true;
            }
            else
            {
                toolStripMenuItem1.Visible = true;
                updateToolStripMenuItem.Visible = true;
                CopyStripMenuItem.Visible = false;
                PasteStripMenuItem.Visible = false;
                InsertStripMenuItem.Visible = false;
                RemoveStripMenuItem.Visible = false;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            isFromBindingFlag = false;
            if (currentColumnIndex >= 0 && currentRowIndex >= 0 && (dgvLD.Rows[currentRowIndex].Cells[currentColumnIndex].Style.BackColor != Color.Yellow))
            {
                int consCount = 0;
                int updateCount = 0;
                for (int j = 0; j < dgvLD.Columns.Count; j++)
                {
                    if (dgvLD.Rows[currentRowIndex].Cells[j].Style.BackColor == Color.Yellow)
                    {
                        consCount++;
                        if (consCount == 4)
                        {
                            MessageBox.Show("Consolidate cells can't exceed 4 cells per row ! - Data error in journal tab, output settings !", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            LogHelper.WriteLog(typeof(OutputContainer), "Consolidate cells can't exceed 4 cells per row ! - Data error in journal tab, output settings !");
                            return;
                        }
                    }
                    if (dgvLD.Rows[currentRowIndex].Cells[j].Style.BackColor == Color.Aqua)
                    {
                        updateCount++;
                        if (updateCount == 1)
                        {
                            MessageBox.Show("Update cells and consolidate cells should in different rows ! - Data error in journal tab, output settings !", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            LogHelper.WriteLog(typeof(OutputContainer), "Update cells and consolidate cells should in different rows ! - Data error in journal tab, output settings !");
                            return;
                        }
                    }
                }
                dgvLD.Rows[currentRowIndex].Cells[currentColumnIndex].Style.BackColor = Color.Yellow;
                NameValueCollection.Add(new KeyValuePair<int, string>(currentRowIndex, dgvLD.Columns[currentColumnIndex].Name));
            }
            else if (currentColumnIndex >= 0 && currentRowIndex >= 0 && (dgvLD.Rows[currentRowIndex].Cells[currentColumnIndex].Style.BackColor == Color.Yellow))
            {
                dgvLD.Rows[currentRowIndex].Cells[currentColumnIndex].Style.BackColor = Color.Empty;
                NameValueCollection.Remove(new KeyValuePair<int, string>(currentRowIndex, dgvLD.Columns[currentColumnIndex].Name));
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void updateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            isFromBindingFlag = false;
            if (currentColumnIndex >= 0 && currentRowIndex >= 0 && (dgvLD.Rows[currentRowIndex].Cells[currentColumnIndex].Style.BackColor != Color.Aqua))
            {
                int consCount = 0;
                for (int j = 0; j < dgvLD.Columns.Count; j++)
                {
                    if (dgvLD.Rows[currentRowIndex].Cells[j].Style.BackColor == Color.Yellow)
                    {
                        consCount++;
                        if (consCount == 1)
                        {
                            MessageBox.Show("Update cells and consolidate cells should in different rows ! - Data error in journal tab, output settings !", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            LogHelper.WriteLog(typeof(OutputContainer), "Update cells and consolidate cells should in different rows ! - Data error in journal tab, output settings !");
                            return;
                        }
                    }
                }
                dgvLD.Rows[currentRowIndex].Cells[currentColumnIndex].Style.BackColor = Color.Aqua;
                NameValueUpdate.Add(new KeyValuePair<int, string>(currentRowIndex, dgvLD.Columns[currentColumnIndex].Name));
            }
            else if (currentColumnIndex >= 0 && currentRowIndex >= 0 && (dgvLD.Rows[currentRowIndex].Cells[currentColumnIndex].Style.BackColor == Color.Aqua))
            {
                dgvLD.Rows[currentRowIndex].Cells[currentColumnIndex].Style.BackColor = Color.Empty;
                NameValueUpdate.Remove(new KeyValuePair<int, string>(currentRowIndex, dgvLD.Columns[currentColumnIndex].Name));
            }
        }
        int copyrow = -2;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CopyStripMenuItem_Click(object sender, EventArgs e)
        {
            copyrow = currentRowIndex;
        }
        /// <summary>
        /// 
        /// </summary>
        private void reduceBinding()
        {
            dgvLD.DataBindingComplete -= new DataGridViewBindingCompleteEventHandler(dgvLD_DataBindingComplete);
            dgvLD.CellValueChanged -= new DataGridViewCellEventHandler(dgvLD_CellValueChanged);
            dgvLD.RowPostPaint -= new DataGridViewRowPostPaintEventHandler(dgv_RowPostPaint);
            dgvLD.CellFormatting -= new DataGridViewCellFormattingEventHandler(dgvLD_CellFormatting);
        }
        /// <summary>
        /// 
        /// </summary>
        private void plusBinding()
        {
            dgvLD.DataBindingComplete += new DataGridViewBindingCompleteEventHandler(dgvLD_DataBindingComplete);
            dgvLD.CellValueChanged += new DataGridViewCellEventHandler(dgvLD_CellValueChanged);
            dgvLD.RowPostPaint += new DataGridViewRowPostPaintEventHandler(dgv_RowPostPaint);
            dgvLD.CellFormatting += new DataGridViewCellFormattingEventHandler(dgvLD_CellFormatting);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PasteStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                reduceBinding();
                for (int i = 0; i < dgvLD.Columns.Count; i++)
                    if (dgvLD[i, currentRowIndex].Visible)
                        dgvLD[i, currentRowIndex].Value = dgvLD[i, copyrow].Value;
            }
            catch { }
            finally
            {
                plusBinding();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void InsertStripMenuItem_Click(object sender, EventArgs e)
        {
            reduceBinding();
            DataTable dt = (DataTable)this.dgvLD.DataSource;
            DataRow dr = dt.NewRow();
            dt.Rows.InsertAt(dr, currentRowIndex);
            this.dgvLD.DataSource = dt;
            bindCount = dt.Rows.Count;
            plusBinding();
            ChangeKeyValuePair(currentRowIndex, 1);
            if (copyrow != -2 && copyrow >= currentRowIndex) copyrow++;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="currentRowIndex"></param>
        /// <param name="number"></param>
        private void ChangeKeyValuePair(int currentRowIndex, int number)
        {
            List<KeyValuePair<int, string>> templist = new List<KeyValuePair<int, string>>();
            foreach (KeyValuePair<int, string> a in NameValueCollection)
            {
                if ((a.Key > currentRowIndex) || ((a.Key == currentRowIndex) && number > 0))
                {
                    if (!string.IsNullOrEmpty(a.Value))
                    {
                        this.dgvLD.Rows[a.Key].Cells[a.Value].Style.BackColor = Color.Empty;
                        int key = a.Key + number;
                        string value = a.Value;
                        templist.Add(new KeyValuePair<int, string>(key, value));
                    }
                }
                else if ((a.Key == currentRowIndex) && number < 0)
                { }
                else
                    templist.Add(a);
            }
            NameValueCollection = templist;
            templist = new List<KeyValuePair<int, string>>();
            foreach (KeyValuePair<int, string> a in NameValueUpdate)
            {
                if ((a.Key > currentRowIndex) || ((a.Key == currentRowIndex) && number > 0))
                {
                    if (!string.IsNullOrEmpty(a.Value))
                    {
                        this.dgvLD.Rows[a.Key].Cells[a.Value].Style.BackColor = Color.Empty;
                        int key = a.Key + number;
                        string value = a.Value;
                        templist.Add(new KeyValuePair<int, string>(key, value));
                    }
                }
                else if ((a.Key == currentRowIndex) && number < 0)
                { }
                else
                    templist.Add(a);
            }
            NameValueUpdate = templist;
            isFromBindingFlag = true;
            dgvLD_DataBindingComplete(null, null);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RemoveStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                reduceBinding();
                DataTable dt = (DataTable)this.dgvLD.DataSource;
                dt.Rows.RemoveAt(currentRowIndex);
                this.dgvLD.DataSource = dt;
                bindCount = dt.Rows.Count;
                plusBinding();
                ChangeKeyValuePair(currentRowIndex, -1);
                if (copyrow != -2 && copyrow > currentRowIndex)
                    copyrow--;
                else if (copyrow == currentRowIndex)
                    copyrow = -2;
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void dgvLD_KeyDown(object sender, KeyEventArgs e)
        {
            ft.EditingControlWantsInputKey(e.KeyCode, dgvLD);
        }
        /// <summary>
        /// Generate datarow number before the Rows
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgv_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            System.Drawing.Rectangle rectangle = new System.Drawing.Rectangle(e.RowBounds.Location.X,
                e.RowBounds.Location.Y,
                dgvLD.RowHeadersWidth - 4,
                e.RowBounds.Height);
            TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(),
                dgvLD.RowHeadersDefaultCellStyle.Font,
                rectangle,
                dgvLD.RowHeadersDefaultCellStyle.ForeColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgv_CellMouseClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 0 || (e.ColumnIndex == 2) || (e.ColumnIndex == 3) || (e.ColumnIndex == 4) || (e.ColumnIndex == 5))
                    return;
                if (dgvLD[0, e.RowIndex].Value == "" || dgvLD[0, e.RowIndex].Value == null || dgvLD[0, e.RowIndex].Value.ToString() == "") dgvLD[0, e.RowIndex].Value = e.RowIndex + 1;
                dgvLD[e.ColumnIndex, e.RowIndex].Value = "";
                var xlRange = Globals.ThisAddIn.Application.ActiveCell.Address;
                if (xlRange != null)
                {
                    dgvLD[e.ColumnIndex, e.RowIndex].Value = xlRange.Replace("$", "");
                    dgvLD.Focus();
                }
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgv_CellMouseDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0 || (e.ColumnIndex == 2) || (e.ColumnIndex == 3) || (e.ColumnIndex == 4) || (e.ColumnIndex == 5))
                return;
            try
            {
                if (dgvLD[e.ColumnIndex, e.RowIndex].Value != null)
                {
                    if (dgvLD[e.ColumnIndex, e.RowIndex].Value.ToString().Contains("$"))
                    {
                        string KeyWord = dgvLD[e.ColumnIndex, e.RowIndex].Value.ToString().Replace("$", "");
                        dgvLD[e.ColumnIndex, e.RowIndex].Value = "";
                        dgvLD[e.ColumnIndex, e.RowIndex].Value = KeyWord;
                    }
                    else
                    {
                        string KeyWord = dgvLD[e.ColumnIndex, e.RowIndex].Value.ToString();
                        string res = Regex.Replace(KeyWord, @"(\d+)|(\s+) ", " $1 $2 ", RegexOptions.Compiled | RegexOptions.IgnoreCase);
                        KeyWord = "$" + res.Trim().Replace(" ", "$");
                        dgvLD[e.ColumnIndex, e.RowIndex].Value = "";
                        dgvLD[e.ColumnIndex, e.RowIndex].Value = KeyWord;
                    }
                    dgvLD.EndEdit();
                }
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        private void BindData()
        {
            if (NameValueCollection == null) NameValueCollection = new List<KeyValuePair<int, string>>();
            if (NameValueUpdate == null) NameValueUpdate = new List<KeyValuePair<int, string>>();
            DataTable refdt = ft.GetAllJournalRefOfTemplate();
            DataTable dtfinal = DataConversionTools.ConvertToDataTableStructure<rsTemplateJournal>();
            foreach (DataRow dr in refdt.Rows)
            {
                string refnum = dr["references"].ToString().Trim();
                DataTable dt = ft.GetLineDetailDataFromDB("0", refnum);
                foreach (DataRow dr2 in dt.Rows)
                {
                    DataRow drnew = dtfinal.NewRow();
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        drnew[dt.Columns[j].ColumnName] = dr2[dt.Columns[j].ColumnName];//asign dt2's data to dt
                    }
                    dtfinal.Rows.Add(drnew);
                }
                BindConsolidationDGV(ref dtfinal, ref NameValueCollection, refnum, dtfinal.Rows.Count);
                BindUpdateDGV(ref dtfinal, ref NameValueUpdate, refnum, dtfinal.Rows.Count);
            }
            isFromBindingFlag = true;
            this.dgvLD.DataSource = dtfinal;
            bindCount = dtfinal.Rows.Count;
            this.panel4.AutoSize = true;
            this.panel5.Controls.Add(dgvLD);
            for (int i = 0; i < dgvLD.Columns.Count; i++) { dgvLD.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable; if (i < 8) dgvLD.Columns[i].DefaultCellStyle.BackColor = System.Drawing.Color.Ivory; }
        }
        private int bindCount = 0;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Save(sender, null);
                SaveCons(sender, null);
                SaveTransUpd(sender, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            Ribbon2._MyOutputCustomTaskPane.Visible = false;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="ws"></param>
        public void Save(object sender, Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            List<string> LineIndicatorList = new List<string>();
            List<string> startInCellList = new List<string>();
            SessionInfo.UserInfo.AllowBalTran = "";
            SessionInfo.UserInfo.AllowPostToSuspended = "";
            SessionInfo.UserInfo.PostProvisional = "";
            if (!string.IsNullOrEmpty(SessionInfo.UserInfo.File_ftid))
            {
                SqlConnection conn = null;
                SqlDataReader rdr = null;
                try
                {
                    conn = new
                        SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                    conn.Open();
                    SqlCommand cmd = new SqlCommand("rsTemplateJournal_Del", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                    cmd.Parameters.Add(new SqlParameter("@Type", "0"));
                    if (sender != null)
                    {
                        rdr = cmd.ExecuteReader();
                        rdr.Close();
                    }
                    finallist.Clear(); //Initiate data
                    SqlCommand cmd2 = new SqlCommand("rsTemplateJournal_Ins", conn);
                    cmd2.CommandType = CommandType.StoredProcedure;
                    for (int i = 0; i < this.dgvLD.Rows.Count; i++)
                    {
                        bool iscontinue = false;
                        for (int j = 0; j < dgvLD.Columns.Count; j++)
                        {
                            if (dgvLD.Rows[i].Cells[j].Style.BackColor == Color.Yellow || dgvLD.Rows[i].Cells[j].Style.BackColor == Color.Aqua)
                            {
                                iscontinue = true;
                            }
                        }
                        if (iscontinue)
                            continue;

                        string startCell = string.Empty;
                        string journalN = string.Empty;
                        string journalLN = string.Empty;
                        string abt = string.Empty;
                        string apsa = string.Empty;
                        string pp = string.Empty;

                        Specialist re = new Specialist();
                        re.Reference = this.dgvLD.Rows[i].Cells[0].Value == null ? "" : this.dgvLD.Rows[i].Cells[0].Value.ToString().Replace(" ", "");
                        re.SaveReference = this.dgvLD.Rows[i].Cells[1].Value == null ? "" : this.dgvLD.Rows[i].Cells[1].Value.ToString();
                        re.BalanceBy = this.dgvLD.Rows[i].Cells[2].Value == null ? "" : this.dgvLD.Rows[i].Cells[2].Value.ToString();
                        abt = this.dgvLD.Rows[i].Cells[3].Value == null ? "" : this.dgvLD.Rows[i].Cells[3].Value.ToString();
                        apsa = this.dgvLD.Rows[i].Cells[4].Value == null ? "" : this.dgvLD.Rows[i].Cells[4].Value.ToString();
                        pp = this.dgvLD.Rows[i].Cells[5].Value == null ? "" : this.dgvLD.Rows[i].Cells[5].Value.ToString();
                        re.AllowBalTrans = abt;
                        re.AllowPostSuspAcco = apsa;
                        re.PostProvisional = pp;
                        re.LineIndicator = this.dgvLD.Rows[i].Cells[6].Value == null ? "" : this.dgvLD.Rows[i].Cells[6].Value.ToString();
                        startCell = this.dgvLD.Rows[i].Cells[7].Value == null ? "" : this.dgvLD.Rows[i].Cells[7].Value.ToString();
                        re.StartInCell = startCell;
                        re.populatecellwithJN = this.dgvLD.Rows[i].Cells[8].Value == null ? "" : this.dgvLD.Rows[i].Cells[8].Value.ToString();
                        journalN = this.dgvLD.Rows[i].Cells[9].Value == null ? "" : this.dgvLD.Rows[i].Cells[9].Value.ToString();
                        journalLN = this.dgvLD.Rows[i].Cells[10].Value == null ? "" : this.dgvLD.Rows[i].Cells[10].Value.ToString();

                        re.Ledger = this.dgvLD.Rows[i].Cells[11].Value == null ? "" : this.dgvLD.Rows[i].Cells[11].Value.ToString();
                        re.AccountCode = this.dgvLD.Rows[i].Cells[12].Value == null ? "" : this.dgvLD.Rows[i].Cells[12].Value.ToString();
                        re.AccountingPeriod = this.dgvLD.Rows[i].Cells[13].Value == null ? "" : this.dgvLD.Rows[i].Cells[13].Value.ToString();
                        re.TransactionDate = this.dgvLD.Rows[i].Cells[14].Value == null ? "" : this.dgvLD.Rows[i].Cells[14].Value.ToString();
                        re.DueDate = this.dgvLD.Rows[i].Cells[15].Value == null ? "" : this.dgvLD.Rows[i].Cells[15].Value.ToString();
                        re.JournalType = this.dgvLD.Rows[i].Cells[16].Value == null ? "" : this.dgvLD.Rows[i].Cells[16].Value.ToString();
                        re.JournalSource = this.dgvLD.Rows[i].Cells[17].Value == null ? "" : this.dgvLD.Rows[i].Cells[17].Value.ToString();
                        re.TransactionReference = this.dgvLD.Rows[i].Cells[18].Value == null ? "" : this.dgvLD.Rows[i].Cells[18].Value.ToString();
                        re.Description = this.dgvLD.Rows[i].Cells[19].Value == null ? "" : this.dgvLD.Rows[i].Cells[19].Value.ToString();
                        re.AllocationMarker = this.dgvLD.Rows[i].Cells[20].Value == null ? "" : this.dgvLD.Rows[i].Cells[20].Value.ToString();
                        re.AnalysisCode1 = this.dgvLD.Rows[i].Cells[21].Value == null ? "" : this.dgvLD.Rows[i].Cells[21].Value.ToString();
                        re.AnalysisCode2 = this.dgvLD.Rows[i].Cells[22].Value == null ? "" : this.dgvLD.Rows[i].Cells[22].Value.ToString();
                        re.AnalysisCode3 = this.dgvLD.Rows[i].Cells[23].Value == null ? "" : this.dgvLD.Rows[i].Cells[23].Value.ToString();
                        re.AnalysisCode4 = this.dgvLD.Rows[i].Cells[24].Value == null ? "" : this.dgvLD.Rows[i].Cells[24].Value.ToString();
                        re.AnalysisCode5 = this.dgvLD.Rows[i].Cells[25].Value == null ? "" : this.dgvLD.Rows[i].Cells[25].Value.ToString();
                        re.AnalysisCode6 = this.dgvLD.Rows[i].Cells[26].Value == null ? "" : this.dgvLD.Rows[i].Cells[26].Value.ToString();
                        re.AnalysisCode7 = this.dgvLD.Rows[i].Cells[27].Value == null ? "" : this.dgvLD.Rows[i].Cells[27].Value.ToString();
                        re.AnalysisCode8 = this.dgvLD.Rows[i].Cells[28].Value == null ? "" : this.dgvLD.Rows[i].Cells[28].Value.ToString();
                        re.AnalysisCode9 = this.dgvLD.Rows[i].Cells[29].Value == null ? "" : this.dgvLD.Rows[i].Cells[29].Value.ToString();
                        re.AnalysisCode10 = this.dgvLD.Rows[i].Cells[30].Value == null ? "" : this.dgvLD.Rows[i].Cells[30].Value.ToString();
                        re.GenDesc1 = this.dgvLD.Rows[i].Cells[31].Value == null ? "" : this.dgvLD.Rows[i].Cells[31].Value.ToString();
                        re.GenDesc2 = this.dgvLD.Rows[i].Cells[32].Value == null ? "" : this.dgvLD.Rows[i].Cells[32].Value.ToString();
                        re.GenDesc3 = this.dgvLD.Rows[i].Cells[33].Value == null ? "" : this.dgvLD.Rows[i].Cells[33].Value.ToString();
                        re.GenDesc4 = this.dgvLD.Rows[i].Cells[34].Value == null ? "" : this.dgvLD.Rows[i].Cells[34].Value.ToString();
                        re.GenDesc5 = this.dgvLD.Rows[i].Cells[35].Value == null ? "" : this.dgvLD.Rows[i].Cells[35].Value.ToString();
                        re.GenDesc6 = this.dgvLD.Rows[i].Cells[36].Value == null ? "" : this.dgvLD.Rows[i].Cells[36].Value.ToString();
                        re.GenDesc7 = this.dgvLD.Rows[i].Cells[37].Value == null ? "" : this.dgvLD.Rows[i].Cells[37].Value.ToString();
                        re.GenDesc8 = this.dgvLD.Rows[i].Cells[38].Value == null ? "" : this.dgvLD.Rows[i].Cells[38].Value.ToString();
                        re.GenDesc9 = this.dgvLD.Rows[i].Cells[39].Value == null ? "" : this.dgvLD.Rows[i].Cells[39].Value.ToString();
                        re.GenDesc10 = this.dgvLD.Rows[i].Cells[40].Value == null ? "" : this.dgvLD.Rows[i].Cells[40].Value.ToString();
                        re.GenDesc11 = this.dgvLD.Rows[i].Cells[41].Value == null ? "" : this.dgvLD.Rows[i].Cells[41].Value.ToString();
                        re.GenDesc12 = this.dgvLD.Rows[i].Cells[42].Value == null ? "" : this.dgvLD.Rows[i].Cells[42].Value.ToString();
                        re.GenDesc13 = this.dgvLD.Rows[i].Cells[43].Value == null ? "" : this.dgvLD.Rows[i].Cells[43].Value.ToString();
                        re.GenDesc14 = this.dgvLD.Rows[i].Cells[44].Value == null ? "" : this.dgvLD.Rows[i].Cells[44].Value.ToString();
                        re.GenDesc15 = this.dgvLD.Rows[i].Cells[45].Value == null ? "" : this.dgvLD.Rows[i].Cells[45].Value.ToString();
                        re.GenDesc16 = this.dgvLD.Rows[i].Cells[46].Value == null ? "" : this.dgvLD.Rows[i].Cells[46].Value.ToString();
                        re.GenDesc17 = this.dgvLD.Rows[i].Cells[47].Value == null ? "" : this.dgvLD.Rows[i].Cells[47].Value.ToString();
                        re.GenDesc18 = this.dgvLD.Rows[i].Cells[48].Value == null ? "" : this.dgvLD.Rows[i].Cells[48].Value.ToString();
                        re.GenDesc19 = this.dgvLD.Rows[i].Cells[49].Value == null ? "" : this.dgvLD.Rows[i].Cells[49].Value.ToString();
                        re.GenDesc20 = this.dgvLD.Rows[i].Cells[50].Value == null ? "" : this.dgvLD.Rows[i].Cells[50].Value.ToString();
                        re.GenDesc21 = this.dgvLD.Rows[i].Cells[51].Value == null ? "" : this.dgvLD.Rows[i].Cells[51].Value.ToString();
                        re.GenDesc22 = this.dgvLD.Rows[i].Cells[52].Value == null ? "" : this.dgvLD.Rows[i].Cells[52].Value.ToString();
                        re.GenDesc23 = this.dgvLD.Rows[i].Cells[53].Value == null ? "" : this.dgvLD.Rows[i].Cells[53].Value.ToString();
                        re.GenDesc24 = this.dgvLD.Rows[i].Cells[54].Value == null ? "" : this.dgvLD.Rows[i].Cells[54].Value.ToString();
                        re.GenDesc25 = this.dgvLD.Rows[i].Cells[55].Value == null ? "" : this.dgvLD.Rows[i].Cells[55].Value.ToString();
                        re.TransactionAmount = this.dgvLD.Rows[i].Cells[56].Value == null ? "" : this.dgvLD.Rows[i].Cells[56].Value.ToString();
                        re.CurrencyCode = this.dgvLD.Rows[i].Cells[57].Value == null ? "" : this.dgvLD.Rows[i].Cells[57].Value.ToString();
                        re.DebitCredit = "";
                        re.BaseAmount = this.dgvLD.Rows[i].Cells[58].Value == null ? "" : this.dgvLD.Rows[i].Cells[58].Value.ToString();
                        re.Base2ReportingAmount = this.dgvLD.Rows[i].Cells[59].Value == null ? "" : this.dgvLD.Rows[i].Cells[59].Value.ToString();
                        re.Value4Amount = this.dgvLD.Rows[i].Cells[60].Value == null ? "" : this.dgvLD.Rows[i].Cells[60].Value.ToString();
                        if (string.IsNullOrEmpty(re.ToString())) continue;
                        if (string.IsNullOrEmpty(re.SaveReference))
                            this.dgvLD.Rows[i].Cells[1].ErrorText = "nullable, but unexpected result would happen when save transaction.";
                        else
                            this.dgvLD.Rows[i].Cells[1].ErrorText = string.Empty;

                        if (string.IsNullOrEmpty(re.LineIndicator))
                            this.dgvLD.Rows[i].Cells[6].ErrorText = "Not null.";
                        else
                            this.dgvLD.Rows[i].Cells[6].ErrorText = string.Empty;

                        cmd2.Parameters.Add(new SqlParameter("@Ledger", re.Ledger));
                        cmd2.Parameters.Add(new SqlParameter("@ft_Account", re.AccountCode));
                        cmd2.Parameters.Add(new SqlParameter("@Period", re.AccountingPeriod));
                        cmd2.Parameters.Add(new SqlParameter("@TransDate", re.TransactionDate));
                        cmd2.Parameters.Add(new SqlParameter("@DueDate", re.DueDate));
                        cmd2.Parameters.Add(new SqlParameter("@JrnlType", re.JournalType));
                        cmd2.Parameters.Add(new SqlParameter("@JrnlSource", re.JournalSource));
                        cmd2.Parameters.Add(new SqlParameter("@TransRef", re.TransactionReference));
                        cmd2.Parameters.Add(new SqlParameter("@Description", re.Description));
                        cmd2.Parameters.Add(new SqlParameter("@AlloctnMarker", re.AllocationMarker));
                        cmd2.Parameters.Add(new SqlParameter("@LA1", re.AnalysisCode1));
                        cmd2.Parameters.Add(new SqlParameter("@LA2", re.AnalysisCode2));
                        cmd2.Parameters.Add(new SqlParameter("@LA3", re.AnalysisCode3));
                        cmd2.Parameters.Add(new SqlParameter("@LA4", re.AnalysisCode4));
                        cmd2.Parameters.Add(new SqlParameter("@LA5", re.AnalysisCode5));
                        cmd2.Parameters.Add(new SqlParameter("@LA6", re.AnalysisCode6));
                        cmd2.Parameters.Add(new SqlParameter("@LA7", re.AnalysisCode7));
                        cmd2.Parameters.Add(new SqlParameter("@LA8", re.AnalysisCode8));
                        cmd2.Parameters.Add(new SqlParameter("@LA9", re.AnalysisCode9));
                        cmd2.Parameters.Add(new SqlParameter("@LA10", re.AnalysisCode10));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc1", re.GenDesc1));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc2", re.GenDesc2));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc3", re.GenDesc3));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc4", re.GenDesc4));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc5", re.GenDesc5));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc6", re.GenDesc6));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc7", re.GenDesc7));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc8", re.GenDesc8));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc9", re.GenDesc9));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc10", re.GenDesc10));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc11", re.GenDesc11));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc12", re.GenDesc12));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc13", re.GenDesc13));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc14", re.GenDesc14));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc15", re.GenDesc15));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc16", re.GenDesc16));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc17", re.GenDesc17));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc18", re.GenDesc18));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc19", re.GenDesc19));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc20", re.GenDesc20));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc21", re.GenDesc21));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc22", re.GenDesc22));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc23", re.GenDesc23));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc24", re.GenDesc24));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc25", re.GenDesc25));
                        cmd2.Parameters.Add(new SqlParameter("@TransAmount", re.TransactionAmount));
                        cmd2.Parameters.Add(new SqlParameter("@Currency", re.CurrencyCode));
                        cmd2.Parameters.Add(new SqlParameter("@BaseAmount", re.BaseAmount));
                        cmd2.Parameters.Add(new SqlParameter("@2ndBase", re.Base2ReportingAmount));
                        cmd2.Parameters.Add(new SqlParameter("@4thAmount", re.Value4Amount));
                        cmd2.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                        cmd2.Parameters.Add(new SqlParameter("@LineIndicator", re.LineIndicator));
                        cmd2.Parameters.Add(new SqlParameter("@StartinginCell", startCell));
                        cmd2.Parameters.Add(new SqlParameter("@BalanceBy", re.BalanceBy));
                        cmd2.Parameters.Add(new SqlParameter("@PopWithJNNumber", re.populatecellwithJN));
                        cmd2.Parameters.Add(new SqlParameter("@Reference", re.Reference));
                        cmd2.Parameters.Add(new SqlParameter("@SaveReference", re.SaveReference));

                        cmd2.Parameters.Add(new SqlParameter("@JournalNumber", journalN));
                        cmd2.Parameters.Add(new SqlParameter("@JournalLineNumber", journalLN));
                        cmd2.Parameters.Add(new SqlParameter("@InputFields", ""));
                        cmd2.Parameters.Add(new SqlParameter("@UpdateFields", ""));
                        cmd2.Parameters.Add(new SqlParameter("@ConsolidateBy1", ""));
                        cmd2.Parameters.Add(new SqlParameter("@ConsolidateBy2", ""));
                        cmd2.Parameters.Add(new SqlParameter("@ConsolidateBy3", ""));
                        cmd2.Parameters.Add(new SqlParameter("@ConsolidateBy4", ""));
                        cmd2.Parameters.Add(new SqlParameter("@Type", "0"));//0.post 1.consolidation 2.update
                        cmd2.Parameters.Add(new SqlParameter("@AllowBalTrans", abt));
                        cmd2.Parameters.Add(new SqlParameter("@AllowPostSuspAcco", apsa));
                        cmd2.Parameters.Add(new SqlParameter("@PostProvisional", pp));
                        if (sender != null)
                        {
                            rdr = cmd2.ExecuteReader();
                            rdr.Close();
                        }
                        cmd2.Parameters.Clear();

                        if (!string.IsNullOrEmpty(re.LineIndicator) && !string.IsNullOrEmpty(startCell))
                        {
                            finallist.Add(re);
                            LineIndicatorList.Add(re.LineIndicator);
                            startInCellList.Add(startCell);
                        }
                    }
                    this.Invalidate();
                    if (sender == null)
                        AddLineDetailEntityListIntoFinalList(finallist, LineIndicatorList, startInCellList, ws, ref finallist);
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message + " - Data Error in Journal tab, Output settings !");
                }
                finally
                {
                    if (conn != null)
                    {
                        conn.Close();
                    }
                    if (rdr != null)
                    {
                        rdr.Close();
                    }
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="list"></param>
        /// <param name="LineIndicatorList"></param>
        /// <param name="StartingInCell"></param>
        /// <param name="ws"></param>
        /// <param name="final"></param>
        private void AddLineDetailEntityListIntoFinalListForTransUpd(List<ExcelAddIn4.Common2.Specialist> list, List<string> LineIndicatorList, List<string> StartingInCellList, Microsoft.Office.Interop.Excel.Worksheet ws, ref List<ExcelAddIn4.Common2.Specialist> final)
        {
            List<ExcelAddIn4.Common2.Specialist> tmplist2 = new List<ExcelAddIn4.Common2.Specialist>();
            List<string> usedString = new List<string>();
            Predicate<string> pred = EquaWithName;
            for (int i = 0; i < LineIndicatorList.Count; i++)
            {
                tmpstr = StartingInCellList[i] + "," + LineIndicatorList[i];
                if (!usedString.Exists(pred))
                {
                    List<ExcelAddIn4.Common2.Specialist> tmplist = ft.GetEntityListFromDGVForTransUpd(StartingInCellList[i], LineIndicatorList[i], list.FindAll((ExcelAddIn4.Common2.Specialist p) => { return p.LineIndicator == LineIndicatorList[i] & p.StartInCell == StartingInCellList[i]; }), ws);//ft.GetEntityListFromDGVForTransUpd(StartingInCellList[i], LineIndicatorList[i], list[i], ws);
                    if (tmplist != null)
                        foreach (ExcelAddIn4.Common2.Specialist sp in tmplist)
                        {
                            tmplist2.Add(sp);
                        }
                    usedString.Add(StartingInCellList[i] + "," + LineIndicatorList[i]);
                }
            }
            final = tmplist2;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="elem"></param>
        /// <returns></returns>
        private bool EquaWithName(string elem)
        {
            if (elem == tmpstr)
                return true;
            return false;
        }
        string tmpstr = string.Empty;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="list"></param>
        /// <param name="LineIndicatorList"></param>
        /// <param name="StartingInCell"></param>
        /// <param name="ws"></param>
        /// <param name="final"></param>
        private void AddLineDetailEntityListIntoFinalList(List<Specialist> list, List<string> LineIndicatorList, List<string> StartingInCellList, Microsoft.Office.Interop.Excel.Worksheet ws, ref List<Specialist> final)
        {
            List<Specialist> tmplist2 = new List<Specialist>();
            List<string> usedString = new List<string>();
            Predicate<string> pred = EquaWithName;
            for (int i = 0; i < LineIndicatorList.Count; i++)
            {
                tmpstr = StartingInCellList[i] + "," + LineIndicatorList[i];
                if (!usedString.Exists(pred))
                {
                    List<Specialist> tmplist = ft.GetEntityListFromDGV(StartingInCellList[i], LineIndicatorList[i], list.FindAll((Specialist p) => { return p.LineIndicator == LineIndicatorList[i] & p.StartInCell == StartingInCellList[i]; }), ws);//ft.GetEntityListFromDGV(StartingInCellList[i], LineIndicatorList[i], list[i], ws);
                    if (tmplist != null)
                        foreach (Specialist sp in tmplist)
                        {
                            tmplist2.Add(sp);
                        }
                    usedString.Add(StartingInCellList[i] + "," + LineIndicatorList[i]);
                }
            }
            final = tmplist2;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="list"></param>
        /// <param name="LineIndicatorList"></param>
        /// <param name="StartingInCell"></param>
        /// <param name="ws"></param>
        /// <param name="final"></param>
        private void AddCreateTextFileEntityListIntoFinalList(List<RowCreateTextFile> list, List<string> LineIndicatorList, List<string> StartingInCellList, Microsoft.Office.Interop.Excel.Worksheet ws, ref List<RowCreateTextFile> final)
        {
            List<RowCreateTextFile> tmplist2 = new List<RowCreateTextFile>();
            List<string> usedString = new List<string>();
            Predicate<string> pred = EquaWithName;
            for (int i = 0; i < LineIndicatorList.Count; i++)
            {
                tmpstr = StartingInCellList[i] + "," + LineIndicatorList[i];
                if (!usedString.Exists(pred))
                {
                    List<RowCreateTextFile> tmplist = ft.GetEntityListFromDGVForCreateTextFile(StartingInCellList[i], LineIndicatorList[i], list.FindAll((RowCreateTextFile p) => { return p.LineIndicator == LineIndicatorList[i] & p.StartinginCell == StartingInCellList[i]; }), ws);//ft.GetEntityListFromDGVForCreateTextFile(StartingInCellList[i], LineIndicatorList[i], list[i], ws);
                    if (tmplist != null)
                        foreach (RowCreateTextFile sp in tmplist)
                        {
                            tmplist2.Add(sp);
                        }
                    usedString.Add(StartingInCellList[i] + "," + LineIndicatorList[i]);
                }
            }
            final = tmplist2;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="ws"></param>
        public void SaveCons(object sender, Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            List<string> LineIndicatorList = new List<string>();
            List<string> startInCellList = new List<string>();
            List<string> consStr = new List<string>();
            string str = string.Empty;
            List<int> ConsolidationRowNumber = new List<int>();
            int consCurrentCount = 0;
            for (int i = 0; i < dgvLD.Rows.Count; i++)
            {
                for (int j = 0; j < dgvLD.Columns.Count; j++)
                {
                    if (dgvLD.Rows[i].Cells[j].Style.BackColor == Color.Yellow)
                    {
                        str += dgvLD.Columns[j].Name + ",";
                        ConsolidationRowNumber.Add(i);
                    }
                }
                if (!string.IsNullOrEmpty(str))
                {
                    consStr.Add(str);
                    str = "";
                }
            }
            if (dgvLD.RowCount == 1)
            {
                return;
            }
            if (!string.IsNullOrEmpty(SessionInfo.UserInfo.File_ftid))
            {
                SqlConnection conn = null;
                SqlDataReader rdr = null;
                try
                {
                    conn = new
                        SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());// create and open a connection object  "Server=(local);Database=RSFinanceTools;User ID=sa;Password=as
                    conn.Open();
                    SqlCommand cmd = new SqlCommand("rsTemplateJournal_Del", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                    cmd.Parameters.Add(new SqlParameter("@Type", "1"));
                    if (sender != null)
                    {
                        rdr = cmd.ExecuteReader();
                        rdr.Close();
                    }
                    List<Specialist> ll = new List<Specialist>();//Initiate data //finallist.Clear();
                    SqlCommand cmd2 = new SqlCommand("rsTemplateJournal_Ins", conn);
                    cmd2.CommandType = CommandType.StoredProcedure;
                    for (int i = 0; i < this.dgvLD.Rows.Count; i++)
                    {
                        if (!ConsolidationRowNumber.Contains(i))
                        {
                            continue;
                        }
                        string startCell = string.Empty;
                        string journalN = string.Empty;
                        string journalLN = string.Empty;
                        string abt = string.Empty;
                        string apsa = string.Empty;
                        string pp = string.Empty;
                        Specialist re = new Specialist();
                        re.Reference = this.dgvLD.Rows[i].Cells[0].Value == null ? "" : this.dgvLD.Rows[i].Cells[0].Value.ToString().Replace(" ", "");
                        re.SaveReference = this.dgvLD.Rows[i].Cells[1].Value == null ? "" : this.dgvLD.Rows[i].Cells[1].Value.ToString();
                        re.BalanceBy = this.dgvLD.Rows[i].Cells[2].Value == null ? "" : this.dgvLD.Rows[i].Cells[2].Value.ToString();
                        abt = this.dgvLD.Rows[i].Cells[3].Value == null ? "" : this.dgvLD.Rows[i].Cells[3].Value.ToString();
                        apsa = this.dgvLD.Rows[i].Cells[4].Value == null ? "" : this.dgvLD.Rows[i].Cells[4].Value.ToString();
                        pp = this.dgvLD.Rows[i].Cells[5].Value == null ? "" : this.dgvLD.Rows[i].Cells[5].Value.ToString();
                        re.AllowBalTrans = abt;
                        re.AllowPostSuspAcco = apsa;
                        re.PostProvisional = pp;
                        re.LineIndicator = this.dgvLD.Rows[i].Cells[6].Value == null ? "" : this.dgvLD.Rows[i].Cells[6].Value.ToString();
                        startCell = this.dgvLD.Rows[i].Cells[7].Value == null ? "" : this.dgvLD.Rows[i].Cells[7].Value.ToString();
                        re.populatecellwithJN = this.dgvLD.Rows[i].Cells[8].Value == null ? "" : this.dgvLD.Rows[i].Cells[8].Value.ToString();
                        journalN = this.dgvLD.Rows[i].Cells[9].Value == null ? "" : this.dgvLD.Rows[i].Cells[9].Value.ToString();
                        journalLN = this.dgvLD.Rows[i].Cells[10].Value == null ? "" : this.dgvLD.Rows[i].Cells[10].Value.ToString();

                        re.Ledger = this.dgvLD.Rows[i].Cells[11].Value == null ? "" : this.dgvLD.Rows[i].Cells[11].Value.ToString();
                        re.AccountCode = this.dgvLD.Rows[i].Cells[12].Value == null ? "" : this.dgvLD.Rows[i].Cells[12].Value.ToString();
                        re.AccountingPeriod = this.dgvLD.Rows[i].Cells[13].Value == null ? "" : this.dgvLD.Rows[i].Cells[13].Value.ToString();
                        re.TransactionDate = this.dgvLD.Rows[i].Cells[14].Value == null ? "" : this.dgvLD.Rows[i].Cells[14].Value.ToString();
                        re.DueDate = this.dgvLD.Rows[i].Cells[15].Value == null ? "" : this.dgvLD.Rows[i].Cells[15].Value.ToString();
                        re.JournalType = this.dgvLD.Rows[i].Cells[16].Value == null ? "" : this.dgvLD.Rows[i].Cells[16].Value.ToString();
                        re.JournalSource = this.dgvLD.Rows[i].Cells[17].Value == null ? "" : this.dgvLD.Rows[i].Cells[17].Value.ToString();
                        re.TransactionReference = this.dgvLD.Rows[i].Cells[18].Value == null ? "" : this.dgvLD.Rows[i].Cells[18].Value.ToString();
                        re.Description = this.dgvLD.Rows[i].Cells[19].Value == null ? "" : this.dgvLD.Rows[i].Cells[19].Value.ToString();
                        re.AllocationMarker = this.dgvLD.Rows[i].Cells[20].Value == null ? "" : this.dgvLD.Rows[i].Cells[20].Value.ToString();
                        re.AnalysisCode1 = this.dgvLD.Rows[i].Cells[21].Value == null ? "" : this.dgvLD.Rows[i].Cells[21].Value.ToString();
                        re.AnalysisCode2 = this.dgvLD.Rows[i].Cells[22].Value == null ? "" : this.dgvLD.Rows[i].Cells[22].Value.ToString();
                        re.AnalysisCode3 = this.dgvLD.Rows[i].Cells[23].Value == null ? "" : this.dgvLD.Rows[i].Cells[23].Value.ToString();
                        re.AnalysisCode4 = this.dgvLD.Rows[i].Cells[24].Value == null ? "" : this.dgvLD.Rows[i].Cells[24].Value.ToString();
                        re.AnalysisCode5 = this.dgvLD.Rows[i].Cells[25].Value == null ? "" : this.dgvLD.Rows[i].Cells[25].Value.ToString();
                        re.AnalysisCode6 = this.dgvLD.Rows[i].Cells[26].Value == null ? "" : this.dgvLD.Rows[i].Cells[26].Value.ToString();
                        re.AnalysisCode7 = this.dgvLD.Rows[i].Cells[27].Value == null ? "" : this.dgvLD.Rows[i].Cells[27].Value.ToString();
                        re.AnalysisCode8 = this.dgvLD.Rows[i].Cells[28].Value == null ? "" : this.dgvLD.Rows[i].Cells[28].Value.ToString();
                        re.AnalysisCode9 = this.dgvLD.Rows[i].Cells[29].Value == null ? "" : this.dgvLD.Rows[i].Cells[29].Value.ToString();
                        re.AnalysisCode10 = this.dgvLD.Rows[i].Cells[30].Value == null ? "" : this.dgvLD.Rows[i].Cells[30].Value.ToString();
                        re.GenDesc1 = this.dgvLD.Rows[i].Cells[31].Value == null ? "" : this.dgvLD.Rows[i].Cells[31].Value.ToString();
                        re.GenDesc2 = this.dgvLD.Rows[i].Cells[32].Value == null ? "" : this.dgvLD.Rows[i].Cells[32].Value.ToString();
                        re.GenDesc3 = this.dgvLD.Rows[i].Cells[33].Value == null ? "" : this.dgvLD.Rows[i].Cells[33].Value.ToString();
                        re.GenDesc4 = this.dgvLD.Rows[i].Cells[34].Value == null ? "" : this.dgvLD.Rows[i].Cells[34].Value.ToString();
                        re.GenDesc5 = this.dgvLD.Rows[i].Cells[35].Value == null ? "" : this.dgvLD.Rows[i].Cells[35].Value.ToString();
                        re.GenDesc6 = this.dgvLD.Rows[i].Cells[36].Value == null ? "" : this.dgvLD.Rows[i].Cells[36].Value.ToString();
                        re.GenDesc7 = this.dgvLD.Rows[i].Cells[37].Value == null ? "" : this.dgvLD.Rows[i].Cells[37].Value.ToString();
                        re.GenDesc8 = this.dgvLD.Rows[i].Cells[38].Value == null ? "" : this.dgvLD.Rows[i].Cells[38].Value.ToString();
                        re.GenDesc9 = this.dgvLD.Rows[i].Cells[39].Value == null ? "" : this.dgvLD.Rows[i].Cells[39].Value.ToString();
                        re.GenDesc10 = this.dgvLD.Rows[i].Cells[40].Value == null ? "" : this.dgvLD.Rows[i].Cells[40].Value.ToString();
                        re.GenDesc11 = this.dgvLD.Rows[i].Cells[41].Value == null ? "" : this.dgvLD.Rows[i].Cells[41].Value.ToString();
                        re.GenDesc12 = this.dgvLD.Rows[i].Cells[42].Value == null ? "" : this.dgvLD.Rows[i].Cells[42].Value.ToString();
                        re.GenDesc13 = this.dgvLD.Rows[i].Cells[43].Value == null ? "" : this.dgvLD.Rows[i].Cells[43].Value.ToString();
                        re.GenDesc14 = this.dgvLD.Rows[i].Cells[44].Value == null ? "" : this.dgvLD.Rows[i].Cells[44].Value.ToString();
                        re.GenDesc15 = this.dgvLD.Rows[i].Cells[45].Value == null ? "" : this.dgvLD.Rows[i].Cells[45].Value.ToString();
                        re.GenDesc16 = this.dgvLD.Rows[i].Cells[46].Value == null ? "" : this.dgvLD.Rows[i].Cells[46].Value.ToString();
                        re.GenDesc17 = this.dgvLD.Rows[i].Cells[47].Value == null ? "" : this.dgvLD.Rows[i].Cells[47].Value.ToString();
                        re.GenDesc18 = this.dgvLD.Rows[i].Cells[48].Value == null ? "" : this.dgvLD.Rows[i].Cells[48].Value.ToString();
                        re.GenDesc19 = this.dgvLD.Rows[i].Cells[49].Value == null ? "" : this.dgvLD.Rows[i].Cells[49].Value.ToString();
                        re.GenDesc20 = this.dgvLD.Rows[i].Cells[50].Value == null ? "" : this.dgvLD.Rows[i].Cells[50].Value.ToString();
                        re.GenDesc21 = this.dgvLD.Rows[i].Cells[51].Value == null ? "" : this.dgvLD.Rows[i].Cells[51].Value.ToString();
                        re.GenDesc22 = this.dgvLD.Rows[i].Cells[52].Value == null ? "" : this.dgvLD.Rows[i].Cells[52].Value.ToString();
                        re.GenDesc23 = this.dgvLD.Rows[i].Cells[53].Value == null ? "" : this.dgvLD.Rows[i].Cells[53].Value.ToString();
                        re.GenDesc24 = this.dgvLD.Rows[i].Cells[54].Value == null ? "" : this.dgvLD.Rows[i].Cells[54].Value.ToString();
                        re.GenDesc25 = this.dgvLD.Rows[i].Cells[55].Value == null ? "" : this.dgvLD.Rows[i].Cells[55].Value.ToString();
                        re.TransactionAmount = this.dgvLD.Rows[i].Cells[56].Value == null ? "" : this.dgvLD.Rows[i].Cells[56].Value.ToString();
                        re.CurrencyCode = this.dgvLD.Rows[i].Cells[57].Value == null ? "" : this.dgvLD.Rows[i].Cells[57].Value.ToString();
                        re.DebitCredit = "";
                        re.BaseAmount = this.dgvLD.Rows[i].Cells[58].Value == null ? "" : this.dgvLD.Rows[i].Cells[58].Value.ToString();
                        re.Base2ReportingAmount = this.dgvLD.Rows[i].Cells[59].Value == null ? "" : this.dgvLD.Rows[i].Cells[59].Value.ToString();
                        re.Value4Amount = this.dgvLD.Rows[i].Cells[60].Value == null ? "" : this.dgvLD.Rows[i].Cells[60].Value.ToString();
                        if (string.IsNullOrEmpty(re.ToString())) continue;
                        if (string.IsNullOrEmpty(re.SaveReference))
                            this.dgvLD.Rows[i].Cells[1].ErrorText = "nullable, but unexpected result would happen when save transaction.";
                        else
                            this.dgvLD.Rows[i].Cells[1].ErrorText = string.Empty;

                        if (string.IsNullOrEmpty(re.LineIndicator))
                            this.dgvLD.Rows[i].Cells[6].ErrorText = "Not null.";
                        else
                            this.dgvLD.Rows[i].Cells[6].ErrorText = string.Empty;

                        cmd2.Parameters.Add(new SqlParameter("@Ledger", re.Ledger));
                        cmd2.Parameters.Add(new SqlParameter("@ft_Account", re.AccountCode));
                        cmd2.Parameters.Add(new SqlParameter("@Period", re.AccountingPeriod));
                        cmd2.Parameters.Add(new SqlParameter("@TransDate", re.TransactionDate));
                        cmd2.Parameters.Add(new SqlParameter("@DueDate", re.DueDate));
                        cmd2.Parameters.Add(new SqlParameter("@JrnlType", re.JournalType));
                        cmd2.Parameters.Add(new SqlParameter("@JrnlSource", re.JournalSource));
                        cmd2.Parameters.Add(new SqlParameter("@TransRef", re.TransactionReference));
                        cmd2.Parameters.Add(new SqlParameter("@Description", re.Description));
                        cmd2.Parameters.Add(new SqlParameter("@AlloctnMarker", re.AllocationMarker));
                        cmd2.Parameters.Add(new SqlParameter("@LA1", re.AnalysisCode1));
                        cmd2.Parameters.Add(new SqlParameter("@LA2", re.AnalysisCode2));
                        cmd2.Parameters.Add(new SqlParameter("@LA3", re.AnalysisCode3));
                        cmd2.Parameters.Add(new SqlParameter("@LA4", re.AnalysisCode4));
                        cmd2.Parameters.Add(new SqlParameter("@LA5", re.AnalysisCode5));
                        cmd2.Parameters.Add(new SqlParameter("@LA6", re.AnalysisCode6));
                        cmd2.Parameters.Add(new SqlParameter("@LA7", re.AnalysisCode7));
                        cmd2.Parameters.Add(new SqlParameter("@LA8", re.AnalysisCode8));
                        cmd2.Parameters.Add(new SqlParameter("@LA9", re.AnalysisCode9));
                        cmd2.Parameters.Add(new SqlParameter("@LA10", re.AnalysisCode10));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc1", re.GenDesc1));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc2", re.GenDesc2));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc3", re.GenDesc3));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc4", re.GenDesc4));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc5", re.GenDesc5));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc6", re.GenDesc6));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc7", re.GenDesc7));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc8", re.GenDesc8));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc9", re.GenDesc9));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc10", re.GenDesc10));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc11", re.GenDesc11));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc12", re.GenDesc12));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc13", re.GenDesc13));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc14", re.GenDesc14));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc15", re.GenDesc15));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc16", re.GenDesc16));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc17", re.GenDesc17));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc18", re.GenDesc18));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc19", re.GenDesc19));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc20", re.GenDesc20));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc21", re.GenDesc21));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc22", re.GenDesc22));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc23", re.GenDesc23));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc24", re.GenDesc24));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc25", re.GenDesc25));
                        cmd2.Parameters.Add(new SqlParameter("@TransAmount", re.TransactionAmount));
                        cmd2.Parameters.Add(new SqlParameter("@Currency", re.CurrencyCode));
                        cmd2.Parameters.Add(new SqlParameter("@BaseAmount", re.BaseAmount));
                        cmd2.Parameters.Add(new SqlParameter("@2ndBase", re.Base2ReportingAmount));
                        cmd2.Parameters.Add(new SqlParameter("@4thAmount", re.Value4Amount));
                        cmd2.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                        cmd2.Parameters.Add(new SqlParameter("@LineIndicator", re.LineIndicator));
                        cmd2.Parameters.Add(new SqlParameter("@StartinginCell", startCell));
                        cmd2.Parameters.Add(new SqlParameter("@BalanceBy", re.BalanceBy));
                        cmd2.Parameters.Add(new SqlParameter("@PopWithJNNumber", re.populatecellwithJN));
                        cmd2.Parameters.Add(new SqlParameter("@Reference", re.Reference));
                        cmd2.Parameters.Add(new SqlParameter("@SaveReference", re.SaveReference));

                        cmd2.Parameters.Add(new SqlParameter("@JournalNumber", journalN));
                        cmd2.Parameters.Add(new SqlParameter("@JournalLineNumber", journalLN));
                        cmd2.Parameters.Add(new SqlParameter("@InputFields", ""));
                        cmd2.Parameters.Add(new SqlParameter("@UpdateFields", ""));
                        string[] arr = consStr[consCurrentCount].Split(',');
                        cmd2.Parameters.Add(new SqlParameter("@ConsolidateBy1", (arr.Count() > 0) ? arr[0] : ""));
                        cmd2.Parameters.Add(new SqlParameter("@ConsolidateBy2", (arr.Count() > 1) ? arr[1] : ""));
                        cmd2.Parameters.Add(new SqlParameter("@ConsolidateBy3", (arr.Count() > 2) ? arr[2] : ""));
                        cmd2.Parameters.Add(new SqlParameter("@ConsolidateBy4", (arr.Count() > 3) ? arr[3] : ""));
                        cmd2.Parameters.Add(new SqlParameter("@Type", "1"));//0.post 1.consolidation 2.update
                        cmd2.Parameters.Add(new SqlParameter("@AllowBalTrans", abt));
                        cmd2.Parameters.Add(new SqlParameter("@AllowPostSuspAcco", apsa));
                        cmd2.Parameters.Add(new SqlParameter("@PostProvisional", pp));
                        if (sender != null)
                        {
                            rdr = cmd2.ExecuteReader();
                            rdr.Close();
                        }
                        cmd2.Parameters.Clear();

                        if (!string.IsNullOrEmpty(re.LineIndicator) && !string.IsNullOrEmpty(startCell))
                        {
                            ll.Add(re);
                            LineIndicatorList.Add(re.LineIndicator);
                            startInCellList.Add(startCell);
                        }
                        consCurrentCount++;
                    }
                    this.Invalidate();
                    if (sender == null)
                    {
                        for (int ii = 0; ii < ll.Count; ii++)
                        {
                            string[] arr = consStr[ii].Split(',');
                            AddConsEntityListIntoFinalList(ll[ii], LineIndicatorList[ii], startInCellList[ii], (arr.Count() > 0) ? arr[0] : "", (arr.Count() > 1) ? arr[1] : "", (arr.Count() > 2) ? arr[2] : "", (arr.Count() > 3) ? arr[3] : "", ws);
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog(typeof(OutputContainer), ex.Message + " - Data Error in Journal tab, Output settings !");
                    throw new Exception(ex.Message + " - Data Error in Journal tab, Output settings !");
                }
                finally
                {
                    if (conn != null)
                    {
                        conn.Close();
                    }
                    if (rdr != null)
                    {
                        rdr.Close();
                    }
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static bool IsNumeric(string value)
        {
            return Regex.IsMatch(value, @"^[+-]?\d*[.]?\d*$");
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static bool IsInt(string value)
        {
            return Regex.IsMatch(value, @"^[+-]?\d*$");
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static bool IsUnsign(string value)
        {
            return Regex.IsMatch(value, @"^\d*[.]?\d*$");
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="list"></param>
        /// <param name="LineIndicatorList"></param>
        /// <param name="StartingInCell"></param>
        /// <param name="groupbyKey1"></param>
        /// <param name="groupbyKey2"></param>
        /// <param name="groupbyKey3"></param>
        /// <param name="groupbyKey4"></param>
        /// <param name="ws"></param>
        private void AddConsEntityListIntoFinalList(Specialist list, string LineIndicator, string StartingInCell, string groupbyKey1, string groupbyKey2, string groupbyKey3, string groupbyKey4, Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            if (string.IsNullOrEmpty(groupbyKey1) && string.IsNullOrEmpty(groupbyKey2) && string.IsNullOrEmpty(groupbyKey3) && string.IsNullOrEmpty(groupbyKey4)) return;
            List<Specialist> tmplist2 = new List<Specialist>();
            List<Specialist> input = new List<Specialist>();
            input.Add(list);
            List<Specialist> newlist = ft.GetEntityListFromDGV(StartingInCell, LineIndicator, input, ws);
            if (newlist != null)
                foreach (Specialist sp in newlist)
                {
                    tmplist2.Add(sp);
                }
            DataTable dt = ft.ToDataTable(tmplist2);//, t5 = "Ledger", t6 = "AccountCode", t7 = "AccountingPeriod", t8 = "AllocationMarker", t9 = "AnalysisCode1", t10 = "AnalysisCode10", t11 = "AnalysisCode2", t12 = "AnalysisCode3", t13 = "AnalysisCode4", t14 = "AnalysisCode5", t15 = "AnalysisCode6", t16 = "AnalysisCode7", t17 = "AnalysisCode8", t18 = "AnalysisCode9", t19 = "CurrencyCode", t20 = "DebitCredit", t21 = "JournalSource", t22 = "JournalType"
            dt.Columns.Add("");
            var query = from t in dt.AsEnumerable()
                        group t by new { t1 = t.Field<string>(string.IsNullOrEmpty(groupbyKey1) ? "Column1" : groupbyKey1), t2 = t.Field<string>(string.IsNullOrEmpty(groupbyKey2) ? "Column1" : groupbyKey2), t3 = t.Field<string>(string.IsNullOrEmpty(groupbyKey3) ? "Column1" : groupbyKey3), t4 = t.Field<string>(string.IsNullOrEmpty(groupbyKey4.Trim()) ? "Column1" : groupbyKey4) } into m
                        select new
                        {
                            Reference = m.First().Field<string>("Reference"),
                            SaveReference = m.First().Field<string>("SaveReference"),
                            populatecellwithJN = m.First().Field<string>("populatecellwithJN"),
                            BalanceBy = m.First().Field<string>("BalanceBy"),
                            Ledger = m.First().Field<string>("Ledger"),
                            AccountCode = m.First().Field<string>("AccountCode"),
                            AccountingPeriod = m.First().Field<string>("AccountingPeriod"),
                            AllocationMarker = m.First().Field<string>("AllocationMarker"),
                            AnalysisCode1 = m.First().Field<string>("AnalysisCode1"),
                            AnalysisCode10 = m.First().Field<string>("AnalysisCode10"),
                            AnalysisCode2 = m.First().Field<string>("AnalysisCode2"),
                            AnalysisCode3 = m.First().Field<string>("AnalysisCode3"),
                            AnalysisCode4 = m.First().Field<string>("AnalysisCode4"),
                            AnalysisCode5 = m.First().Field<string>("AnalysisCode5"),
                            AnalysisCode6 = m.First().Field<string>("AnalysisCode6"),
                            AnalysisCode7 = m.First().Field<string>("AnalysisCode7"),
                            AnalysisCode8 = m.First().Field<string>("AnalysisCode8"),
                            AnalysisCode9 = m.First().Field<string>("AnalysisCode9"),
                            Base2ReportingAmount = m.Sum(k => Decimal.Parse(k.Field<string>("Base2ReportingAmount"))),
                            BaseAmount = m.Sum(k => Decimal.Parse(k.Field<string>("BaseAmount"))),
                            CurrencyCode = m.First().Field<string>("CurrencyCode"),
                            DebitCredit = m.First().Field<string>("DebitCredit"),
                            Description = m.First().Field<string>("Description"),
                            JournalSource = m.First().Field<string>("JournalSource"),
                            JournalType = m.First().Field<string>("JournalType"),
                            TransactionAmount = m.Sum(k => Decimal.Parse(k.Field<string>("TransactionAmount"))),
                            TransactionDate = m.First().Field<string>("TransactionDate"),
                            DueDate = m.First().Field<string>("DueDate"),
                            TransactionReference = m.First().Field<string>("TransactionReference"),
                            Value4Amount = m.Sum(k => Decimal.Parse(k.Field<string>("Value4Amount"))),
                            GeneralDescription1 = m.First().Field<string>("GenDesc1"),
                            GeneralDescription2 = m.First().Field<string>("GenDesc2"),
                            GeneralDescription3 = m.First().Field<string>("GenDesc3"),
                            GeneralDescription4 = m.First().Field<string>("GenDesc4"),
                            GeneralDescription5 = m.First().Field<string>("GenDesc5"),
                            GeneralDescription6 = m.First().Field<string>("GenDesc6"),
                            GeneralDescription7 = m.First().Field<string>("GenDesc7"),
                            GeneralDescription8 = m.First().Field<string>("GenDesc8"),
                            GeneralDescription9 = m.First().Field<string>("GenDesc9"),
                            GeneralDescription10 = m.First().Field<string>("GenDesc10"),
                            GeneralDescription11 = m.First().Field<string>("GenDesc11"),
                            GeneralDescription12 = m.First().Field<string>("GenDesc12"),
                            GeneralDescription13 = m.First().Field<string>("GenDesc13"),
                            GeneralDescription14 = m.First().Field<string>("GenDesc14"),
                            GeneralDescription15 = m.First().Field<string>("GenDesc15"),
                            GeneralDescription16 = m.First().Field<string>("GenDesc16"),
                            GeneralDescription17 = m.First().Field<string>("GenDesc17"),
                            GeneralDescription18 = m.First().Field<string>("GenDesc18"),
                            GeneralDescription19 = m.First().Field<string>("GenDesc19"),
                            GeneralDescription20 = m.First().Field<string>("GenDesc20"),
                            GeneralDescription21 = m.First().Field<string>("GenDesc21"),
                            GeneralDescription22 = m.First().Field<string>("GenDesc22"),
                            GeneralDescription23 = m.First().Field<string>("GenDesc23"),
                            GeneralDescription24 = m.First().Field<string>("GenDesc24"),
                            GeneralDescription25 = m.First().Field<string>("GenDesc25"),
                            rowcount = m.Count(),
                        };
            foreach (var employee in query)
            {
                Specialist re = new Specialist();
                re.Reference = employee.Reference;
                re.SaveReference = employee.SaveReference;
                re.populatecellwithJN = employee.populatecellwithJN;
                re.BalanceBy = employee.BalanceBy;
                re.Ledger = employee.Ledger;
                re.AccountCode = employee.AccountCode;
                re.AccountingPeriod = employee.AccountingPeriod;
                re.TransactionDate = employee.TransactionDate;
                re.DueDate = employee.DueDate;
                re.JournalType = employee.JournalType;
                re.JournalSource = employee.JournalSource;
                re.TransactionReference = employee.TransactionReference;
                re.Description = employee.Description;
                re.AllocationMarker = employee.AllocationMarker;
                re.AnalysisCode1 = employee.AnalysisCode1;
                re.AnalysisCode2 = employee.AnalysisCode2;
                re.AnalysisCode3 = employee.AnalysisCode3;
                re.AnalysisCode4 = employee.AnalysisCode4;
                re.AnalysisCode5 = employee.AnalysisCode5;
                re.AnalysisCode6 = employee.AnalysisCode6;
                re.AnalysisCode7 = employee.AnalysisCode7;
                re.AnalysisCode8 = employee.AnalysisCode8;
                re.AnalysisCode9 = employee.AnalysisCode9;
                re.AnalysisCode10 = employee.AnalysisCode10;
                re.GenDesc1 = employee.GeneralDescription1;
                re.GenDesc2 = employee.GeneralDescription2;
                re.GenDesc3 = employee.GeneralDescription3;
                re.GenDesc4 = employee.GeneralDescription4;
                re.GenDesc5 = employee.GeneralDescription5;
                re.GenDesc6 = employee.GeneralDescription6;
                re.GenDesc7 = employee.GeneralDescription7;
                re.GenDesc8 = employee.GeneralDescription8;
                re.GenDesc9 = employee.GeneralDescription9;
                re.GenDesc10 = employee.GeneralDescription10;
                re.GenDesc11 = employee.GeneralDescription11;
                re.GenDesc12 = employee.GeneralDescription12;
                re.GenDesc13 = employee.GeneralDescription13;
                re.GenDesc14 = employee.GeneralDescription14;
                re.GenDesc15 = employee.GeneralDescription15;
                re.GenDesc16 = employee.GeneralDescription16;
                re.GenDesc17 = employee.GeneralDescription17;
                re.GenDesc18 = employee.GeneralDescription18;
                re.GenDesc19 = employee.GeneralDescription19;
                re.GenDesc20 = employee.GeneralDescription20;
                re.GenDesc21 = employee.GeneralDescription21;
                re.GenDesc22 = employee.GeneralDescription22;
                re.GenDesc23 = employee.GeneralDescription23;
                re.GenDesc24 = employee.GeneralDescription24;
                re.GenDesc25 = employee.GeneralDescription25;
                re.TransactionAmount = employee.TransactionAmount.ToString();
                re.CurrencyCode = employee.CurrencyCode;
                re.DebitCredit = "C";
                re.BaseAmount = employee.BaseAmount.ToString();
                re.Base2ReportingAmount = employee.Base2ReportingAmount.ToString();
                re.Value4Amount = employee.Value4Amount.ToString();
                finallist.Add(re);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="NameValueCollection"></param>
        /// <param name="refnum"></param>
        /// <param name="dtcount"></param>
        private void BindConsolidationDGV(ref DataTable dt, ref List<KeyValuePair<int, string>> NameValueCollection, string refnum, int dtcount)
        {
            try
            {
                DataTable dt2 = ft.GetLineDetailDataFromDB("1", refnum);
                for (int i = 0; i < dt2.Rows.Count; i++)
                {
                    if (dt.Rows.Count == 0)
                    {
                        dt = DataConversionTools.ConvertToDataTableStructure<rsTemplateJournal>();
                    }
                    DataRow dr = dt.NewRow();
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        dr[dt.Columns[j].ColumnName] = dt2.Rows[i][dt.Columns[j].ColumnName];
                    }
                    dt.Rows.Add(dr);
                    if (!string.IsNullOrEmpty(dt2.Rows[i]["ConsolidateBy1"].ToString()))
                        NameValueCollection.Add(new KeyValuePair<int, string>(dtcount + i, dt2.Rows[i]["ConsolidateBy1"].ToString()));
                    else
                        NameValueCollection.Add(new KeyValuePair<int, string>(dtcount + i, ""));
                    if (!string.IsNullOrEmpty(dt2.Rows[i]["ConsolidateBy2"].ToString()))
                        NameValueCollection.Add(new KeyValuePair<int, string>(dtcount + i, dt2.Rows[i]["ConsolidateBy2"].ToString()));
                    else
                        NameValueCollection.Add(new KeyValuePair<int, string>(dtcount + i, ""));
                    if (!string.IsNullOrEmpty(dt2.Rows[i]["ConsolidateBy3"].ToString()))
                        NameValueCollection.Add(new KeyValuePair<int, string>(dtcount + i, dt2.Rows[i]["ConsolidateBy3"].ToString()));
                    else
                        NameValueCollection.Add(new KeyValuePair<int, string>(dtcount + i, ""));
                    if (!string.IsNullOrEmpty(dt2.Rows[i]["ConsolidateBy4"].ToString()))
                        NameValueCollection.Add(new KeyValuePair<int, string>(dtcount + i, dt2.Rows[i]["ConsolidateBy4"].ToString()));
                    else
                        NameValueCollection.Add(new KeyValuePair<int, string>(dtcount + i, ""));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Output settings error");
                LogHelper.WriteLog(typeof(OutputContainer), ex.Message + "Output settings error");
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="NameValueUpdate"></param>
        /// <param name="refnum"></param>
        /// <param name="dtcount"></param>
        private void BindUpdateDGV(ref DataTable dt, ref List<KeyValuePair<int, string>> NameValueUpdate, string refnum, int dtcount)
        {
            try
            {
                DataTable dt2 = ft.GetLineDetailDataFromDB("2", refnum);
                for (int i = 0; i < dt2.Rows.Count; i++)
                {
                    if (dt.Rows.Count == 0)
                    {
                        dt = DataConversionTools.ConvertToDataTableStructure<rsTemplateJournal>();
                    }
                    DataRow dr = dt.NewRow();
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        dr[dt.Columns[j].ColumnName] = dt2.Rows[i][dt.Columns[j].ColumnName];
                    }
                    dt.Rows.Add(dr);
                    string updateStr = dt2.Rows[i]["updateFields"].ToString();
                    string[] sUpdateArray = Regex.Split(updateStr, ",", RegexOptions.IgnoreCase);
                    if (!string.IsNullOrEmpty(updateStr))
                        foreach (string s in sUpdateArray)
                        {
                            NameValueUpdate.Add(new KeyValuePair<int, string>(dtcount + i, s));
                        }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Output settings error");
                LogHelper.WriteLog(typeof(OutputContainer), ex.Message + "Output settings error");
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            DataFieldsSetting dfs = new DataFieldsSetting("Gen");
            dfs.ShowDialog();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void btnSaveCriteria_Click(object sender, EventArgs e)
        {
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsTemplateSetting_Del", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                cmd.ExecuteNonQuery();
                for (int i = 0; i < this.dgvSaveOptions.Rows.Count; i++)
                {
                    string isEmptyStr = string.Empty;
                    for (int k = 0; k < dgvSaveOptions.Rows[i].Cells.Count; k++)
                    {
                        isEmptyStr += dgvSaveOptions.Rows[i].Cells[k].EditedFormattedValue;
                    }
                    if (string.IsNullOrEmpty(isEmptyStr)) continue;

                    string sequencePrifx = string.Empty;
                    string pupulateWithSN = string.Empty;

                    SqlCommand cmd2 = new SqlCommand("rsTemplateSetting_Ins", conn);
                    cmd2.CommandType = CommandType.StoredProcedure;
                    cmd2.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));

                    sequencePrifx = this.dgvSaveOptions.Rows[i].Cells[12].Value == null ? "" : this.dgvSaveOptions.Rows[i].Cells[12].Value.ToString();
                    if (string.IsNullOrEmpty(sequencePrifx))
                        this.dgvSaveOptions.Rows[i].Cells[12].ErrorText = "nullable, but unexpected result would happen when save transaction/view transaction.";
                    else
                        this.dgvSaveOptions.Rows[i].Cells[12].ErrorText = string.Empty;

                    pupulateWithSN = this.dgvSaveOptions.Rows[i].Cells[13].Value == null ? "" : this.dgvSaveOptions.Rows[i].Cells[13].Value.ToString();
                    if (string.IsNullOrEmpty(pupulateWithSN))
                        this.dgvSaveOptions.Rows[i].Cells[13].ErrorText = "nullable, but unexpected result would happen when save transaction/view transaction.";
                    else
                        this.dgvSaveOptions.Rows[i].Cells[13].ErrorText = string.Empty;

                    cmd2.Parameters.Add(new SqlParameter("@CriteriaName1", this.dgvSaveOptions.Rows[i].Cells[1].Value == null ? "" : this.dgvSaveOptions.Rows[i].Cells[1].Value.ToString()));
                    cmd2.Parameters.Add(new SqlParameter("@CellReference1", this.dgvSaveOptions.Rows[i].Cells[2].Value == null ? "" : this.dgvSaveOptions.Rows[i].Cells[2].Value.ToString()));
                    cmd2.Parameters.Add(new SqlParameter("@CriteriaName2", this.dgvSaveOptions.Rows[i].Cells[3].Value == null ? "" : this.dgvSaveOptions.Rows[i].Cells[3].Value.ToString()));
                    cmd2.Parameters.Add(new SqlParameter("@CellReference2", this.dgvSaveOptions.Rows[i].Cells[4].Value == null ? "" : this.dgvSaveOptions.Rows[i].Cells[4].Value.ToString()));
                    cmd2.Parameters.Add(new SqlParameter("@CriteriaName3", this.dgvSaveOptions.Rows[i].Cells[5].Value == null ? "" : this.dgvSaveOptions.Rows[i].Cells[5].Value.ToString()));
                    cmd2.Parameters.Add(new SqlParameter("@CellReference3", this.dgvSaveOptions.Rows[i].Cells[6].Value == null ? "" : this.dgvSaveOptions.Rows[i].Cells[6].Value.ToString()));
                    cmd2.Parameters.Add(new SqlParameter("@CriteriaName4", this.dgvSaveOptions.Rows[i].Cells[7].Value == null ? "" : this.dgvSaveOptions.Rows[i].Cells[7].Value.ToString()));
                    cmd2.Parameters.Add(new SqlParameter("@CellReference4", this.dgvSaveOptions.Rows[i].Cells[8].Value == null ? "" : this.dgvSaveOptions.Rows[i].Cells[8].Value.ToString()));
                    cmd2.Parameters.Add(new SqlParameter("@CriteriaName5", this.dgvSaveOptions.Rows[i].Cells[9].Value == null ? "" : this.dgvSaveOptions.Rows[i].Cells[9].Value.ToString()));
                    cmd2.Parameters.Add(new SqlParameter("@CellReference5", this.dgvSaveOptions.Rows[i].Cells[10].Value == null ? "" : this.dgvSaveOptions.Rows[i].Cells[10].Value.ToString()));
                    cmd2.Parameters.Add(new SqlParameter("@OpenTransUponSave", this.dgvSaveOptions.Rows[i].Cells[11].Value == null ? "" : this.dgvSaveOptions.Rows[i].Cells[11].Value.ToString()));
                    cmd2.Parameters.Add(new SqlParameter("@Reference", this.dgvSaveOptions.Rows[i].Cells[0].Value == null ? "" : this.dgvSaveOptions.Rows[i].Cells[0].Value.ToString()));
                    if (!string.IsNullOrEmpty(sequencePrifx))
                        cmd2.Parameters.Add(new SqlParameter("@UseSequenceNumbering", true));
                    else
                        cmd2.Parameters.Add(new SqlParameter("@UseSequenceNumbering", false));

                    cmd2.Parameters.Add(new SqlParameter("@SequencePrefix", sequencePrifx));
                    cmd2.Parameters.Add(new SqlParameter("@PopulateCell", pupulateWithSN));
                    cmd2.Parameters.Add(new SqlParameter("@PDFFolder", this.dgvSaveOptions.Rows[i].Cells[14].Value == null ? "" : this.dgvSaveOptions.Rows[i].Cells[14].Value.ToString()));
                    cmd2.Parameters.Add(new SqlParameter("@PDFName", this.dgvSaveOptions.Rows[i].Cells[15].Value == null ? "" : this.dgvSaveOptions.Rows[i].Cells[15].Value.ToString()));
                    cmd2.ExecuteNonQuery();
                    cmd2.Parameters.Clear();
                }
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                }
            }
            if (sender != null)
                Ribbon2._MyOutputCustomTaskPane.Visible = false;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="ws"></param>
        public void SaveTransUpd(object sender, Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            searchStatus = string.Empty;
            string cellName = string.Empty;
            updateStatus = string.Empty;
            updateStatusForPost = string.Empty;
            List<string> LineIndicatorList = new List<string>();
            List<string> startInCellList = new List<string>();
            if (!string.IsNullOrEmpty(SessionInfo.UserInfo.File_ftid))
            {
                SqlConnection conn = null;
                SqlDataReader rdr = null;
                try
                {
                    conn = new SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                    conn.Open();
                    SqlCommand cmd = new SqlCommand("rsTemplateJournal_Del", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                    cmd.Parameters.Add(new SqlParameter("@Type", "2"));
                    if (sender != null)
                    {
                        rdr = cmd.ExecuteReader();
                        rdr.Close();
                    }
                    TransUpdFinallist.Clear();
                    SqlCommand cmd2 = new SqlCommand("rsTemplateJournal_Ins", conn);
                    cmd2.CommandType = CommandType.StoredProcedure;
                    for (int j = 0; j < dgvLD.Rows.Count; j++)
                    {
                        bool iscontinue = true;
                        searchStatus = "";
                        updateStatus = "";
                        for (int i = 0; i < dgvLD.Columns.Count; i++)
                        {
                            if (dgvLD.Rows[j].Cells[i].Style.BackColor == Color.Aqua)
                            {
                                iscontinue = false;
                                updateStatus += dgvLD.Columns[i].Name + ",";
                                updateStatusForPost += dgvLD.Columns[i].Name + ",";
                            }
                        }
                        if (iscontinue)
                            continue;

                        string isEmptyStr = string.Empty;
                        for (int k = 0; k < dgvLD.Rows[j].Cells.Count; k++)
                        {
                            isEmptyStr += dgvLD.Rows[j].Cells[k].EditedFormattedValue;
                            if (k > 6 && (dgvLD.Rows[j].Cells[k].Style.BackColor != Color.Aqua))
                            {
                                object tmp = dgvLD.Rows[j].Cells[k].EditedFormattedValue;
                                if (!string.IsNullOrEmpty(tmp.ToString()))
                                    searchStatus += dgvLD.Columns[k].DataPropertyName + ",";
                            }
                        }
                        if (string.IsNullOrEmpty(isEmptyStr)) continue;
                        ExcelAddIn4.Common2.Specialist re = new ExcelAddIn4.Common2.Specialist();
                        re.Actions = new Common2.Actions();

                        string startCell = string.Empty;
                        string journalN = string.Empty;
                        string journalLN = string.Empty;
                        string abt = string.Empty;
                        string apsa = string.Empty;
                        string pp = string.Empty;
                        re.Reference = this.dgvLD.Rows[j].Cells[0].Value == null ? "" : this.dgvLD.Rows[j].Cells[0].Value.ToString().Replace(" ", "");
                        re.SaveReference = this.dgvLD.Rows[j].Cells[1].Value == null ? "" : this.dgvLD.Rows[j].Cells[1].Value.ToString();
                        re.BalanceBy = this.dgvLD.Rows[j].Cells[2].Value == null ? "" : this.dgvLD.Rows[j].Cells[2].Value.ToString();
                        abt = this.dgvLD.Rows[j].Cells[3].Value == null ? "" : this.dgvLD.Rows[j].Cells[3].Value.ToString();
                        apsa = this.dgvLD.Rows[j].Cells[4].Value == null ? "" : this.dgvLD.Rows[j].Cells[4].Value.ToString();
                        pp = this.dgvLD.Rows[j].Cells[5].Value == null ? "" : this.dgvLD.Rows[j].Cells[5].Value.ToString();

                        re.LineIndicator = this.dgvLD.Rows[j].Cells[6].Value == null ? "" : this.dgvLD.Rows[j].Cells[6].Value.ToString();
                        startCell = this.dgvLD.Rows[j].Cells[7].Value == null ? "" : this.dgvLD.Rows[j].Cells[7].Value.ToString();
                        re.StartInCell = startCell;
                        re.populatecellwithJN = this.dgvLD.Rows[j].Cells[8].Value == null ? "" : this.dgvLD.Rows[j].Cells[8].Value.ToString();
                        journalN = this.dgvLD.Rows[j].Cells[9].Value == null ? "" : this.dgvLD.Rows[j].Cells[9].Value.ToString();
                        re.JournalNumber = journalN;
                        journalLN = this.dgvLD.Rows[j].Cells[10].Value == null ? "" : this.dgvLD.Rows[j].Cells[10].Value.ToString();
                        re.JournalLineNumber = journalLN;

                        re.Ledger = this.dgvLD.Rows[j].Cells[11].Value == null ? "" : this.dgvLD.Rows[j].Cells[11].Value.ToString();
                        re.AccountCode = this.dgvLD.Rows[j].Cells[12].Value == null ? "" : this.dgvLD.Rows[j].Cells[12].Value.ToString();
                        re.AccountingPeriod = this.dgvLD.Rows[j].Cells[13].Value == null ? "" : this.dgvLD.Rows[j].Cells[13].Value.ToString();
                        re.TransactionDate = this.dgvLD.Rows[j].Cells[14].Value == null ? "" : this.dgvLD.Rows[j].Cells[14].Value.ToString();
                        re.DueDate = this.dgvLD.Rows[j].Cells[15].Value == null ? "" : this.dgvLD.Rows[j].Cells[15].Value.ToString();
                        re.JournalType = this.dgvLD.Rows[j].Cells[16].Value == null ? "" : this.dgvLD.Rows[j].Cells[16].Value.ToString();
                        re.JournalSource = this.dgvLD.Rows[j].Cells[17].Value == null ? "" : this.dgvLD.Rows[j].Cells[17].Value.ToString();
                        re.TransactionReference = this.dgvLD.Rows[j].Cells[18].Value == null ? "" : this.dgvLD.Rows[j].Cells[18].Value.ToString();
                        re.Description = this.dgvLD.Rows[j].Cells[19].Value == null ? "" : this.dgvLD.Rows[j].Cells[19].Value.ToString();
                        re.AllocationMarker = this.dgvLD.Rows[j].Cells[20].Value == null ? "" : this.dgvLD.Rows[j].Cells[20].Value.ToString();
                        re.AnalysisCode1 = this.dgvLD.Rows[j].Cells[21].Value == null ? "" : this.dgvLD.Rows[j].Cells[21].Value.ToString();
                        re.AnalysisCode2 = this.dgvLD.Rows[j].Cells[22].Value == null ? "" : this.dgvLD.Rows[j].Cells[22].Value.ToString();
                        re.AnalysisCode3 = this.dgvLD.Rows[j].Cells[23].Value == null ? "" : this.dgvLD.Rows[j].Cells[23].Value.ToString();
                        re.AnalysisCode4 = this.dgvLD.Rows[j].Cells[24].Value == null ? "" : this.dgvLD.Rows[j].Cells[24].Value.ToString();
                        re.AnalysisCode5 = this.dgvLD.Rows[j].Cells[25].Value == null ? "" : this.dgvLD.Rows[j].Cells[25].Value.ToString();
                        re.AnalysisCode6 = this.dgvLD.Rows[j].Cells[26].Value == null ? "" : this.dgvLD.Rows[j].Cells[26].Value.ToString();
                        re.AnalysisCode7 = this.dgvLD.Rows[j].Cells[27].Value == null ? "" : this.dgvLD.Rows[j].Cells[27].Value.ToString();
                        re.AnalysisCode8 = this.dgvLD.Rows[j].Cells[28].Value == null ? "" : this.dgvLD.Rows[j].Cells[28].Value.ToString();
                        re.AnalysisCode9 = this.dgvLD.Rows[j].Cells[29].Value == null ? "" : this.dgvLD.Rows[j].Cells[29].Value.ToString();
                        re.AnalysisCode10 = this.dgvLD.Rows[j].Cells[30].Value == null ? "" : this.dgvLD.Rows[j].Cells[30].Value.ToString();
                        re.GenDesc1 = this.dgvLD.Rows[j].Cells[31].Value == null ? "" : this.dgvLD.Rows[j].Cells[31].Value.ToString();
                        re.GenDesc2 = this.dgvLD.Rows[j].Cells[32].Value == null ? "" : this.dgvLD.Rows[j].Cells[32].Value.ToString();
                        re.GenDesc3 = this.dgvLD.Rows[j].Cells[33].Value == null ? "" : this.dgvLD.Rows[j].Cells[33].Value.ToString();
                        re.GenDesc4 = this.dgvLD.Rows[j].Cells[34].Value == null ? "" : this.dgvLD.Rows[j].Cells[34].Value.ToString();
                        re.GenDesc5 = this.dgvLD.Rows[j].Cells[35].Value == null ? "" : this.dgvLD.Rows[j].Cells[35].Value.ToString();
                        re.GenDesc6 = this.dgvLD.Rows[j].Cells[36].Value == null ? "" : this.dgvLD.Rows[j].Cells[36].Value.ToString();
                        re.GenDesc7 = this.dgvLD.Rows[j].Cells[37].Value == null ? "" : this.dgvLD.Rows[j].Cells[37].Value.ToString();
                        re.GenDesc8 = this.dgvLD.Rows[j].Cells[38].Value == null ? "" : this.dgvLD.Rows[j].Cells[38].Value.ToString();
                        re.GenDesc9 = this.dgvLD.Rows[j].Cells[39].Value == null ? "" : this.dgvLD.Rows[j].Cells[39].Value.ToString();
                        re.GenDesc10 = this.dgvLD.Rows[j].Cells[40].Value == null ? "" : this.dgvLD.Rows[j].Cells[40].Value.ToString();
                        re.GenDesc11 = this.dgvLD.Rows[j].Cells[41].Value == null ? "" : this.dgvLD.Rows[j].Cells[41].Value.ToString();
                        re.GenDesc12 = this.dgvLD.Rows[j].Cells[42].Value == null ? "" : this.dgvLD.Rows[j].Cells[42].Value.ToString();
                        re.GenDesc13 = this.dgvLD.Rows[j].Cells[43].Value == null ? "" : this.dgvLD.Rows[j].Cells[43].Value.ToString();
                        re.GenDesc14 = this.dgvLD.Rows[j].Cells[44].Value == null ? "" : this.dgvLD.Rows[j].Cells[44].Value.ToString();
                        re.GenDesc15 = this.dgvLD.Rows[j].Cells[45].Value == null ? "" : this.dgvLD.Rows[j].Cells[45].Value.ToString();
                        re.GenDesc16 = this.dgvLD.Rows[j].Cells[46].Value == null ? "" : this.dgvLD.Rows[j].Cells[46].Value.ToString();
                        re.GenDesc17 = this.dgvLD.Rows[j].Cells[47].Value == null ? "" : this.dgvLD.Rows[j].Cells[47].Value.ToString();
                        re.GenDesc18 = this.dgvLD.Rows[j].Cells[48].Value == null ? "" : this.dgvLD.Rows[j].Cells[48].Value.ToString();
                        re.GenDesc19 = this.dgvLD.Rows[j].Cells[49].Value == null ? "" : this.dgvLD.Rows[j].Cells[49].Value.ToString();
                        re.GenDesc20 = this.dgvLD.Rows[j].Cells[50].Value == null ? "" : this.dgvLD.Rows[j].Cells[50].Value.ToString();
                        re.GenDesc21 = this.dgvLD.Rows[j].Cells[51].Value == null ? "" : this.dgvLD.Rows[j].Cells[51].Value.ToString();
                        re.GenDesc22 = this.dgvLD.Rows[j].Cells[52].Value == null ? "" : this.dgvLD.Rows[j].Cells[52].Value.ToString();
                        re.GenDesc23 = this.dgvLD.Rows[j].Cells[53].Value == null ? "" : this.dgvLD.Rows[j].Cells[53].Value.ToString();
                        re.GenDesc24 = this.dgvLD.Rows[j].Cells[54].Value == null ? "" : this.dgvLD.Rows[j].Cells[54].Value.ToString();
                        re.GenDesc25 = this.dgvLD.Rows[j].Cells[55].Value == null ? "" : this.dgvLD.Rows[j].Cells[55].Value.ToString();
                        re.TransactionAmount = this.dgvLD.Rows[j].Cells[56].Value == null ? "" : this.dgvLD.Rows[j].Cells[56].Value.ToString();
                        re.CurrencyCode = this.dgvLD.Rows[j].Cells[57].Value == null ? "" : this.dgvLD.Rows[j].Cells[57].Value.ToString();
                        re.DebitCredit = "";
                        re.BaseAmount = this.dgvLD.Rows[j].Cells[58].Value == null ? "" : this.dgvLD.Rows[j].Cells[58].Value.ToString();
                        re.Base2ReportingAmount = this.dgvLD.Rows[j].Cells[59].Value == null ? "" : this.dgvLD.Rows[j].Cells[59].Value.ToString();
                        re.Value4Amount = this.dgvLD.Rows[j].Cells[60].Value == null ? "" : this.dgvLD.Rows[j].Cells[60].Value.ToString();
                        if (string.IsNullOrEmpty(re.ToString())) continue;

                        if (string.IsNullOrEmpty(re.LineIndicator))
                            this.dgvLD.Rows[j].Cells[6].ErrorText = "Not null.";
                        else
                            this.dgvLD.Rows[j].Cells[6].ErrorText = string.Empty;

                        cmd2.Parameters.Add(new SqlParameter("@Ledger", re.Ledger));
                        cmd2.Parameters.Add(new SqlParameter("@ft_Account", re.AccountCode));
                        cmd2.Parameters.Add(new SqlParameter("@Period", re.AccountingPeriod));
                        cmd2.Parameters.Add(new SqlParameter("@TransDate", re.TransactionDate));
                        cmd2.Parameters.Add(new SqlParameter("@DueDate", re.DueDate));
                        cmd2.Parameters.Add(new SqlParameter("@JrnlType", re.JournalType));
                        cmd2.Parameters.Add(new SqlParameter("@JrnlSource", re.JournalSource));
                        cmd2.Parameters.Add(new SqlParameter("@TransRef", re.TransactionReference));
                        cmd2.Parameters.Add(new SqlParameter("@Description", re.Description));
                        cmd2.Parameters.Add(new SqlParameter("@AlloctnMarker", re.AllocationMarker));
                        cmd2.Parameters.Add(new SqlParameter("@LA1", re.AnalysisCode1));
                        cmd2.Parameters.Add(new SqlParameter("@LA2", re.AnalysisCode2));
                        cmd2.Parameters.Add(new SqlParameter("@LA3", re.AnalysisCode3));
                        cmd2.Parameters.Add(new SqlParameter("@LA4", re.AnalysisCode4));
                        cmd2.Parameters.Add(new SqlParameter("@LA5", re.AnalysisCode5));
                        cmd2.Parameters.Add(new SqlParameter("@LA6", re.AnalysisCode6));
                        cmd2.Parameters.Add(new SqlParameter("@LA7", re.AnalysisCode7));
                        cmd2.Parameters.Add(new SqlParameter("@LA8", re.AnalysisCode8));
                        cmd2.Parameters.Add(new SqlParameter("@LA9", re.AnalysisCode9));
                        cmd2.Parameters.Add(new SqlParameter("@LA10", re.AnalysisCode10));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc1", re.GenDesc1));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc2", re.GenDesc2));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc3", re.GenDesc3));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc4", re.GenDesc4));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc5", re.GenDesc5));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc6", re.GenDesc6));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc7", re.GenDesc7));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc8", re.GenDesc8));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc9", re.GenDesc9));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc10", re.GenDesc10));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc11", re.GenDesc11));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc12", re.GenDesc12));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc13", re.GenDesc13));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc14", re.GenDesc14));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc15", re.GenDesc15));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc16", re.GenDesc16));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc17", re.GenDesc17));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc18", re.GenDesc18));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc19", re.GenDesc19));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc20", re.GenDesc20));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc21", re.GenDesc21));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc22", re.GenDesc22));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc23", re.GenDesc23));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc24", re.GenDesc24));
                        cmd2.Parameters.Add(new SqlParameter("@GenDesc25", re.GenDesc25));
                        cmd2.Parameters.Add(new SqlParameter("@TransAmount", re.TransactionAmount));
                        cmd2.Parameters.Add(new SqlParameter("@Currency", re.CurrencyCode));
                        cmd2.Parameters.Add(new SqlParameter("@BaseAmount", re.BaseAmount));
                        cmd2.Parameters.Add(new SqlParameter("@2ndBase", re.Base2ReportingAmount));
                        cmd2.Parameters.Add(new SqlParameter("@4thAmount", re.Value4Amount));
                        cmd2.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                        cmd2.Parameters.Add(new SqlParameter("@LineIndicator", re.LineIndicator));
                        cmd2.Parameters.Add(new SqlParameter("@StartinginCell", startCell));
                        cmd2.Parameters.Add(new SqlParameter("@BalanceBy", re.BalanceBy));
                        cmd2.Parameters.Add(new SqlParameter("@PopWithJNNumber", re.populatecellwithJN));
                        cmd2.Parameters.Add(new SqlParameter("@Reference", re.Reference));
                        cmd2.Parameters.Add(new SqlParameter("@SaveReference", re.SaveReference));

                        cmd2.Parameters.Add(new SqlParameter("@JournalNumber", journalN));
                        cmd2.Parameters.Add(new SqlParameter("@JournalLineNumber", journalLN));
                        cmd2.Parameters.Add(new SqlParameter("@InputFields", searchStatus));
                        cmd2.Parameters.Add(new SqlParameter("@UpdateFields", updateStatus));
                        cmd2.Parameters.Add(new SqlParameter("@ConsolidateBy1", ""));
                        cmd2.Parameters.Add(new SqlParameter("@ConsolidateBy2", ""));
                        cmd2.Parameters.Add(new SqlParameter("@ConsolidateBy3", ""));
                        cmd2.Parameters.Add(new SqlParameter("@ConsolidateBy4", ""));
                        cmd2.Parameters.Add(new SqlParameter("@Type", "2"));//0.post 1.consolidation 2.update
                        cmd2.Parameters.Add(new SqlParameter("@AllowBalTrans", abt));
                        cmd2.Parameters.Add(new SqlParameter("@AllowPostSuspAcco", apsa));
                        cmd2.Parameters.Add(new SqlParameter("@PostProvisional", pp));
                        if (sender != null)
                        {
                            rdr = cmd2.ExecuteReader();
                            rdr.Close();
                        }
                        cmd2.Parameters.Clear();

                        if (!string.IsNullOrEmpty(re.LineIndicator) && !string.IsNullOrEmpty(startCell))
                        {
                            TransUpdFinallist.Add(re);
                            LineIndicatorList.Add(re.LineIndicator);
                            startInCellList.Add(startCell);
                        }
                    }
                    this.Invalidate();
                    //if (LineIndicatorList.Count == 0 && TransUpdFinallist.Count != 0)
                    //{
                    //throw new Exception("No data found for specified Line Indicator(s)!");
                    //}
                    if (sender == null)
                        AddLineDetailEntityListIntoFinalListForTransUpd(TransUpdFinallist, LineIndicatorList, startInCellList, ws, ref TransUpdFinallist);
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog(typeof(OutputContainer), ex.Message + " - Data Error in Journal tab, Output settings !" + " Output settings  error");
                    throw new Exception(ex.Message + " - Data Error in Journal tab, Output settings !");
                }
                finally
                {
                    if (conn != null)
                    {
                        conn.Close();
                    }
                    if (rdr != null)
                    {
                        rdr.Close();
                    }
                }
            }
            Ribbon2._MyOutputCustomTaskPane.Visible = false;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private bool isTransUpdate()
        {
            bool isUpdate = false;
            for (int i = 0; i < dgvLD.Columns.Count; i++)
                if (dgvLD.SelectedRows[0].Cells[i].Style.BackColor == Color.Aqua)
                    isUpdate = true;

            return isUpdate;
        }
        /// <summary>
        /// 
        /// </summary>
        private void ClearSession()
        {
            SessionInfo.UserInfo.Criteria1 = "";
            SessionInfo.UserInfo.Criteria2 = "";
            SessionInfo.UserInfo.Criteria3 = "";
            SessionInfo.UserInfo.Criteria4 = "";
            SessionInfo.UserInfo.Criteria5 = "";
            SessionInfo.UserInfo.CellReference1 = "";
            SessionInfo.UserInfo.CellReference2 = "";
            SessionInfo.UserInfo.CellReference3 = "";
            SessionInfo.UserInfo.CellReference4 = "";
            SessionInfo.UserInfo.CellReference5 = "";
            SessionInfo.UserInfo.UseCriteria = false;
            SessionInfo.UserInfo.OpentransuponSave = false;
            SessionInfo.UserInfo.SequencePrefix = "";
            SessionInfo.UserInfo.UseSequenceNumbering = "0";
            SessionInfo.UserInfo.PopulateCell = "";
        }
        /// <summary>
        /// 
        /// </summary>
        public void SetSession()
        {
            ClearSession();
            if (string.IsNullOrEmpty(SessionInfo.UserInfo.CurrentSaveRef)) return;
            DataTable dt = ft.GetReportCriteriaByRef(SessionInfo.UserInfo.File_ftid, SessionInfo.UserInfo.CurrentSaveRef);
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("Save Reference " + SessionInfo.UserInfo.CurrentSaveRef + " error!"); return;
            }
            SessionInfo.UserInfo.OpentransuponSave = (bool)dt.Rows[0]["OpenTransUponSave"];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                try
                {
                    if (!string.IsNullOrEmpty(dt.Rows[i]["SequencePrefix"].ToString()))
                    {
                        SessionInfo.UserInfo.SequencePrefix = dt.Rows[i]["SequencePrefix"].ToString();
                        if (!string.IsNullOrEmpty(SessionInfo.UserInfo.SequencePrefix))
                            SessionInfo.UserInfo.UseSequenceNumbering = "1";
                        else
                            SessionInfo.UserInfo.UseSequenceNumbering = "0";
                    }
                    if (!string.IsNullOrEmpty(dt.Rows[i]["PopulateCell"].ToString()))
                        SessionInfo.UserInfo.PopulateCell = dt.Rows[i]["PopulateCell"].ToString();
                    if (!string.IsNullOrEmpty(dt.Rows[i]["CriteriaName1"].ToString()))
                    {
                        SessionInfo.UserInfo.Criteria1 = dt.Rows[i]["CriteriaName1"].ToString();
                        var cellvalue = "";
                        try
                        {
                            cellvalue = ft.GetValueOfAddress(dt.Rows[i]["CellReference1"].ToString());
                        }
                        catch { cellvalue = dt.Rows[i]["CellReference1"].ToString(); }
                        SessionInfo.UserInfo.CellReference1 = cellvalue;
                        SessionInfo.UserInfo.UseCriteria = true;
                    }
                    else
                    {
                        SessionInfo.UserInfo.Criteria1 = "";
                        SessionInfo.UserInfo.CellReference1 = "";
                    }
                    if (!string.IsNullOrEmpty(dt.Rows[i]["CriteriaName2"].ToString()))
                    {
                        SessionInfo.UserInfo.Criteria2 = dt.Rows[i]["CriteriaName2"].ToString();
                        var cellvalue = "";
                        try
                        {
                            cellvalue = ft.GetValueOfAddress(dt.Rows[i]["CellReference2"].ToString());
                        }
                        catch { cellvalue = dt.Rows[i]["CellReference2"].ToString(); }
                        SessionInfo.UserInfo.CellReference2 = cellvalue;
                        SessionInfo.UserInfo.UseCriteria = true;
                    }
                    else
                    {
                        SessionInfo.UserInfo.Criteria2 = "";
                        SessionInfo.UserInfo.CellReference2 = "";
                    }
                    if (!string.IsNullOrEmpty(dt.Rows[i]["CriteriaName3"].ToString()))
                    {
                        SessionInfo.UserInfo.Criteria3 = dt.Rows[i]["CriteriaName3"].ToString();
                        var cellvalue = "";
                        try
                        {
                            cellvalue = ft.GetValueOfAddress(dt.Rows[i]["CellReference3"].ToString());
                        }
                        catch { cellvalue = dt.Rows[i]["CellReference3"].ToString(); }
                        SessionInfo.UserInfo.CellReference3 = cellvalue;
                        SessionInfo.UserInfo.UseCriteria = true;
                    }
                    else
                    {
                        SessionInfo.UserInfo.Criteria3 = "";
                        SessionInfo.UserInfo.CellReference3 = "";
                    }
                    if (!string.IsNullOrEmpty(dt.Rows[i]["CriteriaName4"].ToString()))
                    {
                        SessionInfo.UserInfo.Criteria4 = dt.Rows[i]["CriteriaName4"].ToString();
                        var cellvalue = "";
                        try
                        {
                            cellvalue = ft.GetValueOfAddress(dt.Rows[i]["CellReference4"].ToString());
                        }
                        catch { cellvalue = dt.Rows[i]["CellReference4"].ToString(); }
                        SessionInfo.UserInfo.CellReference4 = cellvalue;
                        SessionInfo.UserInfo.UseCriteria = true;
                    }
                    else
                    {
                        SessionInfo.UserInfo.Criteria4 = "";
                        SessionInfo.UserInfo.CellReference4 = "";
                    }
                    if (!string.IsNullOrEmpty(dt.Rows[i]["CriteriaName5"].ToString()))
                    {
                        SessionInfo.UserInfo.Criteria5 = dt.Rows[i]["CriteriaName5"].ToString();
                        var cellvalue = "";
                        try
                        {
                            cellvalue = ft.GetValueOfAddress(dt.Rows[i]["CellReference5"].ToString());
                        }
                        catch { cellvalue = dt.Rows[i]["CellReference5"].ToString(); }
                        SessionInfo.UserInfo.CellReference5 = cellvalue;
                        SessionInfo.UserInfo.UseCriteria = true;
                    }
                    else
                    {
                        SessionInfo.UserInfo.Criteria5 = "";
                        SessionInfo.UserInfo.CellReference5 = "";
                    }
                }
                catch { }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTestJournal_Click(object sender, EventArgs e)
        {
            if (dgvLD.SelectedRows.Count == 0)
            {
                MessageBox.Show("Please click the row header number ( before Ref column ) to choose a certain Reference number and enjoy your test ! ", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            else
            {
                if (dgvLD.SelectedRows[0].Cells[0].Value != null)
                    SessionInfo.UserInfo.CurrentRef = dgvLD.SelectedRows[0].Cells[0].Value.ToString().Replace(" ", "");
                if (dgvLD.SelectedRows[0].Cells[2].Value != null)
                    SessionInfo.UserInfo.BalanceBy = dgvLD.SelectedRows[0].Cells[2].Value.ToString().Replace(" ", "");
            }
            if (isTransUpdate())
            {
                btnTestTransUpd_Click(sender, null);
                return;
            }
            if (ft.IsGUID(Path.GetFileNameWithoutExtension(SessionInfo.UserInfo.CachePath)) && !string.IsNullOrEmpty(ft.ProcessJournalNumber()))//(SessionInfo.UserInfo.CachePath != SessionInfo.UserInfo.FilePath)
            {
                MessageBox.Show("Can't be changed! This document has been Posted! ", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LogHelper.WriteLog(typeof(OutputContainer), "Can't be changed! This document has been Posted! ");
                return;
            }
            DateTime starttime = DateTime.Now;
            if (Ribbon2.xpf != null)
                Ribbon2.xpf.Dispose();
            Ribbon2.xpf = new XMLPostFrm();
            Ribbon2.xpf.panel3.Visible = true;
            Ribbon2.xpf.panel1.Visible = false;
            Ribbon2.xpf.panel2.Visible = false;
            Ribbon2.xpf.panel3.Dock = DockStyle.Fill;
            Ribbon2.xpf.ControlBox = false;
            try
            {
                Globals.ThisAddIn.Application.DisplayAlerts = false;
                Ribbon2.wsRrigin = (Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                var lastColumn = Ribbon2.wsRrigin.Cells.Find("*", Ribbon2.wsRrigin.Cells[1, 1], Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, Microsoft.Office.Interop.Excel.XlLookAt.xlPart, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, Type.Missing, Type.Missing);
                Ribbon2.LastColumnName = Finance_Tools.RemoveNumber(lastColumn.Address).Replace("$", "");
                Finance_Tools.MaxColumnCount = lastColumn.Column;
                var lastrow = Ribbon2.wsRrigin.Cells.Find("*", Ribbon2.wsRrigin.Cells[1, 1], Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, Microsoft.Office.Interop.Excel.XlLookAt.xlPart, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, Type.Missing, Type.Missing);
                Ribbon2.LastRowNumber = Finance_Tools.RemoveNotNumber(lastrow.Address);

                Globals.ThisAddIn.Application.ScreenUpdating = false;
                Ribbon2.xpf.Show();
                XMLPostFrm.richTextBox1.Text += "Error List :\r\n";
                Save(null, Ribbon2.wsRrigin);
                SaveCons(null, Ribbon2.wsRrigin);
                SetSession();
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                Globals.ThisAddIn.Application.DisplayAlerts = true;
                if (XMLPostFrm.richTextBox1.Text.Length > 21)
                {
                    LogHelper.WriteLog(typeof(OutputContainer), XMLPostFrm.richTextBox1.Text + " - Post Journal Processing error , Template:" + SessionInfo.UserInfo.FileName);
                }
            }
            catch (Exception ex)
            {
                if (ex.ToString().Contains("Clipboard"))
                {
                    MessageBox.Show("Clipboard not ready, please try again.", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(ex.ToString(), "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                LogHelper.WriteLog(typeof(OutputContainer), ex.Message + "Post error");
                Ribbon2.xpf.Dispose();
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                Clipboard.SetText("\r\n");
            }
            finally
            {
                Ribbon2.xpf.Focus();
                Ribbon2.xpf.bddata();
                Ribbon2.xpf.panel3.Visible = false;
                Ribbon2.xpf.panel1.Visible = true;
                Ribbon2.xpf.panel2.Visible = true;
                Ribbon2.xpf.ControlBox = true;
                DateTime stoptime = DateTime.Now;
                string costtime = Finance_Tools.DateDiff(starttime, stoptime);
                Globals.ThisAddIn.Application.StatusBar = "latest posting process costs " + costtime;
                GC.Collect();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CTF_btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                if (cbXMLOrText.Text == "XML")
                    SaveXML(sender, null);
                else
                    SaveCTF(sender, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            Ribbon2._MyOutputCustomTaskPane.Visible = false;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="ws"></param>
        public void SaveXML(object sender, Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            string HeaderText = string.Empty;
            string HeaderValue = string.Empty;
            List<string> LineIndicatorList = new List<string>();
            List<string> startInCellList = new List<string>();
            for (int k = 0; k < dgvCreateTextFile.Columns.Count; k++)
            {
                HeaderText += dgvCreateTextFile.Columns[k].HeaderText + ",";
                HeaderValue += dgvCreateTextFile.Columns[k].Name + ",";
            }
            if (!string.IsNullOrEmpty(SessionInfo.UserInfo.File_ftid))
            {
                SqlConnection conn = null;
                SqlDataReader rdr = null;
                try
                {
                    conn = new
                        SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                    conn.Open();
                    SqlCommand cmd = new SqlCommand("rsTemplateXMLTextFileDGV_DelByComName", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                    cmd.Parameters.Add(new SqlParameter("@SunComponent", cbItems.Text.Substring(0, cbItems.Text.LastIndexOf(","))));
                    cmd.Parameters.Add(new SqlParameter("@SunMethod", cbItems.Text.Substring(cbItems.Text.LastIndexOf(",") + 1)));
                    if (sender != null)
                    {
                        rdr = cmd.ExecuteReader();
                        rdr.Close();
                    }
                    finallistCTF.Clear();
                    SqlCommand cmd2 = new SqlCommand("rsTemplateXMLTextFileDGV_Ins", conn);
                    cmd2.CommandType = CommandType.StoredProcedure;
                    for (int i = 0; i < this.dgvCreateTextFile.Rows.Count; i++)
                    {
                        string isEmptyStr = string.Empty;
                        for (int k = 0; k < dgvCreateTextFile.Rows[i].Cells.Count; k++)
                        {
                            isEmptyStr += dgvCreateTextFile.Rows[i].Cells[k].EditedFormattedValue;
                        }
                        if (string.IsNullOrEmpty(isEmptyStr)) continue;
                        RowCreateTextFile re = new RowCreateTextFile();

                        re.ReferenceNumber = this.dgvCreateTextFile.Rows[i].Cells[0].Value == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[0].Value.ToString().Replace(" ", "");
                        re.LineIndicator = this.dgvCreateTextFile.Rows[i].Cells[1].Value == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[1].Value.ToString();
                        re.StartinginCell = this.dgvCreateTextFile.Rows[i].Cells[2].Value == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[2].Value.ToString();

                        re.Column1 = this.dgvCreateTextFile.Columns.Count > 3 ? (this.dgvCreateTextFile.Rows[i].Cells[3].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[3].EditedFormattedValue.ToString()) : "";
                        re.Column2 = this.dgvCreateTextFile.Columns.Count > 4 ? (this.dgvCreateTextFile.Rows[i].Cells[4].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[4].EditedFormattedValue.ToString()) : "";
                        re.Column3 = this.dgvCreateTextFile.Columns.Count > 5 ? (this.dgvCreateTextFile.Rows[i].Cells[5].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[5].EditedFormattedValue.ToString()) : "";
                        re.Column4 = this.dgvCreateTextFile.Columns.Count > 6 ? (this.dgvCreateTextFile.Rows[i].Cells[6].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[6].EditedFormattedValue.ToString()) : "";
                        re.Column5 = this.dgvCreateTextFile.Columns.Count > 7 ? (this.dgvCreateTextFile.Rows[i].Cells[7].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[7].EditedFormattedValue.ToString()) : "";
                        re.Column6 = this.dgvCreateTextFile.Columns.Count > 8 ? (this.dgvCreateTextFile.Rows[i].Cells[8].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[8].EditedFormattedValue.ToString()) : "";
                        re.Column7 = this.dgvCreateTextFile.Columns.Count > 9 ? (this.dgvCreateTextFile.Rows[i].Cells[9].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[9].EditedFormattedValue.ToString()) : "";
                        re.Column8 = this.dgvCreateTextFile.Columns.Count > 10 ? (this.dgvCreateTextFile.Rows[i].Cells[10].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[10].EditedFormattedValue.ToString()) : "";
                        re.Column9 = this.dgvCreateTextFile.Columns.Count > 11 ? (this.dgvCreateTextFile.Rows[i].Cells[11].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[11].EditedFormattedValue.ToString()) : "";
                        re.Column10 = this.dgvCreateTextFile.Columns.Count > 12 ? (this.dgvCreateTextFile.Rows[i].Cells[12].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[12].EditedFormattedValue.ToString()) : "";
                        re.Column11 = this.dgvCreateTextFile.Columns.Count > 13 ? (this.dgvCreateTextFile.Rows[i].Cells[13].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[13].EditedFormattedValue.ToString()) : "";
                        re.Column12 = this.dgvCreateTextFile.Columns.Count > 14 ? (this.dgvCreateTextFile.Rows[i].Cells[14].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[14].EditedFormattedValue.ToString()) : "";
                        re.Column13 = this.dgvCreateTextFile.Columns.Count > 15 ? (this.dgvCreateTextFile.Rows[i].Cells[15].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[15].EditedFormattedValue.ToString()) : "";
                        re.Column14 = this.dgvCreateTextFile.Columns.Count > 16 ? (this.dgvCreateTextFile.Rows[i].Cells[16].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[16].EditedFormattedValue.ToString()) : "";
                        re.Column15 = this.dgvCreateTextFile.Columns.Count > 17 ? (this.dgvCreateTextFile.Rows[i].Cells[17].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[17].EditedFormattedValue.ToString()) : "";
                        re.Column16 = this.dgvCreateTextFile.Columns.Count > 18 ? (this.dgvCreateTextFile.Rows[i].Cells[18].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[18].EditedFormattedValue.ToString()) : "";
                        re.Column17 = this.dgvCreateTextFile.Columns.Count > 19 ? (this.dgvCreateTextFile.Rows[i].Cells[19].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[19].EditedFormattedValue.ToString()) : "";
                        re.Column18 = this.dgvCreateTextFile.Columns.Count > 20 ? (this.dgvCreateTextFile.Rows[i].Cells[20].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[20].EditedFormattedValue.ToString()) : "";
                        re.Column19 = this.dgvCreateTextFile.Columns.Count > 21 ? (this.dgvCreateTextFile.Rows[i].Cells[21].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[21].EditedFormattedValue.ToString()) : "";
                        re.Column20 = this.dgvCreateTextFile.Columns.Count > 22 ? (this.dgvCreateTextFile.Rows[i].Cells[22].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[22].EditedFormattedValue.ToString()) : "";
                        re.Column21 = this.dgvCreateTextFile.Columns.Count > 23 ? (this.dgvCreateTextFile.Rows[i].Cells[23].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[23].EditedFormattedValue.ToString()) : "";
                        re.Column22 = this.dgvCreateTextFile.Columns.Count > 24 ? (this.dgvCreateTextFile.Rows[i].Cells[24].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[24].EditedFormattedValue.ToString()) : "";
                        re.Column23 = this.dgvCreateTextFile.Columns.Count > 25 ? (this.dgvCreateTextFile.Rows[i].Cells[25].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[25].EditedFormattedValue.ToString()) : "";
                        re.Column24 = this.dgvCreateTextFile.Columns.Count > 26 ? (this.dgvCreateTextFile.Rows[i].Cells[26].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[26].EditedFormattedValue.ToString()) : "";
                        re.Column25 = this.dgvCreateTextFile.Columns.Count > 27 ? (this.dgvCreateTextFile.Rows[i].Cells[27].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[27].EditedFormattedValue.ToString()) : "";
                        re.Column26 = this.dgvCreateTextFile.Columns.Count > 28 ? (this.dgvCreateTextFile.Rows[i].Cells[28].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[28].EditedFormattedValue.ToString()) : "";
                        re.Column27 = this.dgvCreateTextFile.Columns.Count > 29 ? (this.dgvCreateTextFile.Rows[i].Cells[29].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[29].EditedFormattedValue.ToString()) : "";
                        re.Column28 = this.dgvCreateTextFile.Columns.Count > 30 ? (this.dgvCreateTextFile.Rows[i].Cells[30].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[30].EditedFormattedValue.ToString()) : "";
                        re.Column29 = this.dgvCreateTextFile.Columns.Count > 31 ? (this.dgvCreateTextFile.Rows[i].Cells[31].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[31].EditedFormattedValue.ToString()) : "";
                        re.Column30 = this.dgvCreateTextFile.Columns.Count > 32 ? (this.dgvCreateTextFile.Rows[i].Cells[32].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[32].EditedFormattedValue.ToString()) : "";
                        re.Column31 = this.dgvCreateTextFile.Columns.Count > 33 ? (this.dgvCreateTextFile.Rows[i].Cells[33].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[33].EditedFormattedValue.ToString()) : "";
                        re.Column32 = this.dgvCreateTextFile.Columns.Count > 34 ? (this.dgvCreateTextFile.Rows[i].Cells[34].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[34].EditedFormattedValue.ToString()) : "";
                        re.Column33 = this.dgvCreateTextFile.Columns.Count > 35 ? (this.dgvCreateTextFile.Rows[i].Cells[35].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[35].EditedFormattedValue.ToString()) : "";
                        re.Column34 = this.dgvCreateTextFile.Columns.Count > 36 ? (this.dgvCreateTextFile.Rows[i].Cells[36].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[36].EditedFormattedValue.ToString()) : "";
                        re.Column35 = this.dgvCreateTextFile.Columns.Count > 37 ? (this.dgvCreateTextFile.Rows[i].Cells[37].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[37].EditedFormattedValue.ToString()) : "";
                        re.Column36 = this.dgvCreateTextFile.Columns.Count > 38 ? (this.dgvCreateTextFile.Rows[i].Cells[38].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[38].EditedFormattedValue.ToString()) : "";
                        re.Column37 = this.dgvCreateTextFile.Columns.Count > 39 ? (this.dgvCreateTextFile.Rows[i].Cells[39].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[39].EditedFormattedValue.ToString()) : "";
                        re.Column38 = this.dgvCreateTextFile.Columns.Count > 40 ? (this.dgvCreateTextFile.Rows[i].Cells[40].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[40].EditedFormattedValue.ToString()) : "";
                        re.Column39 = this.dgvCreateTextFile.Columns.Count > 41 ? (this.dgvCreateTextFile.Rows[i].Cells[41].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[41].EditedFormattedValue.ToString()) : "";
                        re.Column40 = this.dgvCreateTextFile.Columns.Count > 42 ? (this.dgvCreateTextFile.Rows[i].Cells[42].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[42].EditedFormattedValue.ToString()) : "";
                        re.Column41 = this.dgvCreateTextFile.Columns.Count > 43 ? (this.dgvCreateTextFile.Rows[i].Cells[43].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[43].EditedFormattedValue.ToString()) : "";
                        re.Column42 = this.dgvCreateTextFile.Columns.Count > 44 ? (this.dgvCreateTextFile.Rows[i].Cells[44].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[44].EditedFormattedValue.ToString()) : "";
                        re.Column43 = this.dgvCreateTextFile.Columns.Count > 45 ? (this.dgvCreateTextFile.Rows[i].Cells[45].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[45].EditedFormattedValue.ToString()) : "";
                        re.Column44 = this.dgvCreateTextFile.Columns.Count > 46 ? (this.dgvCreateTextFile.Rows[i].Cells[46].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[46].EditedFormattedValue.ToString()) : "";
                        re.Column45 = this.dgvCreateTextFile.Columns.Count > 47 ? (this.dgvCreateTextFile.Rows[i].Cells[47].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[47].EditedFormattedValue.ToString()) : "";
                        re.Column46 = this.dgvCreateTextFile.Columns.Count > 48 ? (this.dgvCreateTextFile.Rows[i].Cells[48].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[48].EditedFormattedValue.ToString()) : "";
                        re.Column47 = this.dgvCreateTextFile.Columns.Count > 49 ? (this.dgvCreateTextFile.Rows[i].Cells[49].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[49].EditedFormattedValue.ToString()) : "";
                        re.Column48 = this.dgvCreateTextFile.Columns.Count > 50 ? (this.dgvCreateTextFile.Rows[i].Cells[50].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[50].EditedFormattedValue.ToString()) : "";
                        re.Column49 = this.dgvCreateTextFile.Columns.Count > 51 ? (this.dgvCreateTextFile.Rows[i].Cells[51].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[51].EditedFormattedValue.ToString()) : "";
                        re.Column50 = this.dgvCreateTextFile.Columns.Count > 52 ? (this.dgvCreateTextFile.Rows[i].Cells[52].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[52].EditedFormattedValue.ToString()) : "";

                        if (string.IsNullOrEmpty(re.LineIndicator))
                            this.dgvCreateTextFile.Rows[i].Cells[1].ErrorText = "Not null.";
                        else
                            this.dgvCreateTextFile.Rows[i].Cells[1].ErrorText = string.Empty;

                        cmd2.Parameters.Add(new SqlParameter("@LineIndicator", re.LineIndicator));
                        cmd2.Parameters.Add(new SqlParameter("@Column1", re.Column1));
                        cmd2.Parameters.Add(new SqlParameter("@Column2", re.Column2));
                        cmd2.Parameters.Add(new SqlParameter("@Column3", re.Column3));
                        cmd2.Parameters.Add(new SqlParameter("@Column4", re.Column4));
                        cmd2.Parameters.Add(new SqlParameter("@Column5", re.Column5));
                        cmd2.Parameters.Add(new SqlParameter("@Column6", re.Column6));
                        cmd2.Parameters.Add(new SqlParameter("@Column7", re.Column7));
                        cmd2.Parameters.Add(new SqlParameter("@Column8", re.Column8));
                        cmd2.Parameters.Add(new SqlParameter("@Column9", re.Column9));
                        cmd2.Parameters.Add(new SqlParameter("@Column10", re.Column10));

                        cmd2.Parameters.Add(new SqlParameter("@Column11", re.Column11));
                        cmd2.Parameters.Add(new SqlParameter("@Column12", re.Column12));
                        cmd2.Parameters.Add(new SqlParameter("@Column13", re.Column13));
                        cmd2.Parameters.Add(new SqlParameter("@Column14", re.Column14));
                        cmd2.Parameters.Add(new SqlParameter("@Column15", re.Column15));
                        cmd2.Parameters.Add(new SqlParameter("@Column16", re.Column16));
                        cmd2.Parameters.Add(new SqlParameter("@Column17", re.Column17));
                        cmd2.Parameters.Add(new SqlParameter("@Column18", re.Column18));
                        cmd2.Parameters.Add(new SqlParameter("@Column19", re.Column19));
                        cmd2.Parameters.Add(new SqlParameter("@Column20", re.Column20));

                        cmd2.Parameters.Add(new SqlParameter("@Column21", re.Column21));
                        cmd2.Parameters.Add(new SqlParameter("@Column22", re.Column22));
                        cmd2.Parameters.Add(new SqlParameter("@Column23", re.Column23));
                        cmd2.Parameters.Add(new SqlParameter("@Column24", re.Column24));
                        cmd2.Parameters.Add(new SqlParameter("@Column25", re.Column25));
                        cmd2.Parameters.Add(new SqlParameter("@Column26", re.Column26));
                        cmd2.Parameters.Add(new SqlParameter("@Column27", re.Column27));
                        cmd2.Parameters.Add(new SqlParameter("@Column28", re.Column28));
                        cmd2.Parameters.Add(new SqlParameter("@Column29", re.Column29));
                        cmd2.Parameters.Add(new SqlParameter("@Column30", re.Column30));

                        cmd2.Parameters.Add(new SqlParameter("@Column31", re.Column31));
                        cmd2.Parameters.Add(new SqlParameter("@Column32", re.Column32));
                        cmd2.Parameters.Add(new SqlParameter("@Column33", re.Column33));
                        cmd2.Parameters.Add(new SqlParameter("@Column34", re.Column34));
                        cmd2.Parameters.Add(new SqlParameter("@Column35", re.Column35));
                        cmd2.Parameters.Add(new SqlParameter("@Column36", re.Column36));
                        cmd2.Parameters.Add(new SqlParameter("@Column37", re.Column37));
                        cmd2.Parameters.Add(new SqlParameter("@Column38", re.Column38));
                        cmd2.Parameters.Add(new SqlParameter("@Column39", re.Column39));
                        cmd2.Parameters.Add(new SqlParameter("@Column40", re.Column40));

                        cmd2.Parameters.Add(new SqlParameter("@Column41", re.Column41));
                        cmd2.Parameters.Add(new SqlParameter("@Column42", re.Column42));
                        cmd2.Parameters.Add(new SqlParameter("@Column43", re.Column43));
                        cmd2.Parameters.Add(new SqlParameter("@Column44", re.Column44));
                        cmd2.Parameters.Add(new SqlParameter("@Column45", re.Column45));
                        cmd2.Parameters.Add(new SqlParameter("@Column46", re.Column46));
                        cmd2.Parameters.Add(new SqlParameter("@Column47", re.Column47));
                        cmd2.Parameters.Add(new SqlParameter("@Column48", re.Column48));
                        cmd2.Parameters.Add(new SqlParameter("@Column49", re.Column49));
                        cmd2.Parameters.Add(new SqlParameter("@Column50", re.Column50));

                        cmd2.Parameters.Add(new SqlParameter("@HeaderTextes", HeaderText));
                        cmd2.Parameters.Add(new SqlParameter("@StartinginCell", re.StartinginCell));
                        cmd2.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                        cmd2.Parameters.Add(new SqlParameter("@IncludeHeaderRow", re.IncludeHeaderRow == "True" ? true : false));
                        cmd2.Parameters.Add(new SqlParameter("@SavePath", ""));
                        cmd2.Parameters.Add(new SqlParameter("@SaveName", ""));
                        cmd2.Parameters.Add(new SqlParameter("@ReferenceNumber", re.ReferenceNumber));
                        cmd2.Parameters.Add(new SqlParameter("@SunComponent", cbItems.Text.Substring(0, cbItems.Text.LastIndexOf(","))));
                        cmd2.Parameters.Add(new SqlParameter("@SunMethod", cbItems.Text.Substring(cbItems.Text.LastIndexOf(",") + 1)));
                        cmd2.Parameters.Add(new SqlParameter("@HeaderValue", HeaderValue));
                        cmd2.Parameters.Add(new SqlParameter("@ProcessName", ""));

                        re.SunComponent = cbItems.Text.Substring(0, cbItems.Text.LastIndexOf(","));
                        re.SunMethod = cbItems.Text.Substring(cbItems.Text.LastIndexOf(",") + 1);
                        if (sender != null)
                        {
                            rdr = cmd2.ExecuteReader();
                            rdr.Close();
                        }
                        cmd2.Parameters.Clear();

                        if (!string.IsNullOrEmpty(re.LineIndicator) && !string.IsNullOrEmpty(re.StartinginCell))
                        {
                            finallistCTF.Add(re);
                            LineIndicatorList.Add(re.LineIndicator);
                            startInCellList.Add(re.StartinginCell);
                        }
                    }
                    this.Invalidate();
                    //if (LineIndicatorList.Count == 0)
                    //{
                    //    throw new Exception("No data found for specified Line Indicator(s)!");
                    //}
                    if (sender == null)
                        AddCreateTextFileEntityListIntoFinalList(finallistCTF, LineIndicatorList, startInCellList, ws, ref finallistCTF);
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message + " - Data Error in XML/Text File tab, Output settings !");
                }
                finally
                {
                    if (conn != null)
                    {
                        conn.Close();
                    }
                    if (rdr != null)
                    {
                        rdr.Close();
                    }
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="ws"></param>
        public void SaveCTF(object sender, Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            string HeaderText = string.Empty;
            string HeaderValue = string.Empty;
            List<string> LineIndicatorList = new List<string>();
            List<string> startInCellList = new List<string>();
            for (int k = 0; k < dgvCreateTextFile.Columns.Count; k++)
            {
                HeaderText += dgvCreateTextFile.Columns[k].HeaderText + ",";
                HeaderValue += dgvCreateTextFile.Columns[k].Name + ",";
            }
            if (!string.IsNullOrEmpty(SessionInfo.UserInfo.File_ftid))
            {
                SqlConnection conn = null;
                SqlDataReader rdr = null;
                try
                {
                    conn = new
                        SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                    conn.Open();
                    SqlCommand cmd = new SqlCommand("rsTemplateXMLTextFileDGV_Del", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                    cmd.Parameters.Add(new SqlParameter("@ProcessName", cbItems.Text));
                    rdr = cmd.ExecuteReader();
                    rdr.Close();
                    finallistCTF.Clear();
                    SqlCommand cmd2 = new SqlCommand("rsTemplateXMLTextFileDGV_Ins", conn);
                    cmd2.CommandType = CommandType.StoredProcedure;
                    for (int i = 0; i < this.dgvCreateTextFile.Rows.Count; i++)
                    {
                        string isEmptyStr = string.Empty;
                        for (int k = 0; k < dgvCreateTextFile.Rows[i].Cells.Count; k++)
                        {
                            isEmptyStr += dgvCreateTextFile.Rows[i].Cells[k].EditedFormattedValue;
                        }
                        if (string.IsNullOrEmpty(isEmptyStr)) continue;
                        RowCreateTextFile re = new RowCreateTextFile();

                        re.ReferenceNumber = this.dgvCreateTextFile.Rows[i].Cells[0].Value == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[0].Value.ToString().Replace(" ", "");
                        re.LineIndicator = this.dgvCreateTextFile.Rows[i].Cells[1].Value == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[1].Value.ToString();
                        re.StartinginCell = this.dgvCreateTextFile.Rows[i].Cells[2].Value == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[2].Value.ToString();
                        re.SavePath = this.dgvCreateTextFile.Rows[i].Cells[3].Value == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[3].Value.ToString();
                        re.SaveName = this.dgvCreateTextFile.Rows[i].Cells[4].Value == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[4].Value.ToString();
                        re.IncludeHeaderRow = this.dgvCreateTextFile.Rows[i].Cells[5].Value == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[5].Value.ToString();

                        re.Column1 = this.dgvCreateTextFile.Columns.Count > 6 ? (this.dgvCreateTextFile.Rows[i].Cells[6].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[6].EditedFormattedValue.ToString()) : "";
                        re.Column2 = this.dgvCreateTextFile.Columns.Count > 7 ? (this.dgvCreateTextFile.Rows[i].Cells[7].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[7].EditedFormattedValue.ToString()) : "";
                        re.Column3 = this.dgvCreateTextFile.Columns.Count > 8 ? (this.dgvCreateTextFile.Rows[i].Cells[8].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[8].EditedFormattedValue.ToString()) : "";
                        re.Column4 = this.dgvCreateTextFile.Columns.Count > 9 ? (this.dgvCreateTextFile.Rows[i].Cells[9].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[9].EditedFormattedValue.ToString()) : "";
                        re.Column5 = this.dgvCreateTextFile.Columns.Count > 10 ? (this.dgvCreateTextFile.Rows[i].Cells[10].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[10].EditedFormattedValue.ToString()) : "";
                        re.Column6 = this.dgvCreateTextFile.Columns.Count > 11 ? (this.dgvCreateTextFile.Rows[i].Cells[11].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[11].EditedFormattedValue.ToString()) : "";
                        re.Column7 = this.dgvCreateTextFile.Columns.Count > 12 ? (this.dgvCreateTextFile.Rows[i].Cells[12].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[12].EditedFormattedValue.ToString()) : "";
                        re.Column8 = this.dgvCreateTextFile.Columns.Count > 13 ? (this.dgvCreateTextFile.Rows[i].Cells[13].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[13].EditedFormattedValue.ToString()) : "";
                        re.Column9 = this.dgvCreateTextFile.Columns.Count > 14 ? (this.dgvCreateTextFile.Rows[i].Cells[14].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[14].EditedFormattedValue.ToString()) : "";
                        re.Column10 = this.dgvCreateTextFile.Columns.Count > 15 ? (this.dgvCreateTextFile.Rows[i].Cells[15].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[15].EditedFormattedValue.ToString()) : "";
                        re.Column11 = this.dgvCreateTextFile.Columns.Count > 16 ? (this.dgvCreateTextFile.Rows[i].Cells[16].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[16].EditedFormattedValue.ToString()) : "";
                        re.Column12 = this.dgvCreateTextFile.Columns.Count > 17 ? (this.dgvCreateTextFile.Rows[i].Cells[17].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[17].EditedFormattedValue.ToString()) : "";
                        re.Column13 = this.dgvCreateTextFile.Columns.Count > 18 ? (this.dgvCreateTextFile.Rows[i].Cells[18].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[18].EditedFormattedValue.ToString()) : "";
                        re.Column14 = this.dgvCreateTextFile.Columns.Count > 19 ? (this.dgvCreateTextFile.Rows[i].Cells[19].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[19].EditedFormattedValue.ToString()) : "";
                        re.Column15 = this.dgvCreateTextFile.Columns.Count > 20 ? (this.dgvCreateTextFile.Rows[i].Cells[20].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[20].EditedFormattedValue.ToString()) : "";
                        re.Column16 = this.dgvCreateTextFile.Columns.Count > 21 ? (this.dgvCreateTextFile.Rows[i].Cells[21].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[21].EditedFormattedValue.ToString()) : "";
                        re.Column17 = this.dgvCreateTextFile.Columns.Count > 22 ? (this.dgvCreateTextFile.Rows[i].Cells[22].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[22].EditedFormattedValue.ToString()) : "";
                        re.Column18 = this.dgvCreateTextFile.Columns.Count > 23 ? (this.dgvCreateTextFile.Rows[i].Cells[23].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[23].EditedFormattedValue.ToString()) : "";
                        re.Column19 = this.dgvCreateTextFile.Columns.Count > 24 ? (this.dgvCreateTextFile.Rows[i].Cells[24].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[24].EditedFormattedValue.ToString()) : "";
                        re.Column20 = this.dgvCreateTextFile.Columns.Count > 25 ? (this.dgvCreateTextFile.Rows[i].Cells[25].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[25].EditedFormattedValue.ToString()) : "";
                        re.Column21 = this.dgvCreateTextFile.Columns.Count > 26 ? (this.dgvCreateTextFile.Rows[i].Cells[26].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[26].EditedFormattedValue.ToString()) : "";
                        re.Column22 = this.dgvCreateTextFile.Columns.Count > 27 ? (this.dgvCreateTextFile.Rows[i].Cells[27].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[27].EditedFormattedValue.ToString()) : "";
                        re.Column23 = this.dgvCreateTextFile.Columns.Count > 28 ? (this.dgvCreateTextFile.Rows[i].Cells[28].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[28].EditedFormattedValue.ToString()) : "";
                        re.Column24 = this.dgvCreateTextFile.Columns.Count > 29 ? (this.dgvCreateTextFile.Rows[i].Cells[29].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[29].EditedFormattedValue.ToString()) : "";
                        re.Column25 = this.dgvCreateTextFile.Columns.Count > 30 ? (this.dgvCreateTextFile.Rows[i].Cells[30].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[30].EditedFormattedValue.ToString()) : "";
                        re.Column26 = this.dgvCreateTextFile.Columns.Count > 31 ? (this.dgvCreateTextFile.Rows[i].Cells[31].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[31].EditedFormattedValue.ToString()) : "";
                        re.Column27 = this.dgvCreateTextFile.Columns.Count > 32 ? (this.dgvCreateTextFile.Rows[i].Cells[32].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[32].EditedFormattedValue.ToString()) : "";
                        re.Column28 = this.dgvCreateTextFile.Columns.Count > 33 ? (this.dgvCreateTextFile.Rows[i].Cells[33].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[33].EditedFormattedValue.ToString()) : "";
                        re.Column29 = this.dgvCreateTextFile.Columns.Count > 34 ? (this.dgvCreateTextFile.Rows[i].Cells[34].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[34].EditedFormattedValue.ToString()) : "";
                        re.Column30 = this.dgvCreateTextFile.Columns.Count > 35 ? (this.dgvCreateTextFile.Rows[i].Cells[35].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[35].EditedFormattedValue.ToString()) : "";
                        re.Column31 = this.dgvCreateTextFile.Columns.Count > 36 ? (this.dgvCreateTextFile.Rows[i].Cells[36].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[36].EditedFormattedValue.ToString()) : "";
                        re.Column32 = this.dgvCreateTextFile.Columns.Count > 37 ? (this.dgvCreateTextFile.Rows[i].Cells[37].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[37].EditedFormattedValue.ToString()) : "";
                        re.Column33 = this.dgvCreateTextFile.Columns.Count > 38 ? (this.dgvCreateTextFile.Rows[i].Cells[38].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[38].EditedFormattedValue.ToString()) : "";
                        re.Column34 = this.dgvCreateTextFile.Columns.Count > 39 ? (this.dgvCreateTextFile.Rows[i].Cells[39].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[39].EditedFormattedValue.ToString()) : "";
                        re.Column35 = this.dgvCreateTextFile.Columns.Count > 40 ? (this.dgvCreateTextFile.Rows[i].Cells[40].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[40].EditedFormattedValue.ToString()) : "";
                        re.Column36 = this.dgvCreateTextFile.Columns.Count > 41 ? (this.dgvCreateTextFile.Rows[i].Cells[41].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[41].EditedFormattedValue.ToString()) : "";
                        re.Column37 = this.dgvCreateTextFile.Columns.Count > 42 ? (this.dgvCreateTextFile.Rows[i].Cells[42].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[42].EditedFormattedValue.ToString()) : "";
                        re.Column38 = this.dgvCreateTextFile.Columns.Count > 43 ? (this.dgvCreateTextFile.Rows[i].Cells[43].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[43].EditedFormattedValue.ToString()) : "";
                        re.Column39 = this.dgvCreateTextFile.Columns.Count > 44 ? (this.dgvCreateTextFile.Rows[i].Cells[44].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[44].EditedFormattedValue.ToString()) : "";
                        re.Column40 = this.dgvCreateTextFile.Columns.Count > 45 ? (this.dgvCreateTextFile.Rows[i].Cells[45].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[45].EditedFormattedValue.ToString()) : "";
                        re.Column41 = this.dgvCreateTextFile.Columns.Count > 46 ? (this.dgvCreateTextFile.Rows[i].Cells[46].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[46].EditedFormattedValue.ToString()) : "";
                        re.Column42 = this.dgvCreateTextFile.Columns.Count > 47 ? (this.dgvCreateTextFile.Rows[i].Cells[47].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[47].EditedFormattedValue.ToString()) : "";
                        re.Column43 = this.dgvCreateTextFile.Columns.Count > 48 ? (this.dgvCreateTextFile.Rows[i].Cells[48].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[48].EditedFormattedValue.ToString()) : "";
                        re.Column44 = this.dgvCreateTextFile.Columns.Count > 49 ? (this.dgvCreateTextFile.Rows[i].Cells[49].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[49].EditedFormattedValue.ToString()) : "";
                        re.Column45 = this.dgvCreateTextFile.Columns.Count > 50 ? (this.dgvCreateTextFile.Rows[i].Cells[50].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[50].EditedFormattedValue.ToString()) : "";
                        re.Column46 = this.dgvCreateTextFile.Columns.Count > 51 ? (this.dgvCreateTextFile.Rows[i].Cells[51].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[51].EditedFormattedValue.ToString()) : "";
                        re.Column47 = this.dgvCreateTextFile.Columns.Count > 52 ? (this.dgvCreateTextFile.Rows[i].Cells[52].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[52].EditedFormattedValue.ToString()) : "";
                        re.Column48 = this.dgvCreateTextFile.Columns.Count > 53 ? (this.dgvCreateTextFile.Rows[i].Cells[53].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[53].EditedFormattedValue.ToString()) : "";
                        re.Column49 = this.dgvCreateTextFile.Columns.Count > 54 ? (this.dgvCreateTextFile.Rows[i].Cells[54].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[54].EditedFormattedValue.ToString()) : "";
                        re.Column50 = this.dgvCreateTextFile.Columns.Count > 55 ? (this.dgvCreateTextFile.Rows[i].Cells[55].EditedFormattedValue == null ? "" : this.dgvCreateTextFile.Rows[i].Cells[55].EditedFormattedValue.ToString()) : "";

                        if (string.IsNullOrEmpty(re.LineIndicator))
                            this.dgvCreateTextFile.Rows[i].Cells[1].ErrorText = "Not null.";
                        else
                            this.dgvCreateTextFile.Rows[i].Cells[1].ErrorText = string.Empty;

                        cmd2.Parameters.Add(new SqlParameter("@LineIndicator", re.LineIndicator));
                        cmd2.Parameters.Add(new SqlParameter("@Column1", re.Column1));
                        cmd2.Parameters.Add(new SqlParameter("@Column2", re.Column2));
                        cmd2.Parameters.Add(new SqlParameter("@Column3", re.Column3));
                        cmd2.Parameters.Add(new SqlParameter("@Column4", re.Column4));
                        cmd2.Parameters.Add(new SqlParameter("@Column5", re.Column5));
                        cmd2.Parameters.Add(new SqlParameter("@Column6", re.Column6));
                        cmd2.Parameters.Add(new SqlParameter("@Column7", re.Column7));
                        cmd2.Parameters.Add(new SqlParameter("@Column8", re.Column8));
                        cmd2.Parameters.Add(new SqlParameter("@Column9", re.Column9));
                        cmd2.Parameters.Add(new SqlParameter("@Column10", re.Column10));

                        cmd2.Parameters.Add(new SqlParameter("@Column11", re.Column11));
                        cmd2.Parameters.Add(new SqlParameter("@Column12", re.Column12));
                        cmd2.Parameters.Add(new SqlParameter("@Column13", re.Column13));
                        cmd2.Parameters.Add(new SqlParameter("@Column14", re.Column14));
                        cmd2.Parameters.Add(new SqlParameter("@Column15", re.Column15));
                        cmd2.Parameters.Add(new SqlParameter("@Column16", re.Column16));
                        cmd2.Parameters.Add(new SqlParameter("@Column17", re.Column17));
                        cmd2.Parameters.Add(new SqlParameter("@Column18", re.Column18));
                        cmd2.Parameters.Add(new SqlParameter("@Column19", re.Column19));
                        cmd2.Parameters.Add(new SqlParameter("@Column20", re.Column20));

                        cmd2.Parameters.Add(new SqlParameter("@Column21", re.Column21));
                        cmd2.Parameters.Add(new SqlParameter("@Column22", re.Column22));
                        cmd2.Parameters.Add(new SqlParameter("@Column23", re.Column23));
                        cmd2.Parameters.Add(new SqlParameter("@Column24", re.Column24));
                        cmd2.Parameters.Add(new SqlParameter("@Column25", re.Column25));
                        cmd2.Parameters.Add(new SqlParameter("@Column26", re.Column26));
                        cmd2.Parameters.Add(new SqlParameter("@Column27", re.Column27));
                        cmd2.Parameters.Add(new SqlParameter("@Column28", re.Column28));
                        cmd2.Parameters.Add(new SqlParameter("@Column29", re.Column29));
                        cmd2.Parameters.Add(new SqlParameter("@Column30", re.Column30));

                        cmd2.Parameters.Add(new SqlParameter("@Column31", re.Column31));
                        cmd2.Parameters.Add(new SqlParameter("@Column32", re.Column32));
                        cmd2.Parameters.Add(new SqlParameter("@Column33", re.Column33));
                        cmd2.Parameters.Add(new SqlParameter("@Column34", re.Column34));
                        cmd2.Parameters.Add(new SqlParameter("@Column35", re.Column35));
                        cmd2.Parameters.Add(new SqlParameter("@Column36", re.Column36));
                        cmd2.Parameters.Add(new SqlParameter("@Column37", re.Column37));
                        cmd2.Parameters.Add(new SqlParameter("@Column38", re.Column38));
                        cmd2.Parameters.Add(new SqlParameter("@Column39", re.Column39));
                        cmd2.Parameters.Add(new SqlParameter("@Column40", re.Column40));

                        cmd2.Parameters.Add(new SqlParameter("@Column41", re.Column41));
                        cmd2.Parameters.Add(new SqlParameter("@Column42", re.Column42));
                        cmd2.Parameters.Add(new SqlParameter("@Column43", re.Column43));
                        cmd2.Parameters.Add(new SqlParameter("@Column44", re.Column44));
                        cmd2.Parameters.Add(new SqlParameter("@Column45", re.Column45));
                        cmd2.Parameters.Add(new SqlParameter("@Column46", re.Column46));
                        cmd2.Parameters.Add(new SqlParameter("@Column47", re.Column47));
                        cmd2.Parameters.Add(new SqlParameter("@Column48", re.Column48));
                        cmd2.Parameters.Add(new SqlParameter("@Column49", re.Column49));
                        cmd2.Parameters.Add(new SqlParameter("@Column50", re.Column50));

                        cmd2.Parameters.Add(new SqlParameter("@HeaderTextes", HeaderText));
                        cmd2.Parameters.Add(new SqlParameter("@StartinginCell", re.StartinginCell));
                        cmd2.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                        cmd2.Parameters.Add(new SqlParameter("@IncludeHeaderRow", re.IncludeHeaderRow == "True" ? true : false));
                        cmd2.Parameters.Add(new SqlParameter("@SavePath", re.SavePath));
                        cmd2.Parameters.Add(new SqlParameter("@SaveName", re.SaveName));
                        cmd2.Parameters.Add(new SqlParameter("@ReferenceNumber", re.ReferenceNumber));
                        cmd2.Parameters.Add(new SqlParameter("@SunComponent", ""));
                        cmd2.Parameters.Add(new SqlParameter("@SunMethod", ""));
                        cmd2.Parameters.Add(new SqlParameter("@HeaderValue", HeaderValue));
                        cmd2.Parameters.Add(new SqlParameter("@ProcessName", cbItems.Text));

                        rdr = cmd2.ExecuteReader();
                        rdr.Close();
                        cmd2.Parameters.Clear();
                        if (!string.IsNullOrEmpty(re.LineIndicator) && !string.IsNullOrEmpty(re.StartinginCell))
                        {
                            finallistCTF.Add(re);
                            LineIndicatorList.Add(re.LineIndicator);
                            startInCellList.Add(re.StartinginCell);
                        }
                    }
                    this.Invalidate();
                    //if (LineIndicatorList.Count == 0)
                    //{
                    //    throw new Exception("No data found for specified Line Indicator(s)!");
                    //}
                    if (sender == null)
                        AddCreateTextFileEntityListIntoFinalList(finallistCTF, LineIndicatorList, startInCellList, ws, ref finallistCTF);
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message + " - Data Error in XML/Text File tab, Output settings !");
                }
                finally
                {
                    if (conn != null)
                    {
                        conn.Close();
                    }
                    if (rdr != null)
                    {
                        rdr.Close();
                    }
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTestTransUpd_Click(object sender, EventArgs e)
        {
            if (Ribbon2.tupf != null)
                Ribbon2.tupf.Dispose();
            DateTime starttime = DateTime.Now;
            Ribbon2.tupf = new TransUpdPostFrm();
            Ribbon2.tupf.panel3.Visible = true;
            Ribbon2.tupf.panel1.Visible = false;
            Ribbon2.tupf.panel2.Visible = false;
            Ribbon2.tupf.panel3.Dock = DockStyle.Fill;
            Ribbon2.tupf.ControlBox = false;
            OutputContainer.isTransUpdFlag = true;
            try
            {
                Globals.ThisAddIn.Application.DisplayAlerts = false;
                Ribbon2.wsRrigin = (Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                var lastColumn = Ribbon2.wsRrigin.Cells.Find("*", Ribbon2.wsRrigin.Cells[1, 1], Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, Microsoft.Office.Interop.Excel.XlLookAt.xlPart, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, Type.Missing, Type.Missing);
                Ribbon2.LastColumnName = Finance_Tools.RemoveNumber(lastColumn.Address).Replace("$", "");
                Finance_Tools.MaxColumnCount = lastColumn.Column;
                var lastrow = Ribbon2.wsRrigin.Cells.Find("*", Ribbon2.wsRrigin.Cells[1, 1], Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, Microsoft.Office.Interop.Excel.XlLookAt.xlPart, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, Type.Missing, Type.Missing);
                Ribbon2.LastRowNumber = Finance_Tools.RemoveNotNumber(lastrow.Address);
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                Ribbon2.tupf.Show();
                TransUpdPostFrm.richTextBox1.Text += "Error List :\r\n";
                SaveTransUpd(null, Ribbon2.wsRrigin);
                SetSession();
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                Globals.ThisAddIn.Application.DisplayAlerts = true;
                if (TransUpdPostFrm.richTextBox1.Text.Length > 21)
                {
                    LogHelper.WriteLog(typeof(Ribbon2), TransUpdPostFrm.richTextBox1.Text + " - Journal update processing error , Template:" + SessionInfo.UserInfo.FileName);
                }
            }
            catch (Exception ex)
            {
                if (ex.ToString().Contains("Clipboard"))
                {
                    MessageBox.Show("Clipboard not ready, please try again.", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(ex.ToString(), "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                LogHelper.WriteLog(typeof(OutputContainer), ex.Message + "Journal update error");
                Ribbon2.tupf.Dispose();
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                Clipboard.SetText("\r\n");
            }
            finally
            {
                Ribbon2.tupf.Focus();
                Ribbon2.tupf.bddata();
                Ribbon2.tupf.panel3.Visible = false;
                Ribbon2.tupf.panel1.Visible = true;
                Ribbon2.tupf.panel2.Visible = true;
                Ribbon2.tupf.ControlBox = true;
                DateTime stoptime = DateTime.Now;
                string costtime = Finance_Tools.DateDiff(starttime, stoptime);
                Globals.ThisAddIn.Application.StatusBar = "latest journal update process costs " + costtime;
                GC.Collect();
                OutputContainer.isTransUpdFlag = false;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTestCTF_Click(object sender, EventArgs e)
        {
            if (dgvCreateTextFile.SelectedRows.Count == 0)
            {
                MessageBox.Show("Please click the row header number ( before Ref column ) to choose a certain Reference number and enjoy your test ! ", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            else
            {
                if (dgvCreateTextFile.SelectedRows[0].Cells[0].Value != null)
                    SessionInfo.UserInfo.CurrentRef = dgvCreateTextFile.SelectedRows[0].Cells[0].Value.ToString().Replace(" ", "");
            }
            DateTime starttime = DateTime.Now;
            if (Ribbon2.ctff != null)
                Ribbon2.ctff.Dispose();
            Ribbon2.ctff = new CreateTextFileForm();
            Ribbon2.ctff.panel3.Visible = true;
            Ribbon2.ctff.panel1.Visible = false;
            Ribbon2.ctff.panel2.Visible = false;
            Ribbon2.ctff.panel3.Dock = DockStyle.Fill;
            Ribbon2.ctff.ControlBox = false;
            try
            {
                Globals.ThisAddIn.Application.DisplayAlerts = false;
                Ribbon2.wsRrigin = (Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                var lastColumn = Ribbon2.wsRrigin.Cells.Find("*", Ribbon2.wsRrigin.Cells[1, 1], Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, Microsoft.Office.Interop.Excel.XlLookAt.xlPart, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, Type.Missing, Type.Missing);
                Ribbon2.LastColumnName = Finance_Tools.RemoveNumber(lastColumn.Address).Replace("$", "");
                Finance_Tools.MaxColumnCount = lastColumn.Column;
                var lastrow = Ribbon2.wsRrigin.Cells.Find("*", Ribbon2.wsRrigin.Cells[1, 1], Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, Microsoft.Office.Interop.Excel.XlLookAt.xlPart, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, Type.Missing, Type.Missing);
                Ribbon2.LastRowNumber = Finance_Tools.RemoveNotNumber(lastrow.Address);
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                CreateTextFileForm.richTextBox1.Text += "Error List :\r\n";
                if (cbXMLOrText.Text == "XML")
                {
                    Ribbon2.ctff.Show();
                    SessionInfo.UserInfo.ComName = cbItems.Text.Substring(0, cbItems.Text.LastIndexOf(","));
                    SessionInfo.UserInfo.MethodName = cbItems.Text.Substring(cbItems.Text.LastIndexOf(",") + 1);
                    SaveXML(null, Ribbon2.wsRrigin);
                }
                else
                {
                    SaveCTF(null, Ribbon2.wsRrigin);
                }
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                Globals.ThisAddIn.Application.DisplayAlerts = true;
                if (CreateTextFileForm.richTextBox1.Text.Length > 21)
                {
                    LogHelper.WriteLog(typeof(Ribbon2), CreateTextFileForm.richTextBox1.Text + " - Create Text File Processing error , Template:" + SessionInfo.UserInfo.FileName);
                }
                if (cbXMLOrText.Text != "XML" && (finallistCTF.Count != 0))
                {
                    string fileName = "";
                    string filepath = "";
                    bool includeHeaderRow = false;
                    if (dgvCreateTextFile.SelectedRows[0].Cells[4].Value != null)
                        fileName = dgvCreateTextFile.SelectedRows[0].Cells[4].Value.ToString().Replace(" ", "");

                    if (dgvCreateTextFile.SelectedRows[0].Cells[3].Value != null)
                        filepath = dgvCreateTextFile.SelectedRows[0].Cells[3].Value.ToString().Replace(" ", "");

                    if (dgvCreateTextFile.SelectedRows[0].Cells[5].Value != null)
                        includeHeaderRow = dgvCreateTextFile.SelectedRows[0].Cells[5].Value.ToString().Replace(" ", "") == "True" ? true : false;
                    GenTextFile(sender, fileName, filepath, includeHeaderRow);
                    return;
                }
            }
            catch (Exception ex)
            {
                if (ex.ToString().Contains("Clipboard"))
                {
                    MessageBox.Show("Clipboard not ready, please try again.", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(ex.ToString(), "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                LogHelper.WriteLog(typeof(Ribbon2), ex.Message + "Create Text File error");
                Ribbon2.ctff.Dispose();
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                Clipboard.SetText("\r\n");
            }
            finally
            {
                Ribbon2.ctff.Focus();
                Ribbon2.ctff.bddata();
                Ribbon2.ctff.panel3.Visible = false;
                Ribbon2.ctff.panel1.Visible = true;
                Ribbon2.ctff.panel2.Visible = true;
                Ribbon2.ctff.ControlBox = true;
                DateTime stoptime = DateTime.Now;
                string costtime = Finance_Tools.DateDiff(starttime, stoptime);
                Globals.ThisAddIn.Application.StatusBar = "latest create text file process costs " + costtime;
                GC.Collect();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="fileName"></param>
        /// <param name="filepath"></param>
        /// <param name="includeHeaderRow"></param>
        public void GenTextFile(object sender, string fileName, string filepath, bool includeHeaderRow)
        {
            try
            {
                bool error = Ribbon2.wsRrigin.Cells.get_Range(fileName).Errors.Item[1].Value;
                if (error || string.IsNullOrEmpty(fileName))
                { }
                else
                    fileName = Ribbon2.wsRrigin.Cells.get_Range(fileName).Value.ToString();
            }
            catch { }

            if (File.Exists(filepath + "\\" + fileName + ".txt"))
            {
                File.Delete(filepath + "\\" + fileName + ".txt");
            }
            FileStream aFile = new FileStream(filepath + "\\" + fileName + ".txt", FileMode.OpenOrCreate, FileAccess.Write, FileShare.ReadWrite);//txtSavePath.Text
            StreamWriter sw = new StreamWriter(aFile);
            int columnCount = dgvCreateTextFile.Columns.Count - 6;
            if (includeHeaderRow)
            {
                sw.WriteLine("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}{12}{13}{14}{15}{16}{17}{18}{19}{20}{21}{22}{23}{24}{25}{26}{27}{28}{29}{30}{31}{32}{33}{34}{35}{36}{37}{38}{39}{40}{41}{42}{43}{44}{45}{46}{47}{48}{49}{50}{51}{52}{53}{54}{55}{56}{57}{58}{59}{60}{61}{62}{63}{64}{65}{66}{67}{68}{69}{70}{71}{72}{73}{74}{75}{76}{77}{78}{79}{80}{81}{82}{83}{84}{85}{86}{87}{88}{89}{90}{91}{92}{93}{94}{95}{96}{97}{98}{99}", "  ", columnCount > 0 ? dgvCreateTextFile.Columns[6].HeaderText : "", "  ", columnCount > 1 ? dgvCreateTextFile.Columns[7].HeaderText : "",
                    "  ", columnCount > 2 ? dgvCreateTextFile.Columns[8].HeaderText : "", "  ", columnCount > 3 ? dgvCreateTextFile.Columns[9].HeaderText : "",
                    "  ", columnCount > 4 ? dgvCreateTextFile.Columns[10].HeaderText : "", "  ", columnCount > 5 ? dgvCreateTextFile.Columns[11].HeaderText : "",
                    "  ", columnCount > 6 ? dgvCreateTextFile.Columns[12].HeaderText : "", "  ", columnCount > 7 ? dgvCreateTextFile.Columns[13].HeaderText : "",
                    "  ", columnCount > 8 ? dgvCreateTextFile.Columns[14].HeaderText : "", "  ", columnCount > 9 ? dgvCreateTextFile.Columns[15].HeaderText : "",
                    "  ", columnCount > 10 ? dgvCreateTextFile.Columns[16].HeaderText : "", "  ", columnCount > 11 ? dgvCreateTextFile.Columns[17].HeaderText : "",
                    "  ", columnCount > 12 ? dgvCreateTextFile.Columns[18].HeaderText : "", "  ", columnCount > 13 ? dgvCreateTextFile.Columns[19].HeaderText : "",
                    "  ", columnCount > 14 ? dgvCreateTextFile.Columns[20].HeaderText : "", "  ", columnCount > 15 ? dgvCreateTextFile.Columns[21].HeaderText : "",
                    "  ", columnCount > 16 ? dgvCreateTextFile.Columns[22].HeaderText : "", "  ", columnCount > 17 ? dgvCreateTextFile.Columns[23].HeaderText : "",
                    "  ", columnCount > 18 ? dgvCreateTextFile.Columns[24].HeaderText : "", "  ", columnCount > 19 ? dgvCreateTextFile.Columns[25].HeaderText : "",
                    "  ", columnCount > 20 ? dgvCreateTextFile.Columns[26].HeaderText : "", "  ", columnCount > 21 ? dgvCreateTextFile.Columns[27].HeaderText : "",
                    "  ", columnCount > 22 ? dgvCreateTextFile.Columns[28].HeaderText : "", "  ", columnCount > 23 ? dgvCreateTextFile.Columns[29].HeaderText : "",
                    "  ", columnCount > 24 ? dgvCreateTextFile.Columns[30].HeaderText : "", "  ", columnCount > 25 ? dgvCreateTextFile.Columns[31].HeaderText : "",
                    "  ", columnCount > 26 ? dgvCreateTextFile.Columns[32].HeaderText : "", "  ", columnCount > 27 ? dgvCreateTextFile.Columns[33].HeaderText : "",
                    "  ", columnCount > 28 ? dgvCreateTextFile.Columns[34].HeaderText : "", "  ", columnCount > 29 ? dgvCreateTextFile.Columns[35].HeaderText : "",
                    "  ", columnCount > 30 ? dgvCreateTextFile.Columns[36].HeaderText : "", "  ", columnCount > 31 ? dgvCreateTextFile.Columns[37].HeaderText : "",
                    "  ", columnCount > 32 ? dgvCreateTextFile.Columns[38].HeaderText : "", "  ", columnCount > 33 ? dgvCreateTextFile.Columns[39].HeaderText : "",
                    "  ", columnCount > 34 ? dgvCreateTextFile.Columns[40].HeaderText : "", "  ", columnCount > 35 ? dgvCreateTextFile.Columns[41].HeaderText : "",
                    "  ", columnCount > 36 ? dgvCreateTextFile.Columns[42].HeaderText : "", "  ", columnCount > 37 ? dgvCreateTextFile.Columns[43].HeaderText : "",
                    "  ", columnCount > 38 ? dgvCreateTextFile.Columns[44].HeaderText : "", "  ", columnCount > 39 ? dgvCreateTextFile.Columns[45].HeaderText : "",
                    "  ", columnCount > 40 ? dgvCreateTextFile.Columns[46].HeaderText : "", "  ", columnCount > 41 ? dgvCreateTextFile.Columns[47].HeaderText : "",
                    "  ", columnCount > 42 ? dgvCreateTextFile.Columns[48].HeaderText : "", "  ", columnCount > 43 ? dgvCreateTextFile.Columns[49].HeaderText : "",
                    "  ", columnCount > 44 ? dgvCreateTextFile.Columns[50].HeaderText : "", "  ", columnCount > 45 ? dgvCreateTextFile.Columns[51].HeaderText : "",
                    "  ", columnCount > 46 ? dgvCreateTextFile.Columns[52].HeaderText : "", "  ", columnCount > 47 ? dgvCreateTextFile.Columns[53].HeaderText : "",
                    "  ", columnCount > 48 ? dgvCreateTextFile.Columns[54].HeaderText : "", "  ", columnCount > 49 ? dgvCreateTextFile.Columns[55].HeaderText : ""
                    );
            }
            DataTable dtFile = ft.GetXMLorTextFileFieldsByFileName(cbItems.Text);
            for (int i = 0; i < finallistCTF.Count; i++)
            {
                string str = ft.getFormatStr(finallistCTF[i].Column1, dtFile, 0) + ft.getSeprStr(dtFile, 0)
                    + ft.getFormatStr(finallistCTF[i].Column2, dtFile, 1) + ft.getSeprStr(dtFile, 1)
                    + ft.getFormatStr(finallistCTF[i].Column3, dtFile, 2) + ft.getSeprStr(dtFile, 2)
                    + ft.getFormatStr(finallistCTF[i].Column4, dtFile, 3) + ft.getSeprStr(dtFile, 3)
                    + ft.getFormatStr(finallistCTF[i].Column5, dtFile, 4) + ft.getSeprStr(dtFile, 4)
                    + ft.getFormatStr(finallistCTF[i].Column6, dtFile, 5) + ft.getSeprStr(dtFile, 5)
                    + ft.getFormatStr(finallistCTF[i].Column7, dtFile, 6) + ft.getSeprStr(dtFile, 6)
                    + ft.getFormatStr(finallistCTF[i].Column8, dtFile, 7) + ft.getSeprStr(dtFile, 7)
                    + ft.getFormatStr(finallistCTF[i].Column9, dtFile, 8) + ft.getSeprStr(dtFile, 8)
                    + ft.getFormatStr(finallistCTF[i].Column10, dtFile, 9) + ft.getSeprStr(dtFile, 9)
                    + ft.getFormatStr(finallistCTF[i].Column11, dtFile, 10) + ft.getSeprStr(dtFile, 10)
                    + ft.getFormatStr(finallistCTF[i].Column12, dtFile, 11) + ft.getSeprStr(dtFile, 11)
                    + ft.getFormatStr(finallistCTF[i].Column13, dtFile, 12) + ft.getSeprStr(dtFile, 12)
                    + ft.getFormatStr(finallistCTF[i].Column14, dtFile, 13) + ft.getSeprStr(dtFile, 13)
                    + ft.getFormatStr(finallistCTF[i].Column15, dtFile, 14) + ft.getSeprStr(dtFile, 14)
                    + ft.getFormatStr(finallistCTF[i].Column16, dtFile, 15) + ft.getSeprStr(dtFile, 15)
                    + ft.getFormatStr(finallistCTF[i].Column17, dtFile, 16) + ft.getSeprStr(dtFile, 16)
                    + ft.getFormatStr(finallistCTF[i].Column18, dtFile, 17) + ft.getSeprStr(dtFile, 17)
                    + ft.getFormatStr(finallistCTF[i].Column19, dtFile, 18) + ft.getSeprStr(dtFile, 18)
                    + ft.getFormatStr(finallistCTF[i].Column20, dtFile, 19) + ft.getSeprStr(dtFile, 19)
                    + ft.getFormatStr(finallistCTF[i].Column21, dtFile, 20) + ft.getSeprStr(dtFile, 20)
                    + ft.getFormatStr(finallistCTF[i].Column22, dtFile, 21) + ft.getSeprStr(dtFile, 21)
                    + ft.getFormatStr(finallistCTF[i].Column23, dtFile, 22) + ft.getSeprStr(dtFile, 22)
                    + ft.getFormatStr(finallistCTF[i].Column24, dtFile, 23) + ft.getSeprStr(dtFile, 23)
                    + ft.getFormatStr(finallistCTF[i].Column25, dtFile, 24) + ft.getSeprStr(dtFile, 24)
                    + ft.getFormatStr(finallistCTF[i].Column26, dtFile, 25) + ft.getSeprStr(dtFile, 25)
                    + ft.getFormatStr(finallistCTF[i].Column27, dtFile, 26) + ft.getSeprStr(dtFile, 26)
                    + ft.getFormatStr(finallistCTF[i].Column28, dtFile, 27) + ft.getSeprStr(dtFile, 27)
                    + ft.getFormatStr(finallistCTF[i].Column29, dtFile, 28) + ft.getSeprStr(dtFile, 28)
                    + ft.getFormatStr(finallistCTF[i].Column30, dtFile, 29) + ft.getSeprStr(dtFile, 29)
                    + ft.getFormatStr(finallistCTF[i].Column31, dtFile, 30) + ft.getSeprStr(dtFile, 30)
                    + ft.getFormatStr(finallistCTF[i].Column32, dtFile, 31) + ft.getSeprStr(dtFile, 31)
                    + ft.getFormatStr(finallistCTF[i].Column33, dtFile, 32) + ft.getSeprStr(dtFile, 32)
                    + ft.getFormatStr(finallistCTF[i].Column34, dtFile, 33) + ft.getSeprStr(dtFile, 33)
                    + ft.getFormatStr(finallistCTF[i].Column35, dtFile, 34) + ft.getSeprStr(dtFile, 34)
                    + ft.getFormatStr(finallistCTF[i].Column36, dtFile, 35) + ft.getSeprStr(dtFile, 35)
                    + ft.getFormatStr(finallistCTF[i].Column37, dtFile, 36) + ft.getSeprStr(dtFile, 36)
                    + ft.getFormatStr(finallistCTF[i].Column38, dtFile, 37) + ft.getSeprStr(dtFile, 37)
                    + ft.getFormatStr(finallistCTF[i].Column39, dtFile, 38) + ft.getSeprStr(dtFile, 38)
                    + ft.getFormatStr(finallistCTF[i].Column40, dtFile, 39) + ft.getSeprStr(dtFile, 39)
                    + ft.getFormatStr(finallistCTF[i].Column41, dtFile, 40) + ft.getSeprStr(dtFile, 40)
                    + ft.getFormatStr(finallistCTF[i].Column42, dtFile, 41) + ft.getSeprStr(dtFile, 41)
                    + ft.getFormatStr(finallistCTF[i].Column43, dtFile, 42) + ft.getSeprStr(dtFile, 42)
                    + ft.getFormatStr(finallistCTF[i].Column44, dtFile, 43) + ft.getSeprStr(dtFile, 43)
                    + ft.getFormatStr(finallistCTF[i].Column45, dtFile, 44) + ft.getSeprStr(dtFile, 44)
                    + ft.getFormatStr(finallistCTF[i].Column46, dtFile, 45) + ft.getSeprStr(dtFile, 45)
                    + ft.getFormatStr(finallistCTF[i].Column47, dtFile, 46) + ft.getSeprStr(dtFile, 46)
                    + ft.getFormatStr(finallistCTF[i].Column48, dtFile, 47) + ft.getSeprStr(dtFile, 47)
                    + ft.getFormatStr(finallistCTF[i].Column49, dtFile, 48) + ft.getSeprStr(dtFile, 48)
                    + ft.getFormatStr(finallistCTF[i].Column50, dtFile, 49) + ft.getSeprStr(dtFile, 49)
                    ;

                sw.WriteLine(str.Replace("\\r", "\r").Replace("\\n", "\n"));
            }
            sw.Close();
            if (sender != null)
                MessageBox.Show("Success!", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                SessionInfo.UserInfo.GlobalError += "Process:" + cbItems.Text + "(" + SessionInfo.UserInfo.CurrentRef + ") - Success! ";
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void cbXMLOrText_SelectedIndexChanged(object sender, EventArgs e)
        {
            cbItems.Items.Clear();
            this.panel17.Controls.Clear();
            if (cbXMLOrText.Text == "XML")
            {
                DataTable dt = ft.GetXMLorTextFileNames(1);
                for (int i = 0; i < dt.Rows.Count; i++)
                    cbItems.Items.Add(dt.Rows[i]["RelatedName"].ToString());
            }
            else
            {
                DataTable dt = ft.GetXMLorTextFileNames(0);
                for (int i = 0; i < dt.Rows.Count; i++)
                    cbItems.Items.Add(dt.Rows[i]["RelatedName"].ToString());
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void cbItems_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbXMLOrText.Text == "XML")
            {
                string comname = cbItems.Text.Substring(0, cbItems.Text.LastIndexOf(","));
                string methodName = cbItems.Text.Substring(cbItems.Text.LastIndexOf(",") + 1);
                DataTable dt = ft.GetXMLorTextFileFieldsByComName(comname, methodName);
                BindCreateTextFileDGV(dt);
            }
            else
            {
                string fileName = cbItems.Text;
                DataTable dt = ft.GetXMLorTextFileFieldsByFileName(fileName);
                BindCreateTextFileDGV(dt);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSetMax_Click(object sender, EventArgs e)
        {
            setMax s = new setMax();
            s.ShowDialog();
        }
    }
}
