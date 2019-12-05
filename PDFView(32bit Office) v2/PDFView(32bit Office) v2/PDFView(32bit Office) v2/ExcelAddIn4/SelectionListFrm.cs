using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace ExcelAddIn3
{
    public partial class SelectionListFrm : Form
    {
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
        internal static string cell
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        internal static string sheet
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgv_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            try
            {
                System.Drawing.Rectangle rectangle = new System.Drawing.Rectangle(e.RowBounds.Location.X,
                    e.RowBounds.Location.Y,
                    dataGridView1.RowHeadersWidth - 4,
                    e.RowBounds.Height);
                TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(),
                    dataGridView1.RowHeadersDefaultCellStyle.Font,
                    rectangle,
                    dataGridView1.RowHeadersDefaultCellStyle.ForeColor,
                    TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
            }
            catch { }
        }
        DataTable dt = new DataTable();
        /// <summary>
        /// 
        /// </summary>
        private void InitializeCharacters()
        {
            if (string.IsNullOrEmpty(SessionInfo.UserInfo.FilePath)) return;
            cell = Globals.ThisAddIn.Application.ActiveCell.Address;
            var wstmp = (Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            sheet = wstmp.Name;
            if (!string.IsNullOrEmpty(ft.CSLTableName(SessionInfo.UserInfo.FilePath, sheet, cell)))
            {
                Finance_Tools.sCSLTableName = ft.CSLTableName(SessionInfo.UserInfo.FilePath, sheet, cell);
            }
            if (!string.IsNullOrEmpty(ft.CSLColumnName(SessionInfo.UserInfo.FilePath, sheet, cell)))
            {
                Finance_Tools.sCSLColumnName = ft.CSLColumnName(SessionInfo.UserInfo.FilePath, sheet, cell);
            }
            if (!string.IsNullOrEmpty(ft.CSLColumnName2(SessionInfo.UserInfo.FilePath, sheet, cell)))
            {
                Finance_Tools.sCSLColumnName2 = ft.CSLColumnName2(SessionInfo.UserInfo.FilePath, sheet, cell);
            }
            if (!string.IsNullOrEmpty(ft.CSLColumnName3(SessionInfo.UserInfo.FilePath, sheet, cell)))
            {
                Finance_Tools.sCSLColumnName3 = ft.CSLColumnName3(SessionInfo.UserInfo.FilePath, sheet, cell);
            }
            if (!string.IsNullOrEmpty(ft.CSLFilter(SessionInfo.UserInfo.FilePath, sheet, cell)))
            {
                Finance_Tools.sCSLFilter = ft.CSLFilter(SessionInfo.UserInfo.FilePath, sheet, cell);
            }
            if (!string.IsNullOrEmpty(ft.CSLFilter2(SessionInfo.UserInfo.FilePath, sheet, cell)))
            {
                Finance_Tools.sCSLFilter2 = ft.CSLFilter2(SessionInfo.UserInfo.FilePath, sheet, cell);
            }
            if (!string.IsNullOrEmpty(ft.CSLFilter3(SessionInfo.UserInfo.FilePath, sheet, cell)))
            {
                Finance_Tools.sCSLFilter3 = ft.CSLFilter3(SessionInfo.UserInfo.FilePath, sheet, cell);
            }
            if (!string.IsNullOrEmpty(ft.CSLOperator(SessionInfo.UserInfo.FilePath, sheet, cell)))
            {
                Finance_Tools.sCSLOperator = ft.CSLOperator(SessionInfo.UserInfo.FilePath, sheet, cell);
            }
            if (!string.IsNullOrEmpty(ft.CSLOperator2(SessionInfo.UserInfo.FilePath, sheet, cell)))
            {
                Finance_Tools.sCSLOperator2 = ft.CSLOperator2(SessionInfo.UserInfo.FilePath, sheet, cell);
            }
            if (!string.IsNullOrEmpty(ft.CSLOperator3(SessionInfo.UserInfo.FilePath, sheet, cell)))
            {
                Finance_Tools.sCSLOperator3 = ft.CSLOperator3(SessionInfo.UserInfo.FilePath, sheet, cell);
            }
            if (!string.IsNullOrEmpty(ft.CSLOutPut(SessionInfo.UserInfo.FilePath, sheet, cell)))
            {
                Finance_Tools.sCSLOutPut = ft.CSLOutPut(SessionInfo.UserInfo.FilePath, sheet, cell);
            }
            Finance_Tools.sCSLTemplatePath = SessionInfo.UserInfo.FilePath;
            Finance_Tools.sCSLSheet = sheet;
            Finance_Tools.sCSLCell = cell;
        }
        /// <summary>
        /// 
        /// </summary>
        public void DisposeCharacters()
        {
            Finance_Tools.sCSLTableName = "";
            Finance_Tools.sCSLColumnName = "";
            Finance_Tools.sCSLColumnName2 = "";
            Finance_Tools.sCSLColumnName3 = "";
            Finance_Tools.sCSLFilter = "";
            Finance_Tools.sCSLFilter2 = "";
            Finance_Tools.sCSLFilter3 = "";
            Finance_Tools.sCSLOperator = "";
            Finance_Tools.sCSLOperator2 = "";
            Finance_Tools.sCSLOperator3 = "";
            Finance_Tools.sCSLOutPut = "";
            Finance_Tools.sCSLTemplatePath = "";
            Finance_Tools.sCSLSheet = "";
            Finance_Tools.sCSLCell = "";
        }
        /// <summary>
        /// 
        /// </summary>
        public SelectionListFrm()
        {
            InitializeComponent();
            dataGridView1.RowPostPaint += new DataGridViewRowPostPaintEventHandler(dgv_RowPostPaint);
            try
            {
                InitializeCharacters();
                string result = Finance_Tools.GetListForm();
                dt.Columns.Add(Finance_Tools.sCSLColumnName + "_c1");
                dt.Columns.Add(Finance_Tools.sCSLColumnName2 + "_c2");
                dt.Columns.Add(Finance_Tools.sCSLColumnName3 + "_c3");

                XmlNodeList nodeList;
                try
                {
                    XmlDocument xdoc = new XmlDocument();
                    xdoc.LoadXml(result);
                    nodeList = xdoc.SelectSingleNode("SSC").ChildNodes;
                }
                catch
                {
                    throw new Exception();
                }
                foreach (XmlNode xn in nodeList)
                {
                    try
                    {
                        XmlElement xe = xn as XmlElement;
                        string Name = xe.Name;
                        if (Name == "Payload")
                        {
                            XmlNodeList nodeListSub = xe.ChildNodes;
                            foreach (XmlNode xnSub in nodeListSub)
                            {
                                XmlElement xeSub = xnSub as XmlElement;
                                string subName = xeSub.Name;
                                if (subName == Finance_Tools.sCSLTableName)
                                {
                                    XmlNodeList nodeListSubSub = xeSub.ChildNodes;
                                    string cn1 = string.Empty;
                                    string cn2 = string.Empty;
                                    string cn3 = string.Empty;
                                    foreach (XmlNode xnSubSub in nodeListSubSub)
                                    {
                                        XmlElement xeSubSub = xnSubSub as XmlElement;
                                        string subsubName = xeSubSub.Name;
                                        if (subsubName == Finance_Tools.sCSLColumnName)
                                        {
                                            cn1 = xeSubSub.InnerText;
                                        }
                                        if (subsubName == Finance_Tools.sCSLColumnName2)
                                        {
                                            cn2 = xeSubSub.InnerText;
                                        }
                                        if (subsubName == Finance_Tools.sCSLColumnName3)
                                        {
                                            cn3 = xeSubSub.InnerText;
                                        }
                                    }
                                    dt.Rows.Add(cn1, cn2, cn3);
                                }
                            }
                        }
                    }
                    catch
                    {
                        throw new Exception();
                    }
                }
                DataTable tmp = GetFilterDataTable(dt);
                dataGridView1.DataSource = tmp;
                label1.Text = tmp.Rows.Count + " records";
                this.Focus();
            }
            catch (Exception ex)
            {
                DisposeCharacters();
                //MessageBox.Show(ex.Message);
                //this.Close();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private DataTable GetFilterDataTable(DataTable dt)
        {
            string filterExpression = string.Empty;
            //if (Finance_Tools.sCSLColumnName != Finance_Tools.sCSLColumnName2 && Finance_Tools.sCSLColumnName != Finance_Tools.sCSLColumnName3 && Finance_Tools.sCSLColumnName2 != Finance_Tools.sCSLColumnName3)
            //{
            //FILTER 1
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter) && Finance_Tools.sCSLOperator == "BETWEEN")
            {
                filterExpression += "";
                string[] sArray = Finance_Tools.sCSLFilter.Split(',');
                string minvalue = sArray[0];
                string maxvalue = sArray[1];
                filterExpression += Finance_Tools.sCSLColumnName + "_c1" + ">= '" + minvalue + "' and  " + Finance_Tools.sCSLColumnName + "_c1" + "<='" + maxvalue + "'";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter) && Finance_Tools.sCSLOperator == "EQU")
            {
                filterExpression += "";
                filterExpression += Finance_Tools.sCSLColumnName + "_c1" + "= '" + Finance_Tools.sCSLFilter + "'";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter) && Finance_Tools.sCSLOperator == "NEQU")
            {
                filterExpression += "";
                filterExpression += Finance_Tools.sCSLColumnName + "_c1" + " <> '" + Finance_Tools.sCSLFilter + "'";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter) && Finance_Tools.sCSLOperator == "GT")
            {
                filterExpression += "";
                filterExpression += Finance_Tools.sCSLColumnName + "_c1" + "> '" + Finance_Tools.sCSLFilter + "'";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter) && Finance_Tools.sCSLOperator == "GTE")
            {
                filterExpression += "";
                filterExpression += Finance_Tools.sCSLColumnName + "_c1" + ">= '" + Finance_Tools.sCSLFilter + "'";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter) && Finance_Tools.sCSLOperator == "LT")
            {
                filterExpression += "";
                filterExpression += Finance_Tools.sCSLColumnName + "_c1" + "< '" + Finance_Tools.sCSLFilter + "'";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter) && Finance_Tools.sCSLOperator == "LTE")
            {
                filterExpression += "";
                filterExpression += Finance_Tools.sCSLColumnName + "_c1" + "<= '" + Finance_Tools.sCSLFilter + "'";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter) && Finance_Tools.sCSLOperator == "IN")
            {
                filterExpression += "";
                filterExpression += Finance_Tools.sCSLColumnName + "_c1" + " in ('" + Finance_Tools.sCSLFilter.Replace(",", "','") + "')";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter) && Finance_Tools.sCSLOperator == "LIKE")
            {
                filterExpression += "";
                filterExpression += Finance_Tools.sCSLColumnName + "_c1" + " like '" + Finance_Tools.sCSLFilter + "'";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter) && !string.IsNullOrEmpty(Finance_Tools.sCSLFilter2)) filterExpression += " AND ";
            //FILTER 2
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter2) && Finance_Tools.sCSLOperator2 == "BETWEEN")
            {
                filterExpression += "";
                string[] sArray = Finance_Tools.sCSLFilter2.Split(',');
                string minvalue = sArray[0];
                string maxvalue = sArray[1];
                filterExpression += Finance_Tools.sCSLColumnName2 + "_c2" + ">= '" + minvalue + "' and  " + Finance_Tools.sCSLColumnName2 + "_c2" + "<='" + maxvalue + "'";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter2) && Finance_Tools.sCSLOperator2 == "EQU")
            {
                filterExpression += "";
                filterExpression += Finance_Tools.sCSLColumnName2 + "_c2" + "= '" + Finance_Tools.sCSLFilter2 + "'";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter2) && Finance_Tools.sCSLOperator2 == "NEQU")
            {
                filterExpression += "";
                filterExpression += Finance_Tools.sCSLColumnName2 + "_c2" + "<> '" + Finance_Tools.sCSLFilter2 + "'";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter2) && Finance_Tools.sCSLOperator2 == "GT")
            {
                filterExpression += "";
                filterExpression += Finance_Tools.sCSLColumnName2 + "_c2" + "> '" + Finance_Tools.sCSLFilter2 + "'";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter2) && Finance_Tools.sCSLOperator2 == "GTE")
            {
                filterExpression += "";
                filterExpression += Finance_Tools.sCSLColumnName2 + "_c2" + ">= '" + Finance_Tools.sCSLFilter2 + "'";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter2) && Finance_Tools.sCSLOperator2 == "LT")
            {
                filterExpression += "";
                filterExpression += Finance_Tools.sCSLColumnName2 + "_c2" + "< '" + Finance_Tools.sCSLFilter2 + "'";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter2) && Finance_Tools.sCSLOperator2 == "LTE")
            {
                filterExpression += "";
                filterExpression += Finance_Tools.sCSLColumnName2 + "_c2" + "<= '" + Finance_Tools.sCSLFilter2 + "'";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter2) && Finance_Tools.sCSLOperator2 == "IN")
            {
                filterExpression += "";
                filterExpression += Finance_Tools.sCSLColumnName2 + "_c2" + " in ('" + Finance_Tools.sCSLFilter2.Replace(",", "','") + "')";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter2) && Finance_Tools.sCSLOperator2 == "LIKE")
            {
                filterExpression += "";
                filterExpression += Finance_Tools.sCSLColumnName2 + "_c2" + " like '" + Finance_Tools.sCSLFilter2 + "'";
            }
            //FILTER 3
            if ((!string.IsNullOrEmpty(Finance_Tools.sCSLFilter) || !string.IsNullOrEmpty(Finance_Tools.sCSLFilter2)) && !string.IsNullOrEmpty(Finance_Tools.sCSLFilter3)) filterExpression += " AND ";
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter3) && Finance_Tools.sCSLOperator3 == "BETWEEN")
            {
                filterExpression += "";
                string[] sArray = Finance_Tools.sCSLFilter3.Split(',');
                string minvalue = sArray[0];
                string maxvalue = sArray[1];
                filterExpression += Finance_Tools.sCSLColumnName3 + "_c3" + ">= '" + minvalue + "' and  " + Finance_Tools.sCSLColumnName3 + "_c3" + "<='" + maxvalue + "'";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter3) && Finance_Tools.sCSLOperator3 == "EQU")
            {
                filterExpression += "";
                filterExpression += Finance_Tools.sCSLColumnName3 + "_c3" + "= '" + Finance_Tools.sCSLFilter3 + "'";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter3) && Finance_Tools.sCSLOperator3 == "NEQU")
            {
                filterExpression += "";
                filterExpression += Finance_Tools.sCSLColumnName3 + "_c3" + "<> '" + Finance_Tools.sCSLFilter3 + "'";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter3) && Finance_Tools.sCSLOperator3 == "GT")
            {
                filterExpression += "";
                filterExpression += Finance_Tools.sCSLColumnName3 + "_c3" + "> '" + Finance_Tools.sCSLFilter3 + "'";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter3) && Finance_Tools.sCSLOperator3 == "GTE")
            {
                filterExpression += "";
                filterExpression += Finance_Tools.sCSLColumnName3 + "_c3" + ">= '" + Finance_Tools.sCSLFilter3 + "'";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter3) && Finance_Tools.sCSLOperator3 == "LT")
            {
                filterExpression += "";
                filterExpression += Finance_Tools.sCSLColumnName3 + "_c3" + "< '" + Finance_Tools.sCSLFilter3 + "'";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter3) && Finance_Tools.sCSLOperator3 == "LTE")
            {
                filterExpression += "";
                filterExpression += Finance_Tools.sCSLColumnName3 + "_c3" + "<= '" + Finance_Tools.sCSLFilter3 + "'";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter3) && Finance_Tools.sCSLOperator3 == "IN")
            {
                filterExpression += "";
                filterExpression += Finance_Tools.sCSLColumnName3 + "_c3" + " in ('" + Finance_Tools.sCSLFilter3.Replace(",", "','") + "')";
            }
            if (!string.IsNullOrEmpty(Finance_Tools.sCSLFilter3) && Finance_Tools.sCSLOperator3 == "LIKE")
            {
                filterExpression += "";
                filterExpression += Finance_Tools.sCSLColumnName3 + "_c3" + " like '" + Finance_Tools.sCSLFilter3 + "'";
            }
            //}
            DataRow[] returnValue = dt.Select(filterExpression);
            DataTable returnDT = dt.Clone();
            foreach (DataRow dr in returnValue)
                returnDT.ImportRow(dr);

            this.label2.Text = Finance_Tools.sCSLTableName + " ：" + filterExpression + " Reference Data:" + dt.Rows.Count.ToString() + "," + returnValue.Length.ToString();

            return returnDT;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, EventArgs e)
        {
            var xlapp = Globals.ThisAddIn.Application;
            int cellIndex = 0;
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                if (dataGridView1.Columns[i].Name.Replace("_c1", "").Replace("_c2", "").Replace("_c3", "") == Finance_Tools.sCSLOutPut)
                    cellIndex = i;
            }
            xlapp.ActiveCell.Value = dataGridView1.CurrentRow.Cells[cellIndex].Value;
            DisposeCharacters();
            this.Close();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            DisposeCharacters();
            this.Close();
        }

    }
}
