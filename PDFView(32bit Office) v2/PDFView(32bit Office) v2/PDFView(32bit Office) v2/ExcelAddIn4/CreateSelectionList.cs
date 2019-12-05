using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;

namespace ExcelAddIn3
{
    public partial class CreateSelectionList : Form
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
        internal static string output
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public CreateSelectionList()
        {
            InitializeComponent();
            cbFolders.Items.Add("Accounts");
            if (string.IsNullOrEmpty(SessionInfo.UserInfo.FilePath)) return;
            cell = Globals.ThisAddIn.Application.ActiveCell.Address;
            var wstmp = (Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            sheet = wstmp.Name;
            if (!string.IsNullOrEmpty(ft.CSLTableName(SessionInfo.UserInfo.FilePath, sheet, cell)))
            {
                cbFolders.SelectedItem = ft.CSLTableName(SessionInfo.UserInfo.FilePath, sheet, cell);
            }
            if (!string.IsNullOrEmpty(ft.CSLColumnName(SessionInfo.UserInfo.FilePath, sheet, cell)))
            {
                cbColumnName.SelectedItem = ft.CSLColumnName(SessionInfo.UserInfo.FilePath, sheet, cell);
            }
            if (!string.IsNullOrEmpty(ft.CSLColumnName2(SessionInfo.UserInfo.FilePath, sheet, cell)))
            {
                cbColumnName2.SelectedItem = ft.CSLColumnName2(SessionInfo.UserInfo.FilePath, sheet, cell);
            }
            if (!string.IsNullOrEmpty(ft.CSLColumnName3(SessionInfo.UserInfo.FilePath, sheet, cell)))
            {
                cbColumnName3.SelectedItem = ft.CSLColumnName3(SessionInfo.UserInfo.FilePath, sheet, cell);
            }
            if (!string.IsNullOrEmpty(ft.CSLFilter(SessionInfo.UserInfo.FilePath, sheet, cell)))
            {
                txtFilter.Text = ft.CSLFilter(SessionInfo.UserInfo.FilePath, sheet, cell);
            }
            if (!string.IsNullOrEmpty(ft.CSLFilter2(SessionInfo.UserInfo.FilePath, sheet, cell)))
            {
                txtFilter2.Text = ft.CSLFilter2(SessionInfo.UserInfo.FilePath, sheet, cell);
            }
            if (!string.IsNullOrEmpty(ft.CSLFilter3(SessionInfo.UserInfo.FilePath, sheet, cell)))
            {
                txtFilter3.Text = ft.CSLFilter3(SessionInfo.UserInfo.FilePath, sheet, cell);
            }
            if (!string.IsNullOrEmpty(ft.CSLOperator(SessionInfo.UserInfo.FilePath, sheet, cell)))
            {
                cbOperator.SelectedItem = ft.CSLOperator(SessionInfo.UserInfo.FilePath, sheet, cell);
            }
            if (!string.IsNullOrEmpty(ft.CSLOperator2(SessionInfo.UserInfo.FilePath, sheet, cell)))
            {
                cbOperator2.SelectedItem = ft.CSLOperator2(SessionInfo.UserInfo.FilePath, sheet, cell);
            }
            if (!string.IsNullOrEmpty(ft.CSLOperator3(SessionInfo.UserInfo.FilePath, sheet, cell)))
            {
                cbOperator3.SelectedItem = ft.CSLOperator3(SessionInfo.UserInfo.FilePath, sheet, cell);
            }
            if (!string.IsNullOrEmpty(ft.CSLOutPut(SessionInfo.UserInfo.FilePath, sheet, cell)))
            {
                output = ft.CSLOutPut(SessionInfo.UserInfo.FilePath, sheet, cell);
            }
            if (output != null && cbColumnName.SelectedItem != null && output == cbColumnName.SelectedItem.ToString()) cbOutPut.SelectedItem = "True";
            else if (output != null && cbColumnName2.SelectedItem != null && output == cbColumnName2.SelectedItem.ToString()) cbOutPut2.SelectedItem = "True";
            else if (output != null && cbColumnName3.SelectedItem != null && output == cbColumnName3.SelectedItem.ToString()) cbOutPut3.SelectedItem = "True";

            InitializeCharacters();
        }
        /// <summary>
        /// 
        /// </summary>
        private void InitializeCharacters()
        {
            Finance_Tools.sCSLTableName = cbFolders.SelectedItem == null ? "" : cbFolders.SelectedItem.ToString();
            Finance_Tools.sCSLColumnName = cbColumnName.SelectedItem == null ? "" : cbColumnName.SelectedItem.ToString();
            Finance_Tools.sCSLColumnName2 = cbColumnName2.SelectedItem == null ? "" : cbColumnName2.SelectedItem.ToString();
            Finance_Tools.sCSLColumnName3 = cbColumnName3.SelectedItem == null ? "" : cbColumnName3.SelectedItem.ToString();
            Finance_Tools.sCSLFilter = txtFilter.Text;
            Finance_Tools.sCSLFilter2 = txtFilter2.Text;
            Finance_Tools.sCSLFilter3 = txtFilter3.Text;
            Finance_Tools.sCSLOperator = cbOperator.SelectedItem == null ? "" : cbOperator.SelectedItem.ToString();
            Finance_Tools.sCSLOperator2 = cbOperator2.SelectedItem == null ? "" : cbOperator2.SelectedItem.ToString();
            Finance_Tools.sCSLOperator3 = cbOperator3.SelectedItem == null ? "" : cbOperator3.SelectedItem.ToString();
            Finance_Tools.sCSLOutPut = output;
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
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void CreateSelectionList_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Visible = false;
            //Ribbon1.CSL_VisibleChanged(null, null);
            DisposeCharacters();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, EventArgs e)
        {
            if (this.cbFolders.SelectedItem == null || this.cbColumnName.SelectedItem == null)
            {
                MessageBox.Show("Table name and column name can't be empty!", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            //if (!(this.cbColumnName.SelectedItem != this.cbColumnName2.SelectedItem && this.cbColumnName.SelectedItem != this.cbColumnName3.SelectedItem && this.cbColumnName3.SelectedItem != this.cbColumnName2.SelectedItem))
            //{
            //    MessageBox.Show("Column Names can't be the same!", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    return;
            //}
            SqlConnection conn = null;
            SqlDataReader rdr = null;

            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("FT_CSL_Delete", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@ft_filepath", SessionInfo.UserInfo.FilePath));
                cmd.Parameters.Add(new SqlParameter("@ft_sheet", sheet));
                cmd.Parameters.Add(new SqlParameter("@ft_cell", cell));
                rdr = cmd.ExecuteReader();

                SqlCommand cmd2 = new SqlCommand("FT_CSL_Ins", conn);
                cmd2.CommandType = CommandType.StoredProcedure;

                cmd2.Parameters.Add(new SqlParameter("@CSLTableName", cbFolders.SelectedItem));
                cmd2.Parameters.Add(new SqlParameter("@CSLColumnName", cbColumnName.SelectedItem));
                cmd2.Parameters.Add(new SqlParameter("@CSLFilter", txtFilter.Text));
                cmd2.Parameters.Add(new SqlParameter("@CSLOperator", cbOperator.SelectedItem == null ? "" : cbOperator.SelectedItem));
                cmd2.Parameters.Add(new SqlParameter("@CSLColumnName2", cbColumnName2.SelectedItem == null ? "" : cbColumnName2.SelectedItem));
                cmd2.Parameters.Add(new SqlParameter("@CSLFilter2", txtFilter2.Text));
                cmd2.Parameters.Add(new SqlParameter("@CSLOperator2", cbOperator2.SelectedItem == null ? "" : cbOperator2.SelectedItem));
                cmd2.Parameters.Add(new SqlParameter("@CSLColumnName3", cbColumnName3.SelectedItem == null ? "" : cbColumnName3.SelectedItem));
                cmd2.Parameters.Add(new SqlParameter("@CSLFilter3", txtFilter3.Text));
                cmd2.Parameters.Add(new SqlParameter("@CSLOperator3", cbOperator3.SelectedItem == null ? "" : cbOperator3.SelectedItem));

                cmd2.Parameters.Add(new SqlParameter("@TemplatePath", SessionInfo.UserInfo.FilePath));
                cmd2.Parameters.Add(new SqlParameter("@SheetName", sheet));
                cmd2.Parameters.Add(new SqlParameter("@CellName", cell));
                if (cbOutPut.SelectedItem == "True") output = cbColumnName.SelectedItem.ToString();
                else if (cbOutPut2.SelectedItem == "True") output = cbColumnName2.SelectedItem.ToString();
                else if (cbOutPut3.SelectedItem == "True") output = cbColumnName3.SelectedItem.ToString();

                cmd2.Parameters.Add(new SqlParameter("@OutPut", output));
                cmd2.Parameters.Add(new SqlParameter("@TemplateName", ""));

                rdr.Close();
                rdr = cmd2.ExecuteReader();

                InitializeCharacters();
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

            this.Visible = false;
            //Ribbon1.CSL_VisibleChanged(null, null);
            DisposeCharacters();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            //Ribbon1.CSL_VisibleChanged(null, null);
            DisposeCharacters();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbFolders_SelectedIndexChanged(object sender, EventArgs e)
        {
            cbOperator.Items.Add("EQU");
            cbOperator.Items.Add("NEQU");
            cbOperator.Items.Add("GT");
            cbOperator.Items.Add("GTE");
            cbOperator.Items.Add("LT");
            cbOperator.Items.Add("LTE");
            cbOperator.Items.Add("BETWEEN");
            cbOperator.Items.Add("IN");
            cbOperator.Items.Add("LIKE");

            cbOperator2.Items.Add("EQU");
            cbOperator2.Items.Add("NEQU");
            cbOperator2.Items.Add("GT");
            cbOperator2.Items.Add("GTE");
            cbOperator2.Items.Add("LT");
            cbOperator2.Items.Add("LTE");
            cbOperator2.Items.Add("BETWEEN");
            cbOperator2.Items.Add("IN");
            cbOperator2.Items.Add("LIKE");

            cbOperator3.Items.Add("EQU");
            cbOperator3.Items.Add("NEQU");
            cbOperator3.Items.Add("GT");
            cbOperator3.Items.Add("GTE");
            cbOperator3.Items.Add("LT");
            cbOperator3.Items.Add("LTE");
            cbOperator3.Items.Add("BETWEEN");
            cbOperator3.Items.Add("IN");
            cbOperator3.Items.Add("LIKE");

            cbOutPut.Items.Add("True");
            cbOutPut.Items.Add("False");

            cbOutPut2.Items.Add("True");
            cbOutPut2.Items.Add("False");

            cbOutPut3.Items.Add("True");
            cbOutPut3.Items.Add("False");

            if (this.cbFolders.SelectedItem.ToString() == "Accounts")
            {
                cbColumnName.Items.Add("AccountCode");
                cbColumnName.Items.Add("AccountType");
                cbColumnName.Items.Add("DataAccessGroupCode");
                cbColumnName.Items.Add("DefaultCurrencyCode");
                cbColumnName.Items.Add("Description");
                cbColumnName.Items.Add("EnterAnalysis1");
                cbColumnName.Items.Add("EnterAnalysis10");
                cbColumnName.Items.Add("EnterAnalysis2");
                cbColumnName.Items.Add("EnterAnalysis3");
                cbColumnName.Items.Add("EnterAnalysis4");
                cbColumnName.Items.Add("EnterAnalysis5");
                cbColumnName.Items.Add("EnterAnalysis6");
                cbColumnName.Items.Add("EnterAnalysis7");
                cbColumnName.Items.Add("EnterAnalysis8");
                cbColumnName.Items.Add("EnterAnalysis9");
                cbColumnName.Items.Add("LongDescription");
                cbColumnName.Items.Add("LookupCode");
                cbColumnName.Items.Add("ShortHeading");
                cbColumnName.Items.Add("Status");
                cbColumnName.Items.Add("SuppressRevaluation");
                cbColumnName.Items.Add("UserArea");

                cbColumnName.Items.Add("Analysis1/VAcntCatAnalysis_AcntCode");
                cbColumnName.Items.Add("Analysis1/VAcntCatAnalysis_AnlCode");
                cbColumnName.Items.Add("Analysis1/VAcntCatAnalysis_DagCode");
                cbColumnName.Items.Add("Analysis1/VAcntCatAnalysis_Descr");
                cbColumnName.Items.Add("Analysis1/VAcntCatAnalysis_Lookup");
                cbColumnName.Items.Add("Analysis1/VAcntCatAnalysis_SHead");
                cbColumnName.Items.Add("Analysis1/VAcntCatAnalysis_Status");

                cbColumnName.Items.Add("Analysis10/VAcntCatAnalysis_AcntCode");
                cbColumnName.Items.Add("Analysis10/VAcntCatAnalysis_AnlCode");
                cbColumnName.Items.Add("Analysis10/VAcntCatAnalysis_DagCode");
                cbColumnName.Items.Add("Analysis10/VAcntCatAnalysis_Descr");
                cbColumnName.Items.Add("Analysis10/VAcntCatAnalysis_Lookup");
                cbColumnName.Items.Add("Analysis10/VAcntCatAnalysis_SHead");
                cbColumnName.Items.Add("Analysis10/VAcntCatAnalysis_Status");

                cbColumnName.Items.Add("Analysis2/VAcntCatAnalysis_AcntCode");
                cbColumnName.Items.Add("Analysis2/VAcntCatAnalysis_AnlCode");
                cbColumnName.Items.Add("Analysis2/VAcntCatAnalysis_DagCode");
                cbColumnName.Items.Add("Analysis2/VAcntCatAnalysis_Descr");
                cbColumnName.Items.Add("Analysis2/VAcntCatAnalysis_Lookup");
                cbColumnName.Items.Add("Analysis2/VAcntCatAnalysis_SHead");
                cbColumnName.Items.Add("Analysis2/VAcntCatAnalysis_Status");

                cbColumnName.Items.Add("Analysis3/VAcntCatAnalysis_AcntCode");
                cbColumnName.Items.Add("Analysis3/VAcntCatAnalysis_AnlCode");
                cbColumnName.Items.Add("Analysis3/VAcntCatAnalysis_DagCode");
                cbColumnName.Items.Add("Analysis3/VAcntCatAnalysis_Descr");
                cbColumnName.Items.Add("Analysis3/VAcntCatAnalysis_Lookup");
                cbColumnName.Items.Add("Analysis3/VAcntCatAnalysis_SHead");
                cbColumnName.Items.Add("Analysis3/VAcntCatAnalysis_Status");

                cbColumnName.Items.Add("Analysis4/VAcntCatAnalysis_AcntCode");
                cbColumnName.Items.Add("Analysis4/VAcntCatAnalysis_AnlCode");
                cbColumnName.Items.Add("Analysis4/VAcntCatAnalysis_DagCode");
                cbColumnName.Items.Add("Analysis4/VAcntCatAnalysis_Descr");
                cbColumnName.Items.Add("Analysis4/VAcntCatAnalysis_Lookup");
                cbColumnName.Items.Add("Analysis4/VAcntCatAnalysis_SHead");
                cbColumnName.Items.Add("Analysis4/VAcntCatAnalysis_Status");

                cbColumnName.Items.Add("Analysis5/VAcntCatAnalysis_AcntCode");
                cbColumnName.Items.Add("Analysis5/VAcntCatAnalysis_AnlCode");
                cbColumnName.Items.Add("Analysis5/VAcntCatAnalysis_DagCode");
                cbColumnName.Items.Add("Analysis5/VAcntCatAnalysis_Descr");
                cbColumnName.Items.Add("Analysis5/VAcntCatAnalysis_Lookup");
                cbColumnName.Items.Add("Analysis5/VAcntCatAnalysis_SHead");
                cbColumnName.Items.Add("Analysis5/VAcntCatAnalysis_Status");

                cbColumnName.Items.Add("Analysis6/VAcntCatAnalysis_AcntCode");
                cbColumnName.Items.Add("Analysis6/VAcntCatAnalysis_AnlCode");
                cbColumnName.Items.Add("Analysis6/VAcntCatAnalysis_DagCode");
                cbColumnName.Items.Add("Analysis6/VAcntCatAnalysis_Descr");
                cbColumnName.Items.Add("Analysis6/VAcntCatAnalysis_Lookup");
                cbColumnName.Items.Add("Analysis6/VAcntCatAnalysis_SHead");
                cbColumnName.Items.Add("Analysis6/VAcntCatAnalysis_Status");

                cbColumnName.Items.Add("Analysis7/VAcntCatAnalysis_AcntCode");
                cbColumnName.Items.Add("Analysis7/VAcntCatAnalysis_AnlCode");
                cbColumnName.Items.Add("Analysis7/VAcntCatAnalysis_DagCode");
                cbColumnName.Items.Add("Analysis7/VAcntCatAnalysis_Descr");
                cbColumnName.Items.Add("Analysis7/VAcntCatAnalysis_Lookup");
                cbColumnName.Items.Add("Analysis7/VAcntCatAnalysis_SHead");
                cbColumnName.Items.Add("Analysis7/VAcntCatAnalysis_Status");

                cbColumnName.Items.Add("Analysis8/VAcntCatAnalysis_AcntCode");
                cbColumnName.Items.Add("Analysis8/VAcntCatAnalysis_AnlCode");
                cbColumnName.Items.Add("Analysis8/VAcntCatAnalysis_DagCode");
                cbColumnName.Items.Add("Analysis8/VAcntCatAnalysis_Descr");
                cbColumnName.Items.Add("Analysis8/VAcntCatAnalysis_Lookup");
                cbColumnName.Items.Add("Analysis8/VAcntCatAnalysis_SHead");
                cbColumnName.Items.Add("Analysis8/VAcntCatAnalysis_Status");

                cbColumnName.Items.Add("Analysis9/VAcntCatAnalysis_AcntCode");
                cbColumnName.Items.Add("Analysis9/VAcntCatAnalysis_AnlCode");
                cbColumnName.Items.Add("Analysis9/VAcntCatAnalysis_DagCode");
                cbColumnName.Items.Add("Analysis9/VAcntCatAnalysis_Descr");
                cbColumnName.Items.Add("Analysis9/VAcntCatAnalysis_Lookup");
                cbColumnName.Items.Add("Analysis9/VAcntCatAnalysis_SHead");
                cbColumnName.Items.Add("Analysis9/VAcntCatAnalysis_Status");

                cbColumnName2.Items.Add("AccountCode");
                cbColumnName2.Items.Add("AccountType");
                cbColumnName2.Items.Add("DataAccessGroupCode");
                cbColumnName2.Items.Add("DefaultCurrencyCode");
                cbColumnName2.Items.Add("Description");
                cbColumnName2.Items.Add("EnterAnalysis1");
                cbColumnName2.Items.Add("EnterAnalysis10");
                cbColumnName2.Items.Add("EnterAnalysis2");
                cbColumnName2.Items.Add("EnterAnalysis3");
                cbColumnName2.Items.Add("EnterAnalysis4");
                cbColumnName2.Items.Add("EnterAnalysis5");
                cbColumnName2.Items.Add("EnterAnalysis6");
                cbColumnName2.Items.Add("EnterAnalysis7");
                cbColumnName2.Items.Add("EnterAnalysis8");
                cbColumnName2.Items.Add("EnterAnalysis9");
                cbColumnName2.Items.Add("LongDescription");
                cbColumnName2.Items.Add("LookupCode");
                cbColumnName2.Items.Add("ShortHeading");
                cbColumnName2.Items.Add("Status");
                cbColumnName2.Items.Add("SuppressRevaluation");
                cbColumnName2.Items.Add("UserArea");

                cbColumnName2.Items.Add("Analysis1/VAcntCatAnalysis_AcntCode");
                cbColumnName2.Items.Add("Analysis1/VAcntCatAnalysis_AnlCode");
                cbColumnName2.Items.Add("Analysis1/VAcntCatAnalysis_DagCode");
                cbColumnName2.Items.Add("Analysis1/VAcntCatAnalysis_Descr");
                cbColumnName2.Items.Add("Analysis1/VAcntCatAnalysis_Lookup");
                cbColumnName2.Items.Add("Analysis1/VAcntCatAnalysis_SHead");
                cbColumnName2.Items.Add("Analysis1/VAcntCatAnalysis_Status");

                cbColumnName2.Items.Add("Analysis10/VAcntCatAnalysis_AcntCode");
                cbColumnName2.Items.Add("Analysis10/VAcntCatAnalysis_AnlCode");
                cbColumnName2.Items.Add("Analysis10/VAcntCatAnalysis_DagCode");
                cbColumnName2.Items.Add("Analysis10/VAcntCatAnalysis_Descr");
                cbColumnName2.Items.Add("Analysis10/VAcntCatAnalysis_Lookup");
                cbColumnName2.Items.Add("Analysis10/VAcntCatAnalysis_SHead");
                cbColumnName2.Items.Add("Analysis10/VAcntCatAnalysis_Status");

                cbColumnName2.Items.Add("Analysis2/VAcntCatAnalysis_AcntCode");
                cbColumnName2.Items.Add("Analysis2/VAcntCatAnalysis_AnlCode");
                cbColumnName2.Items.Add("Analysis2/VAcntCatAnalysis_DagCode");
                cbColumnName2.Items.Add("Analysis2/VAcntCatAnalysis_Descr");
                cbColumnName2.Items.Add("Analysis2/VAcntCatAnalysis_Lookup");
                cbColumnName2.Items.Add("Analysis2/VAcntCatAnalysis_SHead");
                cbColumnName2.Items.Add("Analysis2/VAcntCatAnalysis_Status");

                cbColumnName2.Items.Add("Analysis3/VAcntCatAnalysis_AcntCode");
                cbColumnName2.Items.Add("Analysis3/VAcntCatAnalysis_AnlCode");
                cbColumnName2.Items.Add("Analysis3/VAcntCatAnalysis_DagCode");
                cbColumnName2.Items.Add("Analysis3/VAcntCatAnalysis_Descr");
                cbColumnName2.Items.Add("Analysis3/VAcntCatAnalysis_Lookup");
                cbColumnName2.Items.Add("Analysis3/VAcntCatAnalysis_SHead");
                cbColumnName2.Items.Add("Analysis3/VAcntCatAnalysis_Status");

                cbColumnName2.Items.Add("Analysis4/VAcntCatAnalysis_AcntCode");
                cbColumnName2.Items.Add("Analysis4/VAcntCatAnalysis_AnlCode");
                cbColumnName2.Items.Add("Analysis4/VAcntCatAnalysis_DagCode");
                cbColumnName2.Items.Add("Analysis4/VAcntCatAnalysis_Descr");
                cbColumnName2.Items.Add("Analysis4/VAcntCatAnalysis_Lookup");
                cbColumnName2.Items.Add("Analysis4/VAcntCatAnalysis_SHead");
                cbColumnName2.Items.Add("Analysis4/VAcntCatAnalysis_Status");

                cbColumnName2.Items.Add("Analysis5/VAcntCatAnalysis_AcntCode");
                cbColumnName2.Items.Add("Analysis5/VAcntCatAnalysis_AnlCode");
                cbColumnName2.Items.Add("Analysis5/VAcntCatAnalysis_DagCode");
                cbColumnName2.Items.Add("Analysis5/VAcntCatAnalysis_Descr");
                cbColumnName2.Items.Add("Analysis5/VAcntCatAnalysis_Lookup");
                cbColumnName2.Items.Add("Analysis5/VAcntCatAnalysis_SHead");
                cbColumnName2.Items.Add("Analysis5/VAcntCatAnalysis_Status");

                cbColumnName2.Items.Add("Analysis6/VAcntCatAnalysis_AcntCode");
                cbColumnName2.Items.Add("Analysis6/VAcntCatAnalysis_AnlCode");
                cbColumnName2.Items.Add("Analysis6/VAcntCatAnalysis_DagCode");
                cbColumnName2.Items.Add("Analysis6/VAcntCatAnalysis_Descr");
                cbColumnName2.Items.Add("Analysis6/VAcntCatAnalysis_Lookup");
                cbColumnName2.Items.Add("Analysis6/VAcntCatAnalysis_SHead");
                cbColumnName2.Items.Add("Analysis6/VAcntCatAnalysis_Status");

                cbColumnName2.Items.Add("Analysis7/VAcntCatAnalysis_AcntCode");
                cbColumnName2.Items.Add("Analysis7/VAcntCatAnalysis_AnlCode");
                cbColumnName2.Items.Add("Analysis7/VAcntCatAnalysis_DagCode");
                cbColumnName2.Items.Add("Analysis7/VAcntCatAnalysis_Descr");
                cbColumnName2.Items.Add("Analysis7/VAcntCatAnalysis_Lookup");
                cbColumnName2.Items.Add("Analysis7/VAcntCatAnalysis_SHead");
                cbColumnName2.Items.Add("Analysis7/VAcntCatAnalysis_Status");

                cbColumnName2.Items.Add("Analysis8/VAcntCatAnalysis_AcntCode");
                cbColumnName2.Items.Add("Analysis8/VAcntCatAnalysis_AnlCode");
                cbColumnName2.Items.Add("Analysis8/VAcntCatAnalysis_DagCode");
                cbColumnName2.Items.Add("Analysis8/VAcntCatAnalysis_Descr");
                cbColumnName2.Items.Add("Analysis8/VAcntCatAnalysis_Lookup");
                cbColumnName2.Items.Add("Analysis8/VAcntCatAnalysis_SHead");
                cbColumnName2.Items.Add("Analysis8/VAcntCatAnalysis_Status");

                cbColumnName2.Items.Add("Analysis9/VAcntCatAnalysis_AcntCode");
                cbColumnName2.Items.Add("Analysis9/VAcntCatAnalysis_AnlCode");
                cbColumnName2.Items.Add("Analysis9/VAcntCatAnalysis_DagCode");
                cbColumnName2.Items.Add("Analysis9/VAcntCatAnalysis_Descr");
                cbColumnName2.Items.Add("Analysis9/VAcntCatAnalysis_Lookup");
                cbColumnName2.Items.Add("Analysis9/VAcntCatAnalysis_SHead");
                cbColumnName2.Items.Add("Analysis9/VAcntCatAnalysis_Status");

                cbColumnName3.Items.Add("AccountCode");
                cbColumnName3.Items.Add("AccountType");
                cbColumnName3.Items.Add("DataAccessGroupCode");
                cbColumnName3.Items.Add("DefaultCurrencyCode");
                cbColumnName3.Items.Add("Description");
                cbColumnName3.Items.Add("EnterAnalysis1");
                cbColumnName3.Items.Add("EnterAnalysis10");
                cbColumnName3.Items.Add("EnterAnalysis2");
                cbColumnName3.Items.Add("EnterAnalysis3");
                cbColumnName3.Items.Add("EnterAnalysis4");
                cbColumnName3.Items.Add("EnterAnalysis5");
                cbColumnName3.Items.Add("EnterAnalysis6");
                cbColumnName3.Items.Add("EnterAnalysis7");
                cbColumnName3.Items.Add("EnterAnalysis8");
                cbColumnName3.Items.Add("EnterAnalysis9");
                cbColumnName3.Items.Add("LongDescription");
                cbColumnName3.Items.Add("LookupCode");
                cbColumnName3.Items.Add("ShortHeading");
                cbColumnName3.Items.Add("Status");
                cbColumnName3.Items.Add("SuppressRevaluation");
                cbColumnName3.Items.Add("UserArea");

                cbColumnName3.Items.Add("Analysis1/VAcntCatAnalysis_AcntCode");
                cbColumnName3.Items.Add("Analysis1/VAcntCatAnalysis_AnlCode");
                cbColumnName3.Items.Add("Analysis1/VAcntCatAnalysis_DagCode");
                cbColumnName3.Items.Add("Analysis1/VAcntCatAnalysis_Descr");
                cbColumnName3.Items.Add("Analysis1/VAcntCatAnalysis_Lookup");
                cbColumnName3.Items.Add("Analysis1/VAcntCatAnalysis_SHead");
                cbColumnName3.Items.Add("Analysis1/VAcntCatAnalysis_Status");

                cbColumnName3.Items.Add("Analysis10/VAcntCatAnalysis_AcntCode");
                cbColumnName3.Items.Add("Analysis10/VAcntCatAnalysis_AnlCode");
                cbColumnName3.Items.Add("Analysis10/VAcntCatAnalysis_DagCode");
                cbColumnName3.Items.Add("Analysis10/VAcntCatAnalysis_Descr");
                cbColumnName3.Items.Add("Analysis10/VAcntCatAnalysis_Lookup");
                cbColumnName3.Items.Add("Analysis10/VAcntCatAnalysis_SHead");
                cbColumnName3.Items.Add("Analysis10/VAcntCatAnalysis_Status");

                cbColumnName3.Items.Add("Analysis2/VAcntCatAnalysis_AcntCode");
                cbColumnName3.Items.Add("Analysis2/VAcntCatAnalysis_AnlCode");
                cbColumnName3.Items.Add("Analysis2/VAcntCatAnalysis_DagCode");
                cbColumnName3.Items.Add("Analysis2/VAcntCatAnalysis_Descr");
                cbColumnName3.Items.Add("Analysis2/VAcntCatAnalysis_Lookup");
                cbColumnName3.Items.Add("Analysis2/VAcntCatAnalysis_SHead");
                cbColumnName3.Items.Add("Analysis2/VAcntCatAnalysis_Status");

                cbColumnName3.Items.Add("Analysis3/VAcntCatAnalysis_AcntCode");
                cbColumnName3.Items.Add("Analysis3/VAcntCatAnalysis_AnlCode");
                cbColumnName3.Items.Add("Analysis3/VAcntCatAnalysis_DagCode");
                cbColumnName3.Items.Add("Analysis3/VAcntCatAnalysis_Descr");
                cbColumnName3.Items.Add("Analysis3/VAcntCatAnalysis_Lookup");
                cbColumnName3.Items.Add("Analysis3/VAcntCatAnalysis_SHead");
                cbColumnName3.Items.Add("Analysis3/VAcntCatAnalysis_Status");

                cbColumnName3.Items.Add("Analysis4/VAcntCatAnalysis_AcntCode");
                cbColumnName3.Items.Add("Analysis4/VAcntCatAnalysis_AnlCode");
                cbColumnName3.Items.Add("Analysis4/VAcntCatAnalysis_DagCode");
                cbColumnName3.Items.Add("Analysis4/VAcntCatAnalysis_Descr");
                cbColumnName3.Items.Add("Analysis4/VAcntCatAnalysis_Lookup");
                cbColumnName3.Items.Add("Analysis4/VAcntCatAnalysis_SHead");
                cbColumnName3.Items.Add("Analysis4/VAcntCatAnalysis_Status");

                cbColumnName3.Items.Add("Analysis5/VAcntCatAnalysis_AcntCode");
                cbColumnName3.Items.Add("Analysis5/VAcntCatAnalysis_AnlCode");
                cbColumnName3.Items.Add("Analysis5/VAcntCatAnalysis_DagCode");
                cbColumnName3.Items.Add("Analysis5/VAcntCatAnalysis_Descr");
                cbColumnName3.Items.Add("Analysis5/VAcntCatAnalysis_Lookup");
                cbColumnName3.Items.Add("Analysis5/VAcntCatAnalysis_SHead");
                cbColumnName3.Items.Add("Analysis5/VAcntCatAnalysis_Status");

                cbColumnName3.Items.Add("Analysis6/VAcntCatAnalysis_AcntCode");
                cbColumnName3.Items.Add("Analysis6/VAcntCatAnalysis_AnlCode");
                cbColumnName3.Items.Add("Analysis6/VAcntCatAnalysis_DagCode");
                cbColumnName3.Items.Add("Analysis6/VAcntCatAnalysis_Descr");
                cbColumnName3.Items.Add("Analysis6/VAcntCatAnalysis_Lookup");
                cbColumnName3.Items.Add("Analysis6/VAcntCatAnalysis_SHead");
                cbColumnName3.Items.Add("Analysis6/VAcntCatAnalysis_Status");

                cbColumnName3.Items.Add("Analysis7/VAcntCatAnalysis_AcntCode");
                cbColumnName3.Items.Add("Analysis7/VAcntCatAnalysis_AnlCode");
                cbColumnName3.Items.Add("Analysis7/VAcntCatAnalysis_DagCode");
                cbColumnName3.Items.Add("Analysis7/VAcntCatAnalysis_Descr");
                cbColumnName3.Items.Add("Analysis7/VAcntCatAnalysis_Lookup");
                cbColumnName3.Items.Add("Analysis7/VAcntCatAnalysis_SHead");
                cbColumnName3.Items.Add("Analysis7/VAcntCatAnalysis_Status");

                cbColumnName3.Items.Add("Analysis8/VAcntCatAnalysis_AcntCode");
                cbColumnName3.Items.Add("Analysis8/VAcntCatAnalysis_AnlCode");
                cbColumnName3.Items.Add("Analysis8/VAcntCatAnalysis_DagCode");
                cbColumnName3.Items.Add("Analysis8/VAcntCatAnalysis_Descr");
                cbColumnName3.Items.Add("Analysis8/VAcntCatAnalysis_Lookup");
                cbColumnName3.Items.Add("Analysis8/VAcntCatAnalysis_SHead");
                cbColumnName3.Items.Add("Analysis8/VAcntCatAnalysis_Status");

                cbColumnName3.Items.Add("Analysis9/VAcntCatAnalysis_AcntCode");
                cbColumnName3.Items.Add("Analysis9/VAcntCatAnalysis_AnlCode");
                cbColumnName3.Items.Add("Analysis9/VAcntCatAnalysis_DagCode");
                cbColumnName3.Items.Add("Analysis9/VAcntCatAnalysis_Descr");
                cbColumnName3.Items.Add("Analysis9/VAcntCatAnalysis_Lookup");
                cbColumnName3.Items.Add("Analysis9/VAcntCatAnalysis_SHead");
                cbColumnName3.Items.Add("Analysis9/VAcntCatAnalysis_Status");
            }
        }

        private void cbOutPut_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.cbOutPut.SelectedItem.ToString() == "True")
            {
                cbOutPut2.SelectedItem = "False";
                cbOutPut3.SelectedItem = "False";
            }
        }

        private void cbOutPut2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.cbOutPut2.SelectedItem.ToString() == "True")
            {
                cbOutPut.SelectedItem = "False";
                cbOutPut3.SelectedItem = "False";
            }
        }

        private void cbOutPut3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.cbOutPut3.SelectedItem.ToString() == "True")
            {
                cbOutPut2.SelectedItem = "False";
                cbOutPut.SelectedItem = "False";
            }
        }

    }
}
