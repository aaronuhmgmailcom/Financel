
/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<frmAd_Doc_View>   
 * Author：Peter.uhm  (yanb@shinetechchina.com)
 * Modify date：2016.04
 * Version : 2.0.0.2
 */
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

namespace ExcelAddIn4
{
    public partial class frmAd_Doc_View : Form
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
        internal static RSFinanceToolsEntities db
        {
            get { return new RSFinanceToolsEntities(); }
        }
        public frmAd_Doc_View()
        {
            InitializeComponent();

            cbADV_FType.Items.Add("xlsx");
            cbADV_FType.Items.Add("xlsm");
            cbADV_FType.Items.Add("pdf");
            BindPDFViewer();
        }
        /// <summary>
        /// 
        /// </summary>
        private void BindPDFViewer()
        {
            try
            {
                SessionInfo.UserInfo.Containerpath = (from FT_sett in db.rsTemplateContainers
                                                      where FT_sett.TemplateID == SessionInfo.UserInfo.File_ftid
                                                      select FT_sett.ft_relatefilepath).First();
                string column = (from FT_sett in db.rsTemplateContainers
                                 where FT_sett.TemplateID == SessionInfo.UserInfo.File_ftid
                                 select FT_sett.column).First();

                bool? viewFromDB = (from FT_sett in db.rsTemplateContainers
                                    where FT_sett.TemplateID == SessionInfo.UserInfo.File_ftid
                                    select FT_sett.FromDB).First();

                bool? IsWebBrowser = (from FT_sett in db.rsTemplateContainers
                                      where FT_sett.TemplateID == SessionInfo.UserInfo.File_ftid
                                      select FT_sett.IsWebBrowser).First();

                bool? IsDataBaseQuery = (from FT_sett in db.rsTemplateContainers
                                         where FT_sett.TemplateID == SessionInfo.UserInfo.File_ftid
                                         select FT_sett.IsDataBaseQuery).First();

                string ConnectionString = (from FT_sett in db.rsTemplateContainers
                                           where FT_sett.TemplateID == SessionInfo.UserInfo.File_ftid
                                           select FT_sett.ConnectionString).First();

                string UserID = (from FT_sett in db.rsTemplateContainers
                                 where FT_sett.TemplateID == SessionInfo.UserInfo.File_ftid
                                 select FT_sett.UserID).First();

                string Password = (from FT_sett in db.rsTemplateContainers
                                   where FT_sett.TemplateID == SessionInfo.UserInfo.File_ftid
                                   select FT_sett.Password).First();

                string SQLString = (from FT_sett in db.rsTemplateContainers
                                    where FT_sett.TemplateID == SessionInfo.UserInfo.File_ftid
                                    select FT_sett.SQLString).First();

                txtPath.Text = SessionInfo.UserInfo.Containerpath;
                this.txtInvName.Text = column;
                this.tbColumn2.Text = column;
                this.chkViewFromDB.Checked = (bool)viewFromDB;
                this.rbWb.Checked = (bool)IsWebBrowser;
                this.cbDbQ.Checked = (bool)IsDataBaseQuery;
                this.tbConnectStr.Text = ConnectionString;
                this.tbUserID.Text = UserID;
                this.tbPassword.Text = Password;
                this.tbSQL.Text = SQLString;
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnADV_Folder_Click(object sender, EventArgs e)
        {
            var path = Finance_Tools.RootPath;
            fbdADV.SelectedPath = path;
            DialogResult dr = fbdADV.ShowDialog();
            if (dr == DialogResult.OK)
                tbADV_Filepath.Text = fbdADV.SelectedPath;
        }
        /// <summary>
        /// 
        /// </summary>
        private void InitializeToolTip()
        {
            ToolTip toolTip1 = new ToolTip();
            toolTip1.AutoPopDelay = 5000;
            toolTip1.InitialDelay = 10;
            toolTip1.ReshowDelay = 50;
            toolTip1.ShowAlways = true;
            toolTip1.SetToolTip(this.tbSQL, "~ is replaced by the active cell value each time in the specified column");
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            DialogResult dr = fbdAd_UpdateFolder.ShowDialog();
            if (dr == DialogResult.OK)
                txtPath.Text = fbdAd_UpdateFolder.SelectedPath;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbDbQ_CheckedChanged(object sender, EventArgs e)
        {
            if (cbDbQ.Checked)
            {
                panel21.Visible = false;
                panel22.Visible = true;
                panel21.Location = new Point(6, 219);
                panel22.Location = new Point(6, 98);
                InitializeToolTip();
            }
            else
            {
                panel21.Visible = true;
                panel22.Visible = false;
                panel22.Location = new Point(6, 219);
                panel21.Location = new Point(6, 98);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rbPdf_CheckedChanged(object sender, EventArgs e)
        {
            if (rbPdf.Checked)
            {
                rbWb.Checked = false;
                label4.Text = "Path:";
                btnBrowse.Visible = true;
            }
            else
            {
                rbWb.Checked = true;
                label4.Text = "  Url:";
                btnBrowse.Visible = false;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnADV_Create_Click(object sender, EventArgs e)
        {
            foreach (Control ctrl in this.tabPage1.Controls)
            {
                if ((ctrl is TextBox) && (string.IsNullOrEmpty(ctrl.Text)) && (ctrl.Name != "tbADV_File") && (ctrl.Name != "tbADV_Macro01") && (ctrl.Name != "txtADV_ModuleName"))
                {
                    MessageBox.Show("Please enter a value for the field: " + ctrl.Name, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ctrl.Select();
                    return;
                }
            }
            try
            {
                string Filetype = string.IsNullOrEmpty(cbADV_FType.Text) ? "html" : cbADV_FType.Text;
                ft.sprocVIEW_DOC_INS(tbADV_Type.Text, tbADV_Prefix.Text, tbADV_Filepath.Text,
                                chkADV_UseRef.Checked, tbADV_File.Text, Filetype, (txtADV_ModuleName.Text + "." + tbADV_Macro01.Text).Replace("..", "."));

                this.Close();
                MessageBox.Show("New Document View routine created for " + tbADV_Type.Text, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                MessageBox.Show("Document View could not be created. Please ensure the details you " +
                "entered are correct and try again", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chkADV_UseRef_CheckedChanged(object sender, EventArgs e)
        {
            if (this.chkADV_UseRef.Checked == true)
            {
                this.tbADV_File.Enabled = false;
            }
            else
            {
                this.tbADV_File.Enabled = true;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbADV_FType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbADV_FType.SelectedItem == "xlsm")
            {
                tbADV_Macro01.Enabled = true;
                txtADV_ModuleName.Enabled = true;
            }
            else
            {
                tbADV_Macro01.Enabled = false;
                txtADV_ModuleName.Enabled = false;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ckb_isAWebFile_CheckedChanged(object sender, EventArgs e)
        {
            if (this.ckb_isAWebFile.Checked == true)
            {
                this.btnADV_Folder.Visible = false;
                tbADV_Filepath.Enabled = true;
                cbADV_FType.Visible = false;
                lbADV_FType.Visible = false;
                lbADV_Folder.Text = "Web Path containing document or view document file:";

                lbADV_Macro.Visible = false;
                label1.Visible = false;
                lbADV_Macro01.Visible = false;
                txtADV_ModuleName.Visible = false;
                tbADV_Macro01.Visible = false;
            }
            else
            {
                this.btnADV_Folder.Visible = true;
                tbADV_Filepath.Enabled = false;
                cbADV_FType.Visible = true;
                lbADV_FType.Visible = true;
                lbADV_Folder.Text = "Folder containing document or view document file:";

                lbADV_Macro.Visible = true;
                label1.Visible = true;
                lbADV_Macro01.Visible = true;
                txtADV_ModuleName.Visible = true;
                tbADV_Macro01.Visible = true;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tbConnectStr_TextChanged(object sender, EventArgs e)
        {
            tbConnectStr.ForeColor = Color.Black;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tbSQL_TextChanged(object sender, EventArgs e)
        {
            tbSQL.ForeColor = Color.Black;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(SessionInfo.UserInfo.File_ftid))
            {
                this.errorProvider1.SetError(this.btnSave, "Please choose a template first!"); return;
            }
            string Path = txtPath.Text;
            string ColumnName = string.Empty;
            this.errorProvider1.Clear();
            if ((this.cbDbQ.Checked == true))
            {
                if ((tbConnectStr.Text == "Data Source=myDatabaseName;Initial Catalog=myDBTables;User ID=[UserID];Password=[Password];") || string.IsNullOrEmpty(tbConnectStr.Text))
                {
                    this.errorProvider1.SetError(this.tbConnectStr, "Please enter your Connection String!"); return;
                }
                if ((tbSQL.Text == "Select di.urlname from docInvoice di Where di.invoiceNum = ‘~’") || string.IsNullOrEmpty(tbSQL.Text))
                {
                    this.errorProvider1.SetError(this.tbSQL, "Please enter your SQL String!"); return;
                }
                if (string.IsNullOrEmpty(tbColumn2.Text))
                {
                    this.errorProvider1.SetError(this.tbColumn2, "Please enter the specified column!"); return;
                }
                ColumnName = tbColumn2.Text;
            }
            else
            {
                ColumnName = this.txtInvName.Text;
            }
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsTemplateContainer_Del", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                cmd.ExecuteNonQuery();
                SqlCommand cmd2 = new SqlCommand("rsTemplateContainer_Ins", conn);
                cmd2.CommandType = CommandType.StoredProcedure;
                cmd2.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                cmd2.Parameters.Add(new SqlParameter("@ft_relatefilepath", Path));
                cmd2.Parameters.Add(new SqlParameter("@column", ColumnName));
                cmd2.Parameters.Add(new SqlParameter("@FromDB", chkViewFromDB.Checked));
                cmd2.Parameters.Add(new SqlParameter("@IsWebBrowser", rbWb.Checked));
                cmd2.Parameters.Add(new SqlParameter("@IsDataBaseQuery", cbDbQ.Checked));
                cmd2.Parameters.Add(new SqlParameter("@ConnectionString", tbConnectStr.Text));
                cmd2.Parameters.Add(new SqlParameter("@UserID", tbUserID.Text));
                cmd2.Parameters.Add(new SqlParameter("@Password", tbPassword.Text));
                cmd2.Parameters.Add(new SqlParameter("@SQLString", tbSQL.Text));

                cmd2.ExecuteNonQuery();
                MessageBox.Show("Success!", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                }
            }
        }
    }
}
