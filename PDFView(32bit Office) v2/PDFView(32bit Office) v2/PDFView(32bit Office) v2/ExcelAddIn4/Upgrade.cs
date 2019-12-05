/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<Upgrade>   
 * Author：Peter.uhm  (54778723@qq.com)
 * Modify date：2015.12
 * Modify date：2016.09
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
using ExcelAddIn4.Common;
using System.Data.SqlClient;
using System.Security.Principal;
using System.Configuration;
using System.IO;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddIn4
{
    public partial class Upgrade : Form
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
        ToolTip toolTip1 = new ToolTip();
        /// <summary>
        /// 
        /// </summary>
        private System.Windows.Forms.FolderBrowserDialog fbdAd_UpdateFolder = new FolderBrowserDialog();
        /// <summary>
        /// 
        /// </summary>
        public Upgrade()
        {
            InitializeComponent();
            toolTip1.SetToolTip(txtpath, "Specify the path for the template file from which to import the data. For example, C:\\MyData, \\Sales\\Northwind. Or, click Browse.");
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
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSync_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            if (chkSyncDB.Checked)
            {

            }
            else
            {
                SqlConnection conn = null;
                try
                {
                    conn = new
                        SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                    conn.Open();
                    SqlCommand cmd = null;
                    ImportTemplatesIntoDB(conn, cmd);
                    ImportAddDocumentViewPermission(conn, cmd);
                    ImportAttachPDFPermission(conn, cmd);
                    ImportCreateNewButtonPermission(conn, cmd);
                    ImportAmendCodeDescription(conn, cmd);
                    //ImportPreferencePermission(conn, cmd);
                    ImportCreateNewTemplatePermission(conn, cmd);
                    ImportCreateXMLorTextFileProfile(conn, cmd);
                    ImportUpgradePermission(conn, cmd);
                    ImportSecurityPermission(conn, cmd);
                }
                finally
                {
                    if (conn != null)
                        conn.Close();
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="cmd"></param>
        private void ImportCreateXMLorTextFileProfile(SqlConnection conn, SqlCommand cmd)
        {
            if (!ft.GetPermissionByName("Create XML or Text File Profile"))
            {
                InsertIntoPermission("Create XML or Text File Profile", conn, cmd);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="cmd"></param>
        private void ImportSecurityPermission(SqlConnection conn, SqlCommand cmd)
        {
            if (!ft.GetPermissionByName("Security"))
            {
                InsertIntoPermission("Security", conn, cmd);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="cmd"></param>
        private void ImportUpgradePermission(SqlConnection conn, SqlCommand cmd)
        {
            if (!ft.GetPermissionByName("Upgrade"))
            {
                InsertIntoPermission("Upgrade", conn, cmd);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="cmd"></param>
        private void ImportCreateNewTemplatePermission(SqlConnection conn, SqlCommand cmd)
        {
            if (!ft.GetPermissionByName("Save/Amend Template"))
            {
                InsertIntoPermission("Save/Amend Template", conn, cmd);
            }
        }
        ///// <summary>
        ///// 
        ///// </summary>
        ///// <param name="conn"></param>
        ///// <param name="cmd"></param>
        //private void ImportPreferencePermission(SqlConnection conn, SqlCommand cmd)
        //{
        //    if (!ft.GetPermissionByName("Preferences"))
        //    {
        //        InsertIntoPermission("Preferences", conn, cmd);
        //    }
        //}
        /// <summary>
        /// 
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="cmd"></param>
        private void ImportAmendCodeDescription(SqlConnection conn, SqlCommand cmd)
        {
            if (!ft.GetPermissionByName("Amend Code Descriptions"))
            {
                InsertIntoPermission("Amend Code Descriptions", conn, cmd);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="cmd"></param>
        private void ImportCreateNewButtonPermission(SqlConnection conn, SqlCommand cmd)
        {
            if (!ft.GetPermissionByName("Create New Action"))
            {
                InsertIntoPermission("Create New Action", conn, cmd);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="cmd"></param>
        private void ImportAddDocumentViewPermission(SqlConnection conn, SqlCommand cmd)
        {
            if (!ft.GetPermissionByName("Add Document View"))
            {
                InsertIntoPermission("Add Document View", conn, cmd);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="cmd"></param>
        private void ImportAttachPDFPermission(SqlConnection conn, SqlCommand cmd)
        {
            if (!ft.GetPermissionByName("Attach PDF"))
            {
                InsertIntoPermission("Attach PDF", conn, cmd);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="permissionName"></param>
        /// <param name="conn"></param>
        /// <param name="cmd"></param>
        private void InsertIntoPermission(string permissionName, SqlConnection conn, SqlCommand cmd)
        {
            int index = this.dataGridView1.Rows.Add();
            try
            {
                Log("Initialize " + permissionName + " 's Permission - Global.", index, "", "");
                cmd = new SqlCommand("rsPermissions_Ins", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@PermissionName", permissionName + " - Global"));
                cmd.Parameters.Add(new SqlParameter("@TemplateID", ""));
                cmd.Parameters.Add(new SqlParameter("@ActionID", ""));
                cmd.Parameters.Add(new SqlParameter("@Per_Type", "3"));//0.Template - Save/Amend,1.Template - Write,2.Template - Read 3. global4.create new action
                cmd.Parameters.Add(new SqlParameter("@remark", ""));
                cmd.Parameters.Add(new SqlParameter("@Folder", ""));
                cmd.ExecuteNonQuery();
                Log("Initialize " + permissionName + " 's Permission - Global.", index, "Success!", "");
            }
            catch (Exception ex)
            {
                Log("Initialize " + permissionName + " 's Permission - Global.", index, "Fail!", ex.Message);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="cmd"></param>
        private void ImportTemplatesIntoDB(SqlConnection conn, SqlCommand cmd)
        {
            var path = string.Empty;
            getTemplateFolder(ref path);
            if (!string.IsNullOrEmpty(path))
            {
                try
                {
                    DirectoryInfo di = new DirectoryInfo(path);
                    DirectoryInfo[] fldrs = di.GetDirectories("*.*");
                    if (Finance_Tools.IsAddedToGallery(di))
                    {
                        try
                        {
                            FileInfo[] myfile = di.GetFiles();
                            foreach (FileInfo f in myfile)
                            {
                                int index = this.dataGridView1.Rows.Add();
                                Log("Importing " + f.Name + "...", index, "", "");
                                DirectoryInfo dir = new DirectoryInfo(di.FullName);
                                if (f.Extension == ".xls" || f.Extension == ".xlsx" || f.Extension == ".xlsm" || f.Extension == ".pdf")
                                {
                                    string Description = string.Empty;
                                    try
                                    {
                                        ShellLib.ShellLibClass sl = new ShellLib.ShellLibClass();
                                        Description = sl.getFileDetail(di.FullName, f.Name);
                                    }
                                    catch { }
                                    InsertIntoDB(f, Description, index, conn, cmd);
                                }
                                else
                                {
                                    Log("Importing " + f.Name + "...", index, "Fail!", "File format is not correct!File format must be xls , xlsx , xlsm or pdf ");
                                }
                            }
                        }
                        catch (Exception ex)
                        { MessageBox.Show(ex.Message, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information); }
                    }
                    foreach (DirectoryInfo d in fldrs)
                    {
                        if (Finance_Tools.IsAddedToGallery(d))
                        {
                            try
                            {
                                FileInfo[] myfile = d.GetFiles();
                                foreach (FileInfo f in myfile)
                                {
                                    int index = this.dataGridView1.Rows.Add();
                                    Log("Importing " + f.Name + "...", index, "", "");
                                    DirectoryInfo dir = new DirectoryInfo(d.FullName);
                                    if (f.Extension == ".xls" || f.Extension == ".xlsx" || f.Extension == ".xlsm" || f.Extension == ".pdf")
                                    {
                                        string Description = string.Empty;
                                        try
                                        {
                                            ShellLib.ShellLibClass sl = new ShellLib.ShellLibClass();
                                            Description = sl.getFileDetail(d.FullName, f.Name);
                                        }
                                        catch { }
                                        InsertIntoDB(f, Description, index, conn, cmd);
                                    }
                                    else
                                    {
                                        Log("Importing " + f.Name + "...", index, "Fail!", "File format is not correct!File format must be xls , xlsx , xlsm or pdf ");
                                    }
                                }
                            }
                            catch (Exception ex)
                            { MessageBox.Show(ex.Message, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information); }
                        }
                    }
                }
                catch (Exception ex)
                { MessageBox.Show(ex.Message, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information); }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ActionText"></param>
        /// <param name="index"></param>
        /// <param name="Result"></param>
        /// <param name="Message"></param>
        public void Log(string ActionText, int index, string Result, string Message)
        {
            try
            {
                this.dataGridView1.Rows[index].Cells[0].Value = ActionText;
                this.dataGridView1.Rows[index].Cells[1].Value = Result + Message;
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="f"></param>
        /// <param name="Description"></param>
        /// <param name="index"></param>
        /// <param name="conn"></param>
        /// <param name="cmd"></param>
        private void InsertIntoDB(FileInfo f, string Description, int index, SqlConnection conn, SqlCommand cmd)
        {
            fileNameforInsertDB = Path.GetFileNameWithoutExtension(f.Name);
            fileTypeForInsertDB = Path.GetExtension(f.Name);
            if (!ft.GetTemplateByPath(f.FullName))
            {
                assignSameName();
                try
                {
                    cmd = new SqlCommand("rsTemplates_Ins", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@TemplateData", ft.GetData(f.FullName)));
                    cmd.Parameters.Add(new SqlParameter("@TemplateName", fileNameforInsertDB));
                    cmd.Parameters.Add(new SqlParameter("@OriginTemplatePath", f.FullName));
                    cmd.Parameters.Add(new SqlParameter("@FileType", fileTypeForInsertDB));
                    cmd.Parameters.Add(new SqlParameter("@Description", Description));
                    SqlParameter parReturn = new SqlParameter("@ReturnValue", SqlDbType.Int);
                    parReturn.Direction = ParameterDirection.ReturnValue;
                    cmd.Parameters.Add(parReturn);
                    cmd.ExecuteNonQuery();
                    Log("Importing " + fileNameforInsertDB + fileTypeForInsertDB + "...", index, "Success!", "");
                    InitializeTemplatePermissions(cmd, conn, parReturn.Value.ToString(), fileNameforInsertDB + fileTypeForInsertDB, Description);
                }
                catch (Exception ex)
                {
                    Log("Importing " + fileNameforInsertDB + fileTypeForInsertDB + "...", index, "Fail!", ex.Message);
                }
            }
            else
                Log("Importing " + fileNameforInsertDB + fileTypeForInsertDB + "...", index, "Fail!", "File already exists!");
        }
        string fileNameforInsertDB = string.Empty;
        string fileTypeForInsertDB = string.Empty;
        /// <summary>
        /// 
        /// </summary>
        private void assignSameName()
        {
            while (ft.GetTemplateByNameAndType(fileNameforInsertDB, fileTypeForInsertDB))
            {
                fileNameforInsertDB += "Copy";
                assignSameName();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="conn"></param>
        /// <param name="id"></param>
        /// <param name="name"></param>
        /// <param name="desc"></param>
        private void InitializeTemplatePermissions(SqlCommand cmd, SqlConnection conn, string id, string name, string desc)
        {
            int index = this.dataGridView1.Rows.Add();
            try
            {
                Log("Initialize " + name + " 's Permission - Save/Amend.", index, "", "");
                cmd = new SqlCommand("rsPermissions_Ins", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@PermissionName", name + " - Save/Amend"));
                cmd.Parameters.Add(new SqlParameter("@TemplateID", id));
                cmd.Parameters.Add(new SqlParameter("@ActionID", ""));
                cmd.Parameters.Add(new SqlParameter("@Per_Type", "0"));//0.Template - Save/Amend,1.Template - Write,2.Template - Read 3.Global4.create new action 5.Template - Delete
                cmd.Parameters.Add(new SqlParameter("@remark", desc));
                cmd.Parameters.Add(new SqlParameter("@Folder", ""));
                cmd.ExecuteNonQuery();
                Log("Initialize " + name + " 's Permission - Save/Amend.", index, "Success!", desc);
                cmd.Parameters.Clear();
                index = this.dataGridView1.Rows.Add();
                Log("Initialize " + name + " 's Permission - Write.", index, "", "");
                cmd.Parameters.Add(new SqlParameter("@PermissionName", name + " - Write"));
                cmd.Parameters.Add(new SqlParameter("@TemplateID", id));
                cmd.Parameters.Add(new SqlParameter("@ActionID", ""));
                cmd.Parameters.Add(new SqlParameter("@Per_Type", "1"));
                cmd.Parameters.Add(new SqlParameter("@remark", desc));
                cmd.Parameters.Add(new SqlParameter("@Folder", ""));
                cmd.ExecuteNonQuery();
                Log("Initialize " + name + " 's Permission - Write.", index, "Success!", "");
                cmd.Parameters.Clear();
                index = this.dataGridView1.Rows.Add();
                Log("Initialize " + name + " 's Permission - Read.", index, "", "");
                cmd.Parameters.Add(new SqlParameter("@PermissionName", name + " - Read"));
                cmd.Parameters.Add(new SqlParameter("@TemplateID", id));
                cmd.Parameters.Add(new SqlParameter("@ActionID", ""));
                cmd.Parameters.Add(new SqlParameter("@Per_Type", "2"));
                cmd.Parameters.Add(new SqlParameter("@remark", desc));
                cmd.Parameters.Add(new SqlParameter("@Folder", ""));
                cmd.ExecuteNonQuery();
                Log("Initialize " + name + " 's Permission - Read.", index, "Success!", "");
                cmd.Parameters.Clear();
                index = this.dataGridView1.Rows.Add();
                Log("Initialize " + name + " 's Permission - Delete.", index, "", "");
                cmd.Parameters.Add(new SqlParameter("@PermissionName", name + " - Delete"));
                cmd.Parameters.Add(new SqlParameter("@TemplateID", id));
                cmd.Parameters.Add(new SqlParameter("@ActionID", ""));
                cmd.Parameters.Add(new SqlParameter("@Per_Type", "5"));
                cmd.Parameters.Add(new SqlParameter("@remark", desc));
                cmd.Parameters.Add(new SqlParameter("@Folder", ""));
                cmd.ExecuteNonQuery();
                Log("Initialize " + name + " 's Permission - Delete.", index, "Success!", "");
            }
            catch (Exception ex)
            {
                Log("Initialize " + name + " 's Permissions.", index, "Fail!", ex.Message);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateFolder"></param>
        private void setTemplateFolder(ref string templateFolder)
        {
            this.errorProvider1.Clear();
            if (string.IsNullOrEmpty(txtpath.Text))
            {
                this.errorProvider1.SetError(this.txtpath, "Specify the path for the template file from which to import the data. For example, C:\\MyData, \\Sales\\Northwind. Or, click Browse."); return;
            }
            templateFolder = txtpath.Text;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private SqlConnection getSqlConnection()
        {
            FileStream aFile = new FileStream("C:\\ProgramData\\RSData\\RSDataConfig\\Server.txt", FileMode.Open);
            StreamReader sr = new StreamReader(aFile);
            string strLine = sr.ReadLine();
            string ServerName = DEncrypt.Decrypt(strLine);
            sr.Close();
            string connString = string.Format("Data Source={0};Initial Catalog=RSData;Integrated Security=True;MultipleActiveResultSets=True", ServerName);
            SqlConnection sqlConn = new SqlConnection(connString);
            return sqlConn;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateFolder"></param>
        private void getTemplateFolder(ref string templateFolder)
        {
            if (cbDataSource.SelectedIndex == 0)
                setTemplateFolder(ref templateFolder);
            else
            {
                SqlConnection sqlConn = getSqlConnection();
                try
                {
                    sqlConn.Open();
                    SqlCommand cmd = new SqlCommand("select ft_folder from FinTools_Settings", sqlConn);
                    SqlDataAdapter sdap = new SqlDataAdapter();
                    sdap.SelectCommand = cmd;
                    DataTable dt = new DataTable();
                    sdap.Fill(dt);
                    if (dt.Rows.Count > 0)
                        templateFolder = dt.Rows[0][0].ToString();
                    else
                        throw new Exception("Previous version addin's templates folder is null. Please choose the Templates Source Folder.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LogHelper.WriteLog(typeof(Ribbon2), ex.Message + " - SYNC error");
                }
                finally
                {
                    if (sqlConn != null)
                        sqlConn.Close();
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
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
                txtpath.Text = fbdAd_UpdateFolder.SelectedPath;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbDataSource_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbDataSource.SelectedIndex == 0)
            {
                txtpath.Visible = true;
                btnBrowse.Visible = true;
                lblFolder.Visible = true;
            }
            else
            {
                txtpath.Visible = false;
                btnBrowse.Visible = false;
                lblFolder.Visible = false;
            }
        }
    }
}
