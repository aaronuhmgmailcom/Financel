
/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<LoginFrm>   
 * Author：Peter.uhm  (yanb@shinetechchina.com)
 * Modify date：2016.04
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
using System.IO;
using System.Reflection;
using System.Data.SqlClient;
using System.Configuration;
using ExcelAddIn4.Common;

namespace ExcelAddIn4
{
    public partial class FileNameForm : Form
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
        public static string filename;
        /// <summary>
        /// 
        /// </summary>
        public static string folderName;
        /// <summary>
        /// 
        /// </summary>
        public static string fileType;
        /// <summary>
        /// 
        /// </summary>
        public FileNameForm()
        {
            InitializeComponent();
            InitializeTemplate();
            initializeFolder();
            if (!string.IsNullOrEmpty(SessionInfo.UserInfo.FilePath))
            {
                DirectoryInfo di2 = new DirectoryInfo(SessionInfo.UserInfo.FilePath);
                cbFolders.Text = di2.Parent.Name;
                cbTmpName.Visible = true;
                btnDelTemp.Visible = true;
            }
            else
            {
                btnDelTemp.Visible = false;
            }
            try
            {
                cbTmpName.SelectedItem = new KeyValuePair<string, string>(SessionInfo.UserInfo.FileName, SessionInfo.UserInfo.File_ftid);
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        private void InitializeTemplate()
        {
            for (int i = 0; i < Ribbon2.TemplateAndPath.Count; i++)
                if (BasePage.VerifyAmendButton(Ribbon2.TemplateAndPath[i].Key))
                    cbTmpName.Items.Add(new KeyValuePair<string, string>(Ribbon2.TemplateAndPath[i].Key, Ribbon2.TemplateAndPath[i].Value));

            cbTmpName.DisplayMember = "Key";
            cbTmpName.ValueMember = "Value";
        }
        /// <summary>
        /// 
        /// </summary>
        private void initializeFolder()
        {
            var path = Finance_Tools.RootPath;
            DirectoryInfo di = new DirectoryInfo(path);
            DirectoryInfo[] fldrs = di.GetDirectories("*.*");
            foreach (DirectoryInfo d in fldrs)
            {
                if (Finance_Tools.IsAddedToGallery(d))
                    cbFolders.Items.Add(new KeyValuePair<string, string>(d.Name, d.FullName));
            }
            cbFolders.DisplayMember = "Key";
            cbFolders.ValueMember = "Value";
        }
        /// <summary>
        /// 
        /// </summary>
        private string TempName
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="elem"></param>
        /// <returns></returns>
        private bool EquaWithButtonName(KeyValuePair<string, string> elem)
        {
            if (elem.Key == TempName)
                return true;
            return false;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbFolder_SelectedIndexChanged(object sender, EventArgs e)
        {
            KeyValuePair<string, string> obj = (KeyValuePair<string, string>)this.cbFolders.SelectedItem;
            DirectoryInfo d = new DirectoryInfo(obj.Value);
            cbTmpName.Items.Clear();
            try
            {
                FileInfo[] myfile = d.GetFiles();
                for (int j = myfile.Length; j > 0; j--)
                {
                    FileInfo f = myfile[j - 1];
                    TempName = Path.GetFileNameWithoutExtension(f.Name);
                    Predicate<KeyValuePair<string, string>> pred2 = EquaWithButtonName;
                    try
                    {
                        KeyValuePair<string, string> kv = Ribbon2.TemplateAndPath.Find(pred2);
                        if ((kv.Key != null) && (BasePage.VerifyAmendButton(kv.Key)))
                            cbTmpName.Items.Add(kv);
                    }
                    catch { }
                }
                cbTmpName.Items.Add(new KeyValuePair<string, string>("<New Template...>", "-1"));
                textBox1.Text = "";
            }
            catch { }
            btnDelTemp.Visible = false;
            textBox1.Visible = false;
            cbTmpName.Visible = true;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, EventArgs e)
        {
            KeyValuePair<string, string> obj = (KeyValuePair<string, string>)this.cbFolders.SelectedItem;
            KeyValuePair<string, string> obj2 = (KeyValuePair<string, string>)this.cbTmpName.SelectedItem;
            if (string.IsNullOrEmpty(this.textBox1.Text.Trim()) || this.cbFolders.SelectedItem == null)
            {
                MessageBox.Show("Folder name or file name is invalid.", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            else
            {
                fileType = Path.GetExtension(Globals.ThisAddIn.Application.ActiveWorkbook.FullName);
                if (string.IsNullOrEmpty(fileType)) fileType = ".xlsm";
                folderName = obj.Value;
                if (textBox1.Visible == true)
                {
                    filename = this.textBox1.Text;
                    assignSameName();
                }
                else
                {
                    filename = cbTmpName.Text;
                }
                this.Close();
            }
            if (SaveToFolder())
            {
                try
                {
                    if (textBox1.Visible == true)
                        insertTemplate();
                    else
                        UpdateTemplate(obj2.Value);
                }
                catch { }
                ft.uPDATETemplateUpdateFlag(true);
                MessageBox.Show("Success!", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        private void assignSameName()
        {
            while (ft.GetTemplateByNameAndType(filename, fileType))
            {
                filename += "Copy";
                assignSameName();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        private void insertTemplate()
        {
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsTemplates_Ins", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@TemplateData", ft.GetData(folderName + "\\" + filename + fileType)));
                cmd.Parameters.Add(new SqlParameter("@TemplateName", filename));
                cmd.Parameters.Add(new SqlParameter("@OriginTemplatePath", folderName + "\\" + filename + fileType));
                cmd.Parameters.Add(new SqlParameter("@FileType", fileType));
                cmd.Parameters.Add(new SqlParameter("@Description", ""));
                SqlParameter parReturn = new SqlParameter("@ReturnValue", SqlDbType.Int);
                parReturn.Direction = ParameterDirection.ReturnValue;
                cmd.Parameters.Add(parReturn);
                cmd.ExecuteNonQuery();
                InitializeTemplatePermissions(cmd, conn, parReturn.Value.ToString(), filename + fileType, "");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LogHelper.WriteLog(typeof(FileNameForm), e.Message + " FileNameForm error");
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
            try
            {
                cmd = new SqlCommand("rsPermissions_Ins", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@PermissionName", name + " - Save/Amend"));
                cmd.Parameters.Add(new SqlParameter("@TemplateID", id));
                cmd.Parameters.Add(new SqlParameter("@ActionID", ""));
                cmd.Parameters.Add(new SqlParameter("@Per_Type", "0"));//0.Template - Save/Amend,1.Template - Write,2.Template - Read 3.Global4.create new action 5.Template - Delete
                cmd.Parameters.Add(new SqlParameter("@remark", desc));
                cmd.Parameters.Add(new SqlParameter("@Folder", ""));
                cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();
                cmd.Parameters.Add(new SqlParameter("@PermissionName", name + " - Write"));
                cmd.Parameters.Add(new SqlParameter("@TemplateID", id));
                cmd.Parameters.Add(new SqlParameter("@ActionID", ""));
                cmd.Parameters.Add(new SqlParameter("@Per_Type", "1"));
                cmd.Parameters.Add(new SqlParameter("@remark", desc));
                cmd.Parameters.Add(new SqlParameter("@Folder", ""));
                cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();
                cmd.Parameters.Add(new SqlParameter("@PermissionName", name + " - Read"));
                cmd.Parameters.Add(new SqlParameter("@TemplateID", id));
                cmd.Parameters.Add(new SqlParameter("@ActionID", ""));
                cmd.Parameters.Add(new SqlParameter("@Per_Type", "2"));
                cmd.Parameters.Add(new SqlParameter("@remark", desc));
                cmd.Parameters.Add(new SqlParameter("@Folder", ""));
                cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();
                cmd.Parameters.Add(new SqlParameter("@PermissionName", name + " - Delete"));
                cmd.Parameters.Add(new SqlParameter("@TemplateID", id));
                cmd.Parameters.Add(new SqlParameter("@ActionID", ""));
                cmd.Parameters.Add(new SqlParameter("@Per_Type", "5"));
                cmd.Parameters.Add(new SqlParameter("@remark", desc));
                cmd.Parameters.Add(new SqlParameter("@Folder", ""));
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LogHelper.WriteLog(typeof(FileNameForm), ex.Message + " FileNameForm error");
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateid"></param>
        private void UpdateTemplate(string templateid)
        {
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsTemplates_Upd", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@TemplateData", ft.GetData(folderName + "\\" + filename + fileType)));
                cmd.Parameters.Add(new SqlParameter("@FileType", fileType));
                cmd.Parameters.Add(new SqlParameter("@templateID", int.Parse(templateid)));
                cmd.ExecuteNonQuery();
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        private bool SaveToFolder()
        {
            try
            {
                string newTemplateFolderName = FileNameForm.folderName;
                if (string.IsNullOrEmpty(newTemplateFolderName))
                {
                    MessageBox.Show("Please choose a folder.", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }
                string savefilename = FileNameForm.filename;
                if (string.IsNullOrEmpty(savefilename))
                {
                    MessageBox.Show("Please input a file name.", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }
                Globals.ThisAddIn.Application.DisplayAlerts = false;
                if (fileType == ".xlsm")
                    Globals.ThisAddIn.Application.ActiveWorkbook.SaveAs(newTemplateFolderName + "\\" + savefilename + fileType, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                else
                    Globals.ThisAddIn.Application.ActiveWorkbook.SaveAs(newTemplateFolderName + "\\" + savefilename + fileType, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                Globals.ThisAddIn.Application.DisplayAlerts = true;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Globals.ThisAddIn.Application.ActiveWorkbook.Close();
                return false;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDelTemp_Click(object sender, EventArgs e)
        {
            ft.uPDATETemplateUpdateFlag(true);
            DoDelete();
            cbFolder_SelectedIndexChanged(null, null);
        }
        /// <summary>
        /// 
        /// </summary>
        private void DoDelete()
        {
            if (MessageBox.Show("This operation will delete template and related permissions,Are you sure to do this operation?", "Message - RSystems FinanceTools", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                KeyValuePair<string, string> obj = (KeyValuePair<string, string>)this.cbTmpName.SelectedItem;
                DeleteFile();
                DeleteTemplate(obj.Value);
                DeleteTemplateButtons(obj.Value);
                DeleteTemplatePermissions(obj.Value);
                DeleteTemplateVisible(obj.Value);
                TempName = obj.Key;
                Predicate<KeyValuePair<string, string>> pred2 = EquaWithButtonName;
                KeyValuePair<string, string> kv = Ribbon2.TemplateAndPath.Find(pred2);
                if (kv.Key != null)
                {
                    cbTmpName.Items.Remove(kv);
                    Ribbon2.TemplateAndPath.Remove(kv);
                    for (int i = 0; i < Ribbon2.ButtonViewCount; i++)
                    {
                        try
                        {
                            Microsoft.Office.Tools.Ribbon.RibbonControl rb = (Microsoft.Office.Tools.Ribbon.RibbonButton)Globals.Ribbons[0].Tabs[0].Groups[2].Items[i];
                            string[] sArray = System.Text.RegularExpressions.Regex.Split(rb.Tag.ToString(), ",");
                            string templateid = sArray[0];
                            if (templateid == obj.Value)
                                rb.Visible = false;
                        }
                        catch { }
                    }
                    Microsoft.Office.Tools.Ribbon.RibbonGroup rg = (Microsoft.Office.Tools.Ribbon.RibbonGroup)Globals.Ribbons[0].Tabs[0].Groups[3];
                    for (int j = 0; j < rg.Items.Count; j++)
                    {
                        Microsoft.Office.Tools.Ribbon.RibbonMenu rm = (Microsoft.Office.Tools.Ribbon.RibbonMenu)rg.Items[j];
                        for (int k = 0; k < rm.Items.Count; k++)
                        {
                            Microsoft.Office.Tools.Ribbon.RibbonToggleButton rtb = (Microsoft.Office.Tools.Ribbon.RibbonToggleButton)rm.Items[k];
                            if (rtb.Tag == obj.Value)
                                rtb.Visible = false;
                        }
                    }
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        private void DeleteFile()
        {
            try
            {
                KeyValuePair<string, string> obj = (KeyValuePair<string, string>)this.cbFolders.SelectedItem;
                if (File.Exists(obj.Value + "\\" + cbTmpName.Text + ".xlsm"))
                {
                    if (Globals.ThisAddIn.Application.ActiveWorkbook.FullName == (obj.Value + "\\" + cbTmpName.Text + ".xlsm"))
                        Globals.ThisAddIn.Application.ActiveWorkbook.Close();
                    File.Delete(obj.Value + "\\" + cbTmpName.Text + ".xlsm");
                }
                else if (File.Exists(obj.Value + "\\" + cbTmpName.Text + ".xlsx"))
                {
                    if (Globals.ThisAddIn.Application.ActiveWorkbook.FullName == (obj.Value + "\\" + cbTmpName.Text + ".xlsx"))
                        Globals.ThisAddIn.Application.ActiveWorkbook.Close();
                    File.Delete(obj.Value + "\\" + cbTmpName.Text + ".xlsx");
                }
                else if (File.Exists(obj.Value + "\\" + cbTmpName.Text + ".xls"))
                {
                    if (Globals.ThisAddIn.Application.ActiveWorkbook.FullName == (obj.Value + "\\" + cbTmpName.Text + ".xls"))
                        Globals.ThisAddIn.Application.ActiveWorkbook.Close();
                    File.Delete(obj.Value + "\\" + cbTmpName.Text + ".xls");
                }
            }
            catch (Exception ex) { LogHelper.WriteLog(typeof(FileNameForm), ex.Message + " FileNameForm error"); }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateid"></param>
        private void DeleteTemplateVisible(string templateid)
        {
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsUsersTemplatesVisible_DelByTmpID", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@TemplateID", templateid));
                cmd.ExecuteNonQuery();
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateid"></param>
        private void DeleteTemplatePermissions(string templateid)
        {
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsPermissions_DelByTmpID", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@templateid", templateid));
                cmd.ExecuteNonQuery();
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateid"></param>
        private void DeleteTemplateButtons(string templateid)
        {
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsTemplateActions_DelByTmpID", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@templateid", templateid));
                cmd.ExecuteNonQuery();
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateid"></param>
        private void DeleteTemplate(string templateid)
        {
            SqlConnection conn = null;
            SqlDataReader rdr = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsTemplates_Del", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@TemplateID", templateid));
                rdr = cmd.ExecuteReader();
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
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbTmpName_SelectedIndexChanged(object sender, EventArgs e)
        {
            KeyValuePair<string, string> obj = (KeyValuePair<string, string>)this.cbTmpName.SelectedItem;
            if (obj.Value == "-1")
            {
                btnDelTemp.Visible = false;
                textBox1.Visible = true;
                cbTmpName.Visible = false;
                textBox1.Text = "";
            }
            else if (obj.Value != "-1")
            {
                textBox1.Visible = false;
                cbTmpName.Visible = true;
                textBox1.Text = obj.Key;
                BasePage.VerifyDelButton(obj.Key, "5", this.btnDelTemp);
            }
        }
    }
}
