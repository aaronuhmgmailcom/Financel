/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<Finance_Tools>   
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
using System.Data.SqlClient;
using System.Configuration;
using System.Text.RegularExpressions;

namespace ExcelAddIn4
{
    public partial class Preferences : Form
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
        string templateIDStr = string.Empty;
        /// <summary>
        /// 
        /// </summary>
        string templateCheckUnCheck = string.Empty;
        /// <summary>
        /// 
        /// </summary>
        public Preferences()
        {
            InitializeComponent();
            BindXMLEntityInfo();
            BindSunInfo();
            BindTreeView();
            treeFiles.AfterCheck += new TreeViewEventHandler(treeFiles_AfterCheck);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void treeFiles_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Level == 0)
            {
                if (e.Node.Checked)
                {
                    foreach (TreeNode tnTemp in e.Node.Nodes)
                    {
                        tnTemp.Checked = true;
                        templateIDStr += tnTemp.Tag + ",";
                        templateCheckUnCheck += "True,";
                    }
                }
                else
                {
                    foreach (TreeNode tnTemp in e.Node.Nodes)
                    {
                        tnTemp.Checked = false;
                        templateIDStr += tnTemp.Tag + ",";
                        templateCheckUnCheck += "False,";
                    }
                }
            }
            else
            {
                templateIDStr += e.Node.Tag + ",";
                templateCheckUnCheck += e.Node.Checked.ToString() + ",";
            }
        }
        /// <summary>
        /// 
        /// </summary>
        private void BindTreeView()
        {
            List<string> list = ft.GetUserGroups(SessionInfo.UserInfo.ID);
            for (int i = 0; i < list.Count; i++)
            {
                try
                {
                    string groupid = list[i];
                    bool? groupDisable = ft.GetGroupDisableByID(int.Parse(groupid));
                    if (!(bool)groupDisable)
                    {
                        List<string> list2 = ft.GetPermissionsByGroupID(groupid);
                        for (int j = 0; j < list2.Count; j++)
                        {
                            try
                            {
                                string permissionid = list2[j];
                                string Templateid = ft.GetTemplateIDByPermissionID(int.Parse(permissionid));
                                string permissionRemark = ft.GetPermissionRemarkByID(int.Parse(permissionid));
                                string templateName = ft.GetTemplateNameByID(int.Parse(Templateid));
                                string templatePath = ft.GetTemplatePathByID(int.Parse(Templateid));
                                string Folder = ft.GetFolderByID(int.Parse(permissionid));
                                if (!string.IsNullOrEmpty(Folder) && !isFolderExist(treeFiles.Nodes, Folder))
                                    treeFiles.Nodes.Add(Folder);

                                if (!isNodeExist(treeFiles.Nodes, templateName))
                                {
                                    TreeNode tn = new TreeNode();
                                    foreach (TreeNode node in treeFiles.Nodes)
                                    {
                                        if (node.Text.ToUpper() == Folder.ToUpper())
                                        {
                                            tn = node.Nodes.Add(templateName);
                                            string str = ft.GetVisibleByUserIDTemplateID(SessionInfo.UserInfo.ID, Templateid);
                                            if (str.Trim() == "True") tn.Checked = true;
                                            tn.ToolTipText = permissionRemark + "{" + templatePath + "}";
                                            tn.Tag = Templateid;
                                        }
                                    }
                                }
                            }
                            catch { continue; }
                        }
                    }
                }
                catch { continue; }
            }
            treeFiles.ShowNodeToolTips = true;
            treeFiles.ExpandAll();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Node"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        private bool isNodeExist(TreeNodeCollection Nodes, string name)
        {
            foreach (TreeNode node in Nodes)
            {
                foreach (TreeNode nd in node.Nodes)
                    if (nd.Text.ToUpper() == name.ToUpper())
                        return true;
            }
            return false;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Node"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        private bool isFolderExist(TreeNodeCollection Nodes, string name)
        {
            foreach (TreeNode node in Nodes)
            {
                if (node.Text.ToUpper() == name.ToUpper())
                    return true;
            }
            return false;
        }
        /// <summary>
        /// 
        /// </summary>
        private void BindXMLEntityInfo()
        {
            this.txtSA.Text = Finance_Tools.GetAppConfig("SuspenseAccount");
            this.txtBU.Text = Finance_Tools.GetAppConfig("BusinessUnit");
        }
        /// <summary>
        /// 
        /// </summary>
        private void BindSunInfo()
        {
            this.txtSunServer.Text = SessionInfo.UserInfo.SunUserIP;
            this.txtSunID.Text = SessionInfo.UserInfo.SunUserID;
            this.txtSunPass.Text = SessionInfo.UserInfo.SunUserPass;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            SqlConnection conn = null;
            SessionInfo.UserInfo.SunUserIP = this.txtSunServer.Text;
            SessionInfo.UserInfo.SunUserID = this.txtSunID.Text;
            SessionInfo.UserInfo.SunUserPass = this.txtSunPass.Text;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsUsers_UserSunInfo_Upd", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@SUNUserIP", SessionInfo.UserInfo.SunUserIP));
                cmd.Parameters.Add(new SqlParameter("@SUNUserID", DEncrypt.Encrypt(SessionInfo.UserInfo.SunUserID)));
                cmd.Parameters.Add(new SqlParameter("@SUNUserPass", DEncrypt.Encrypt(SessionInfo.UserInfo.SunUserPass)));
                cmd.Parameters.Add(new SqlParameter("@id", SessionInfo.UserInfo.ID));
                cmd.ExecuteNonQuery();
                this.Close();
                Finance_Tools.AppSettingSave("SuspenseAccount", this.txtSA.Text);
                Finance_Tools.AppSettingSave("BusinessUnit", this.txtBU.Text);
                SaveVisibleSetting(conn, cmd);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (conn != null)
                    conn.Close();
            }
            ft.uPDATETemplateUpdateFlag(SessionInfo.UserInfo.ID, true);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="cmd"></param>
        private void SaveVisibleSetting(SqlConnection conn, SqlCommand cmd)
        {
            string[] sArray = Regex.Split(templateIDStr, ",", RegexOptions.IgnoreCase);
            string[] sArray2 = Regex.Split(templateCheckUnCheck, ",", RegexOptions.IgnoreCase);
            for (int i = 0; i < sArray.Length; i++)
                Update(conn, cmd, sArray2[i], sArray[i]);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="cmd"></param>
        /// <param name="valu"></param>
        /// <param name="templateid"></param>
        private void Update(SqlConnection conn, SqlCommand cmd, string valu, string templateid)
        {
            cmd = new SqlCommand("rsUsersTemplatesVisible_Upd", conn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Clear();
            cmd.Parameters.Add(new SqlParameter("@TemplateID", templateid));
            cmd.Parameters.Add(new SqlParameter("@UserID", SessionInfo.UserInfo.ID));
            cmd.Parameters.Add(new SqlParameter("@Visible", valu));
            cmd.ExecuteNonQuery();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
