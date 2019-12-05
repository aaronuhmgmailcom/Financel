/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<Security>   
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
using System.Data.SqlClient;
using System.Configuration;
using System.Text.RegularExpressions;
using ExcelAddIn4.Common;

namespace ExcelAddIn4
{
    public partial class Security : Form
    {
        DataGridView dgv = null;
        DataGridView dgvUsers = null;
        DataGridView dgvPermissions = null;
        DataTable dtsearch = new DataTable();
        string UserUpdRows = string.Empty;
        string AddInTabValue = string.Empty;
        string UserUpdRowIndex = string.Empty;
        string PermissionsUpdRows = string.Empty;
        string PermissionNames = string.Empty;
        string PermissionRemark = string.Empty;
        string PermissionFolder = string.Empty;
        string PermissionUpdRowIndex = string.Empty;
        string GroupUpdRows = string.Empty;
        string GroupNames = string.Empty;
        string GroupDisable = string.Empty;
        string GroupFolderMaskName = string.Empty;
        string GroupRemark = string.Empty;
        string GroupUpdRowIndex = string.Empty;
        string DeleteGroupIDs = string.Empty;
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
        public Security()
        {
            InitializeComponent();
            dgv = ft.IniGroups();
            dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
            dgv.AutoGenerateColumns = false;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgv.Dock = DockStyle.Fill;
            dgv.CellMouseDown += new DataGridViewCellMouseEventHandler(DataGridView1_CellMouseDown);
            dgv.RowPostPaint += new DataGridViewRowPostPaintEventHandler(dgv_RowPostPaint);
            dgv.CellValueChanged += new DataGridViewCellEventHandler(dgv_CellValueChanged);
            dgv.CellBeginEdit += new DataGridViewCellCancelEventHandler(dgv_CellBeginEdit);
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv.MultiSelect = false;
            dgv.CellClick += new DataGridViewCellEventHandler(dgv_CellClick);
            dgv.KeyDown += new KeyEventHandler(dgv_KeyDown);
            dgv.DataBindingComplete += new DataGridViewBindingCompleteEventHandler(dgv_DataBindingComplete);
            dgv.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            for (int i = 0; i < dgv.Columns.Count; i++) { dgv.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable; }
            BindData();
            this.tabPage3.Controls.Add(dgv);
            InitializeUsersDGV();
            InitializePermissionsDGV();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgv_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Up)
            {
                DataGridViewCellEventArgs args = new DataGridViewCellEventArgs(dgv.CurrentCell.ColumnIndex, dgv.CurrentRow.Index - 1);
                dgv_CellClick(dgv, args);
            }
            else if (e.KeyCode == Keys.Down)
            {
                DataGridViewCellEventArgs args = new DataGridViewCellEventArgs(dgv.CurrentCell.ColumnIndex, dgv.CurrentRow.Index + 1);
                dgv_CellClick(dgv, args);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgv_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            if (dgv.CurrentRow != null && (dgv.CurrentRow.Index == 0))
            {
                DataGridViewCellEventArgs args = new DataGridViewCellEventArgs(0, 0);
                dgv_CellClick(dgv, args);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgv_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgv.CurrentRow == null) return;
            int i = e.RowIndex < 0 ? 0 : (e.RowIndex >= dgv.Rows.Count ? (dgv.Rows.Count - 1) : e.RowIndex);
            string groupid = this.dgv.Rows[i].Cells["ID"].FormattedValue.ToString();
            BindPermissions(groupid);
            BindUsers(groupid);
            dtsearch = (DataTable)dgvPermissions.DataSource;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="groupid"></param>
        private void BindPermissions(string groupid)
        {
            for (int j = 0; j < dgvPermissions.Rows.Count; j++)
                dgvPermissions.Rows[j].Cells[1].Value = false;

            if (string.IsNullOrEmpty(groupid)) return;
            List<string> list = ft.GetPermissionsByGroupID(groupid);
            for (int i = 0; i < list.Count; i++)
                for (int j = 0; j < dgvPermissions.Rows.Count; j++)
                {
                    if (dgvPermissions.Rows[j].Cells["ID"].Value == null) continue;
                    if (dgvPermissions.Rows[j].Cells["ID"].Value.ToString() == list[i])
                        dgvPermissions.Rows[j].Cells[1].Value = true;
                }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="groupid"></param>
        private void BindUsers(string groupid)
        {
            for (int j = 0; j < dgvUsers.Rows.Count; j++)
                dgvUsers.Rows[j].Cells[1].Value = false;

            if (string.IsNullOrEmpty(groupid)) return;
            List<string> list = ft.GetUsersByGroupID(groupid);
            for (int i = 0; i < list.Count; i++)
                for (int j = 0; j < dgvUsers.Rows.Count; j++)
                {
                    if (dgvUsers.Rows[j].Cells[0].Value == null) continue;
                    if (dgvUsers.Rows[j].Cells[0].Value.ToString() == list[i])
                        dgvUsers.Rows[j].Cells[1].Value = true;
                }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgv_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            CheckCurrentRow(e.RowIndex);
            GroupUpdRows += "," + this.dgv.Rows[e.RowIndex].Cells["ID"].FormattedValue;
            GroupNames += "," + this.dgv.Rows[e.RowIndex].Cells["GroupName"].FormattedValue;
            GroupDisable += "," + this.dgv.Rows[e.RowIndex].Cells["GroupDisable"].FormattedValue;
            GroupFolderMaskName += "," + this.dgv.Rows[e.RowIndex].Cells["FolderMaskName"].FormattedValue;
            GroupRemark += "," + this.dgv.Rows[e.RowIndex].Cells["Remark"].FormattedValue;
            GroupUpdRowIndex += "," + e.RowIndex;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="index"></param>
        private void CheckCurrentRow(int index)
        {
            string[] sArray = Regex.Split(GroupUpdRows, ",", RegexOptions.IgnoreCase);
            string[] sGroupNames = Regex.Split(GroupNames, ",", RegexOptions.IgnoreCase);
            string[] sGroupDisable = Regex.Split(GroupDisable, ",", RegexOptions.IgnoreCase);
            string[] sGroupFolderMaskName = Regex.Split(GroupFolderMaskName, ",", RegexOptions.IgnoreCase);
            string[] sGroupRemark = Regex.Split(GroupRemark, ",", RegexOptions.IgnoreCase);
            string[] sGroupUpdRowIndex = Regex.Split(GroupUpdRowIndex, ",", RegexOptions.IgnoreCase);
            for (int i = 0; i < sGroupUpdRowIndex.Length; i++)
            {
                if (string.IsNullOrEmpty(sGroupUpdRowIndex[i])) continue;
                if (int.Parse(sGroupUpdRowIndex[i]) == index)
                {
                    sArray.SetValue("", i);
                    sGroupNames.SetValue("", i);
                    sGroupDisable.SetValue("", i);
                    sGroupFolderMaskName.SetValue("", i);
                    sGroupRemark.SetValue("", i);
                    sGroupUpdRowIndex.SetValue("", i);
                }
            }
            GroupUpdRows = "";
            GroupNames = "";
            GroupDisable = "";
            GroupFolderMaskName = "";
            GroupRemark = "";
            GroupUpdRowIndex = "";
            for (int i = 0; i < sArray.Length; i++)
            {
                if (string.IsNullOrEmpty(sArray[i]) && string.IsNullOrEmpty(sGroupNames[i]) && string.IsNullOrEmpty(sGroupDisable[i]) && string.IsNullOrEmpty(sGroupFolderMaskName[i]) && string.IsNullOrEmpty(sGroupRemark[i])) continue;
                GroupUpdRows += "," + sArray[i];
                GroupNames += "," + sGroupNames[i];
                GroupDisable += "," + sGroupDisable[i];
                GroupFolderMaskName += "," + sGroupFolderMaskName[i];
                GroupRemark += "," + sGroupRemark[i];
                GroupUpdRowIndex += "," + sGroupUpdRowIndex[i];
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgv_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            btnApply.Enabled = true;
        }
        /// <summary>
        /// 
        /// </summary>
        private void InitializePermissionsDGV()
        {
            dgvPermissions = ft.IniPermissions();
            dgvPermissions.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
            dgvPermissions.AutoGenerateColumns = false;
            dgvPermissions.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvPermissions.Dock = DockStyle.Fill;
            dgvPermissions.RowPostPaint += new DataGridViewRowPostPaintEventHandler(dgvPermissions_RowPostPaint);
            dgvPermissions.CellValueChanged += new DataGridViewCellEventHandler(dgvPermissions_CellValueChanged);
            dgvPermissions.CellBeginEdit += new DataGridViewCellCancelEventHandler(dgvPermissions_CellBeginEdit);
            BindDataPermissions();
            this.tabPage2.Controls.Add(dgvPermissions);
            for (int i = 0; i < dgvPermissions.Columns.Count; i++) { dgvPermissions.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable; }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvPermissions_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            btnApply.Enabled = true;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="index"></param>
        private void CheckCurrentPermissionsRow(int index)
        {
            string[] sArray = Regex.Split(PermissionsUpdRows, ",", RegexOptions.IgnoreCase);
            string[] sPermissionNames = Regex.Split(PermissionNames, ",", RegexOptions.IgnoreCase);
            string[] sPermissionRemark = Regex.Split(PermissionRemark, ",", RegexOptions.IgnoreCase);
            string[] sPermissionFolder = Regex.Split(PermissionFolder, ",", RegexOptions.IgnoreCase);
            string[] sPermissionUpdRowIndex = Regex.Split(PermissionUpdRowIndex, ",", RegexOptions.IgnoreCase);
            for (int i = 0; i < sPermissionUpdRowIndex.Length; i++)
            {
                if (string.IsNullOrEmpty(sPermissionUpdRowIndex[i])) continue;
                if (int.Parse(sPermissionUpdRowIndex[i]) == index)
                {
                    sArray.SetValue("", i);
                    sPermissionNames.SetValue("", i);
                    sPermissionRemark.SetValue("", i);
                    sPermissionFolder.SetValue("", i);
                    sPermissionUpdRowIndex.SetValue("", i);
                }
            }
            PermissionsUpdRows = "";
            PermissionNames = "";
            PermissionRemark = "";
            PermissionFolder = "";
            PermissionUpdRowIndex = "";
            for (int i = 0; i < sArray.Length; i++)
            {
                if (string.IsNullOrEmpty(sArray[i]) && string.IsNullOrEmpty(sPermissionNames[i])) continue;
                PermissionsUpdRows += "," + sArray[i];
                PermissionNames += "," + sPermissionNames[i];
                PermissionRemark += "," + sPermissionRemark[i];
                PermissionFolder += "," + sPermissionFolder[i];
                PermissionUpdRowIndex += "," + sPermissionUpdRowIndex[i];
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvPermissions_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            CheckCurrentPermissionsRow(e.RowIndex);
            PermissionsUpdRows += "," + this.dgvPermissions.Rows[e.RowIndex].Cells["ID"].FormattedValue;
            PermissionNames += "," + this.dgvPermissions.Rows[e.RowIndex].Cells["PermissionName"].FormattedValue;
            PermissionRemark += "," + this.dgvPermissions.Rows[e.RowIndex].Cells["Remark"].FormattedValue;
            PermissionFolder += "," + this.dgvPermissions.Rows[e.RowIndex].Cells["Folder"].FormattedValue;
            PermissionUpdRowIndex += "," + e.RowIndex;
        }
        /// <summary>
        /// 
        /// </summary>
        private void BindDataPermissions()
        {
            DataTable dt = ft.GetPermissionsFromDB();
            this.dgvPermissions.DataSource = dt;
        }
        /// <summary>
        /// 
        /// </summary>
        private void InitializeUsersDGV()
        {
            dgvUsers = ft.IniUsers();
            dgvUsers.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
            dgvUsers.AutoGenerateColumns = false;
            dgvUsers.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvUsers.Dock = DockStyle.Fill;
            dgvUsers.RowPostPaint += new DataGridViewRowPostPaintEventHandler(dgvUsers_RowPostPaint);
            dgvUsers.CellValueChanged += new DataGridViewCellEventHandler(dgvUsers_CellValueChanged);
            dgvUsers.CellBeginEdit += new DataGridViewCellCancelEventHandler(dgvUsers_CellBeginEdit);
            BindDataUsers();
            this.tabPage1.Controls.Add(dgvUsers);
            for (int i = 0; i < dgvUsers.Columns.Count; i++) { dgvUsers.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable; }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvUsers_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            btnApply.Enabled = true;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="index"></param>
        private void CheckCurrentUsersRow(int index)
        {
            string[] sArray = Regex.Split(UserUpdRows, ",", RegexOptions.IgnoreCase);
            string[] sAddInTabValue = Regex.Split(AddInTabValue, ",", RegexOptions.IgnoreCase);
            string[] sUserUpdRowIndex = Regex.Split(UserUpdRowIndex, ",", RegexOptions.IgnoreCase);
            for (int i = 0; i < sUserUpdRowIndex.Length; i++)
            {
                if (string.IsNullOrEmpty(sUserUpdRowIndex[i])) continue;
                if (int.Parse(sUserUpdRowIndex[i]) == index)
                {
                    sArray.SetValue("", i);
                    sAddInTabValue.SetValue("", i);
                    sUserUpdRowIndex.SetValue("", i);
                }
            }
            UserUpdRows = "";
            AddInTabValue = "";
            UserUpdRowIndex = "";
            for (int i = 0; i < sArray.Length; i++)
            {
                if (string.IsNullOrEmpty(sArray[i]) && string.IsNullOrEmpty(sAddInTabValue[i])) continue;
                UserUpdRows += "," + sArray[i];
                AddInTabValue += "," + sAddInTabValue[i];
                UserUpdRowIndex += "," + sUserUpdRowIndex[i];
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvUsers_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            CheckCurrentUsersRow(e.RowIndex);
            UserUpdRows += "," + this.dgvUsers.Rows[e.RowIndex].Cells["UserID"].FormattedValue;
            AddInTabValue += "," + this.dgvUsers.Rows[e.RowIndex].Cells["AddInTabName"].FormattedValue;
            UserUpdRowIndex += "," + e.RowIndex;
        }
        /// <summary>
        /// 
        /// </summary>
        private void BindData()
        {
            DataTable dt = ft.GetGroupsFromDB();
            this.dgv.DataSource = dt;
        }
        /// <summary>
        /// 
        /// </summary>
        private void BindDataUsers()
        {
            DataTable dt = ft.GetUsersFromDB();
            this.dgvUsers.DataSource = dt;
        }
        private int currentRowIndex = 0;
        private void DataGridView1_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                if (e.Button == MouseButtons.Right)
                {
                    if (e.RowIndex >= 0)
                    {
                        dgv.ClearSelection();
                        currentRowIndex = e.RowIndex;
                        this.contextMenuStrip1.Show(MousePosition.X, MousePosition.Y);
                    }
                }
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void contextMenuScript1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (currentRowIndex >= 0)
            {
                try
                {
                    DeleteGroupIDs += this.dgv.Rows[currentRowIndex].Cells["ID"].FormattedValue + ",";
                    dgv.Rows.RemoveAt(currentRowIndex);
                    CheckCurrentRow(currentRowIndex);
                    DataGridViewCellEventArgs args = new DataGridViewCellEventArgs(0, currentRowIndex);
                    dgv_CellClick(dgv, args);
                    btnApply.Enabled = true;
                    dgv.Rows[currentRowIndex].Selected = true;
                }
                catch { }
            }
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
                    dgv.RowHeadersWidth - 4,
                    e.RowBounds.Height);
                TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(),
                    dgv.RowHeadersDefaultCellStyle.Font,
                    rectangle,
                    dgv.RowHeadersDefaultCellStyle.ForeColor,
                    TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvUsers_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            try
            {
                System.Drawing.Rectangle rectangle = new System.Drawing.Rectangle(e.RowBounds.Location.X,
                    e.RowBounds.Location.Y,
                    dgvUsers.RowHeadersWidth - 4,
                    e.RowBounds.Height);
                TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(),
                    dgvUsers.RowHeadersDefaultCellStyle.Font,
                    rectangle,
                    dgvUsers.RowHeadersDefaultCellStyle.ForeColor,
                    TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvPermissions_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            try
            {
                System.Drawing.Rectangle rectangle = new System.Drawing.Rectangle(e.RowBounds.Location.X,
                    e.RowBounds.Location.Y,
                    dgvPermissions.RowHeadersWidth - 4,
                    e.RowBounds.Height);
                TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(),
                    dgvPermissions.RowHeadersDefaultCellStyle.Font,
                    rectangle,
                    dgvPermissions.RowHeadersDefaultCellStyle.ForeColor,
                    TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
            }
            catch { }
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
        private void btnApply_Click(object sender, EventArgs e)
        {
            Apply();
            btnApply.Enabled = false;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOK_Click(object sender, EventArgs e)
        {
            Apply();
            this.Close();
        }
        /// <summary>
        /// 
        /// </summary>
        private void Apply()
        {
            int selectRowIndex = 0;
            try
            {
                selectRowIndex = dgv.SelectedRows[0].Index;
            }
            catch { }
            DeleteGroup();
            SaveGroup();
            SaveUserAndPermission();
            SaveGroupUsersAndGroupPermissions();
            ft.uPDATETemplateUpdateFlag(true);
            clearTempString();
            BindData();
            dgv.Rows[selectRowIndex].Selected = true;
            dgv.CurrentCell = dgv.Rows[selectRowIndex].Cells[1];
            DataGridViewCellEventArgs args = new DataGridViewCellEventArgs(dgv.CurrentCell.ColumnIndex, dgv.CurrentRow.Index);
            dgv_CellClick(dgv, args);
        }
        /// <summary>
        /// 
        /// </summary>
        private void clearTempString()
        {
            PermissionsUpdRows = string.Empty;
            UserUpdRows = string.Empty;
            AddInTabValue = string.Empty;
            UserUpdRowIndex = string.Empty;
            PermissionNames = string.Empty;
            PermissionRemark = string.Empty;
            PermissionFolder = string.Empty;
            PermissionUpdRowIndex = string.Empty;
            GroupUpdRows = string.Empty;
            GroupNames = string.Empty;
            GroupDisable = string.Empty;
            GroupFolderMaskName = string.Empty;
            GroupRemark = string.Empty;
            GroupUpdRowIndex = string.Empty;
            DeleteGroupIDs = string.Empty;
        }
        /// <summary>
        /// 
        /// </summary>
        private void SaveGroupUsersAndGroupPermissions()
        {
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                if (this.dgv.SelectedRows.Count == 0) return;
                string groupid = this.dgv.SelectedRows[0].Cells["ID"].FormattedValue.ToString();
                SqlCommand cmd = null;
                DeleteGroupUsers(groupid, conn, cmd);
                cmd = new SqlCommand("rsUserGroup_Ins", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                for (int j = 0; j < dgvUsers.Rows.Count; j++)
                {
                    if (dgvUsers.Rows[j].Cells[1].Value == null) continue;
                    if (dgvUsers.Rows[j].Cells[1].Value.ToString() == "True")
                    {
                        if (dgvUsers.Rows[j].Cells["UserID"].Value == null) continue;
                        string userid = dgvUsers.Rows[j].Cells["UserID"].Value.ToString();
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(new SqlParameter("@UserID", userid));
                        cmd.Parameters.Add(new SqlParameter("@GroupID", groupid));
                        cmd.ExecuteNonQuery();
                        SaveUserTemplates(userid, conn, cmd);
                    }
                }
                SaveGroupPermissions(conn, cmd);
            }
            finally
            {
                if (conn != null)
                    conn.Close();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="userid"></param>
        /// <param name="conn"></param>
        /// <param name="cmd"></param>
        private void SaveUserTemplates(string userid, SqlConnection conn, SqlCommand cmd)
        {
            cmd = new SqlCommand("rsUsersTemplatesVisible_Ins", conn);
            cmd.CommandType = CommandType.StoredProcedure;
            for (int j = 0; j < dgvPermissions.Rows.Count; j++)
            {
                if (dgvPermissions.Rows[j].Cells[1].Value == null) continue;
                if (dgvPermissions.Rows[j].Cells[1].Value.ToString() == "True")
                {
                    if (dgvPermissions.Rows[j].Cells["ID"].Value == null) continue;
                    string templateid = dgvPermissions.Rows[j].Cells["TemplateID"].Value.ToString();
                    if (!UserTemplateExist(userid, templateid))
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(new SqlParameter("@TemplateID", templateid));
                        cmd.Parameters.Add(new SqlParameter("@UserID", userid));
                        cmd.Parameters.Add(new SqlParameter("@Visible", "True"));
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="userid"></param>
        /// <param name="templateid"></param>
        /// <returns></returns>
        private bool UserTemplateExist(string userid, string templateid)
        {
            string i = ft.GetVisibleByUserIDTemplateID(userid, templateid);
            if (!string.IsNullOrEmpty(i)) return true;
            else return false;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="cmd"></param>
        private void SaveGroupPermissions(SqlConnection conn, SqlCommand cmd)
        {
            string groupid = this.dgv.SelectedRows[0].Cells["ID"].FormattedValue.ToString();
            DeleteGroupPermissions(groupid, conn, cmd);
            cmd = new SqlCommand("rsGroupPermissions_Ins", conn);
            cmd.CommandType = CommandType.StoredProcedure;
            for (int j = 0; j < dgvPermissions.Rows.Count; j++)
            {
                if (dgvPermissions.Rows[j].Cells[1].Value == null) continue;
                if (dgvPermissions.Rows[j].Cells[1].Value.ToString() == "True")
                {
                    if (dgvPermissions.Rows[j].Cells["ID"].Value == null) continue;
                    string permissionid = dgvPermissions.Rows[j].Cells["ID"].Value.ToString();
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add(new SqlParameter("@PermissionID", permissionid));
                    cmd.Parameters.Add(new SqlParameter("@GroupID", groupid));
                    cmd.Parameters.Add(new SqlParameter("@PermissionGroupName", ""));
                    cmd.ExecuteNonQuery();
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        private void DeleteGroup()
        {
            string[] sArray = Regex.Split(DeleteGroupIDs, ",", RegexOptions.IgnoreCase);
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsGroups_Del", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                for (int i = 0; i < sArray.Length; i++)
                {
                    if (string.IsNullOrEmpty(sArray[i])) continue;
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add(new SqlParameter("@id", sArray[i]));
                    cmd.ExecuteNonQuery();
                    DeleteGroupPermissions(sArray[i], conn, cmd);
                    DeleteGroupUsers(sArray[i], conn, cmd);
                }
            }
            finally
            {
                if (conn != null)
                    conn.Close();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="id"></param>
        /// <param name="conn"></param>
        /// <param name="cmd"></param>
        private void DeleteGroupPermissions(string id, SqlConnection conn, SqlCommand cmd)
        {
            cmd = new SqlCommand("rsGroupPermissions_Del", conn);
            cmd.CommandType = CommandType.StoredProcedure;
            if (string.IsNullOrEmpty(id)) return;
            cmd.Parameters.Clear();
            cmd.Parameters.Add(new SqlParameter("@GroupID", id));
            cmd.ExecuteNonQuery();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="id"></param>
        /// <param name="conn"></param>
        /// <param name="cmd"></param>
        private void DeleteGroupUsers(string id, SqlConnection conn, SqlCommand cmd)
        {
            cmd = new SqlCommand("rsUserGroup_Del", conn);
            cmd.CommandType = CommandType.StoredProcedure;
            if (string.IsNullOrEmpty(id)) return;
            cmd.Parameters.Clear();
            cmd.Parameters.Add(new SqlParameter("@GroupID", id));
            cmd.ExecuteNonQuery();
        }
        /// <summary>
        /// 
        /// </summary>
        private void SaveGroup()
        {
            string[] sArray = Regex.Split(GroupUpdRows, ",", RegexOptions.IgnoreCase);
            string[] sGroupNames = Regex.Split(GroupNames, ",", RegexOptions.IgnoreCase);
            string[] sGroupDisable = Regex.Split(GroupDisable, ",", RegexOptions.IgnoreCase);
            string[] sGroupFolderMaskName = Regex.Split(GroupFolderMaskName, ",", RegexOptions.IgnoreCase);
            string[] sGroupRemark = Regex.Split(GroupRemark, ",", RegexOptions.IgnoreCase);
            string[] sGroupUpdRowIndex = Regex.Split(GroupUpdRowIndex, ",", RegexOptions.IgnoreCase);
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                for (int i = 0; i < sArray.Length; i++)
                {
                    if (string.IsNullOrEmpty(sArray[i]) && (!string.IsNullOrEmpty(sGroupNames[i]) || !string.IsNullOrEmpty(sGroupFolderMaskName[i]) || !string.IsNullOrEmpty(sGroupRemark[i])))
                    {
                        SqlCommand cmd = new SqlCommand("rsGroups_Ins", conn);
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(new SqlParameter("@GroupName", sGroupNames[i]));
                        cmd.Parameters.Add(new SqlParameter("@GroupDisable", sGroupDisable[i]));
                        cmd.Parameters.Add(new SqlParameter("@Remark", sGroupRemark[i]));
                        cmd.Parameters.Add(new SqlParameter("@AddInTabName", sGroupFolderMaskName[i]));
                        SqlParameter parReturn = new SqlParameter("@ReturnValue", SqlDbType.Int);
                        parReturn.Direction = ParameterDirection.ReturnValue;
                        cmd.Parameters.Add(parReturn);
                        cmd.ExecuteNonQuery();
                        sArray[i] = parReturn.Value.ToString();
                        this.dgv.Rows[int.Parse(sGroupUpdRowIndex[i])].Cells["ID"].Value = parReturn.Value.ToString();
                    }
                    else if (!string.IsNullOrEmpty(sArray[i]))
                    {
                        SqlCommand cmd = new SqlCommand("rsGroups_Upd", conn);
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(new SqlParameter("@id", sArray[i]));
                        cmd.Parameters.Add(new SqlParameter("@GroupName", sGroupNames[i]));
                        cmd.Parameters.Add(new SqlParameter("@GroupDisable", sGroupDisable[i]));
                        cmd.Parameters.Add(new SqlParameter("@Remark", sGroupRemark[i]));
                        cmd.Parameters.Add(new SqlParameter("@AddInTabName", sGroupFolderMaskName[i]));
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            finally
            {
                if (conn != null)
                    conn.Close();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        private void SaveUserAndPermission()
        {
            string[] sArray = Regex.Split(UserUpdRows, ",", RegexOptions.IgnoreCase);
            string[] sAddInTabValue = Regex.Split(AddInTabValue, ",", RegexOptions.IgnoreCase);
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsUsers_AddInTabName_Upd", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                for (int i = 0; i < sArray.Length; i++)
                {
                    if (string.IsNullOrEmpty(sArray[i])) continue;
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add(new SqlParameter("@id", sArray[i]));
                    cmd.Parameters.Add(new SqlParameter("@AddInTabName", sAddInTabValue[i]));
                    cmd.ExecuteNonQuery();
                }
                SavePermission(conn, cmd);
            }
            finally
            {
                if (conn != null)
                    conn.Close();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="cmd"></param>
        private void SavePermission(SqlConnection conn, SqlCommand cmd)
        {
            string[] sArray = Regex.Split(PermissionsUpdRows, ",", RegexOptions.IgnoreCase);
            string[] sPermissionNames = Regex.Split(PermissionNames, ",", RegexOptions.IgnoreCase);
            string[] sPermissionRemark = Regex.Split(PermissionRemark, ",", RegexOptions.IgnoreCase);
            string[] sPermissionFolder = Regex.Split(PermissionFolder, ",", RegexOptions.IgnoreCase);
            cmd = new SqlCommand("rsPermissions_Upd", conn);
            cmd.CommandType = CommandType.StoredProcedure;
            for (int i = 0; i < sArray.Length; i++)
            {
                if (string.IsNullOrEmpty(sArray[i])) continue;
                cmd.Parameters.Clear();
                cmd.Parameters.Add(new SqlParameter("@id", sArray[i]));
                cmd.Parameters.Add(new SqlParameter("@PermissionName", sPermissionNames[i]));
                cmd.Parameters.Add(new SqlParameter("@remark", sPermissionRemark[i]));
                cmd.Parameters.Add(new SqlParameter("@Folder", sPermissionFolder[i]));
                cmd.ExecuteNonQuery();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = ToDataTable(dtsearch.Select(" PermissionName like '%" + this.txtSearch.Text + "%' or Folder like '%" + this.txtSearch.Text + "%'"));
                if (dt == null) return;
                CurrencyManager currencyManager1 = (CurrencyManager)BindingContext[dgvPermissions.DataSource];
                currencyManager1.SuspendBinding();
                for (int j = 0; j < dgvPermissions.Rows.Count; j++)
                {
                    if (dgvPermissions.Rows[j].Cells["ID"].Value == null) continue;
                    dgvPermissions.Rows[j].Visible = false;
                }
                currencyManager1.ResumeBinding();
                for (int i = 0; i < dt.Rows.Count; i++)
                    for (int j = 0; j < dgvPermissions.Rows.Count; j++)
                    {
                        if (dgvPermissions.Rows[j].Cells["ID"].Value == null) continue;
                        if (dgvPermissions.Rows[j].Cells["ID"].Value.ToString() == dt.Rows[i][0].ToString())
                        {
                            dgvPermissions.Rows[j].Visible = true;
                        }
                    }
            }
            catch (Exception ex) { LogHelper.WriteLog(typeof(Security), ex.Message + " Security error"); }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="rows"></param>
        /// <returns></returns>
        public DataTable ToDataTable(DataRow[] rows)
        {
            if (rows == null || rows.Length == 0) return null;
            DataTable tmp = rows[0].Table.Clone();
            foreach (DataRow row in rows)
                tmp.ImportRow(row);
            return tmp;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtSearch_Leave(object sender, EventArgs e)
        {
            if (txtSearch.Text == "")
                txtSearch.Text = "Search:";
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtSearch_Click(object sender, EventArgs e)
        {
            if (txtSearch.Text == "Search:")
                txtSearch.Text = txtSearch.Text.Replace("Search:", "");
        }
    }
}
