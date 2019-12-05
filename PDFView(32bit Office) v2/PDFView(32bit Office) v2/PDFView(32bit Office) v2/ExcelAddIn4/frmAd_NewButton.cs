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
using System.Data.SqlClient;
using System.Configuration;
using ExcelAddIn4.Common;

namespace ExcelAddIn4
{
    public partial class frmAd_NewButton : Form
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
        internal ComboBox cbProcess = new ComboBox();
        /// <summary>
        /// 
        /// </summary>
        internal TextBox tbNewButton = new TextBox();
        /// <summary>
        /// 
        /// </summary>
        public static string iconName = string.Empty;
        /// <summary>
        /// 
        /// </summary>
        public static int size = 0;
        /// <summary>
        /// 
        /// </summary>
        private string GroupName
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        private int GroupOrder
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        private int ButtonOrder
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public frmAd_NewButton()
        {
            frmAd_NewButton.iconName = "";
            InitializeComponent();
            listBox1.Items.Clear();
            cbTemplateName.DataSource = Ribbon2.TemplateAndPath;
            cbTemplateName.DisplayMember = "Key";
            cbTemplateName.ValueMember = "Value";
            tbButtonName.DisplayMember = "Key";
            tbButtonName.ValueMember = "Value";
            InitializeToolTip();
            try
            {
                cbTemplateName.SelectedValue = SessionInfo.UserInfo.File_ftid;
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="actionid"></param>
        private void BindGroups(string actionid)
        {
            DataTable dt = ft.GetGroupsFromDB();
            clbGroup.DataSource = dt;
            clbGroup.ValueMember = "ID";
            clbGroup.DisplayMember = "GroupName";
            for (int i = 0; i < clbGroup.Items.Count; i++)
                clbGroup.SetItemCheckState(i, CheckState.Unchecked);
            if (!string.IsNullOrEmpty(actionid))
            {
                List<string> list = ft.GetGroupsByActionID(actionid, cbTemplateName.SelectedValue.ToString());
                for (int i = 0; i < clbGroup.Items.Count; i++)
                {
                    DataRowView dv = (DataRowView)clbGroup.Items[i];
                    string groupid = dv["ID"].ToString();
                    if (list.Contains(groupid))
                        clbGroup.SetItemCheckState(i, CheckState.Checked);
                }
            }
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
            toolTip1.SetToolTip(this.btnIcon, "Click here to search a perfect icon!");
            toolTip1.SetToolTip(this.btnGroup, "Click here to config your category!");
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCNB_Create_Click(object sender, EventArgs e)
        {
            this.errorProvider1.Clear();
            if (string.IsNullOrEmpty(tbButtonName.Text))
            {
                this.errorProvider1.SetError(this.tbButtonName, "Action Name can't be empty!");
                tabControl1.SelectedTab = tabPage1;
                return;
            }
            if (tbButtonName.SelectedValue == "-1" && string.IsNullOrEmpty(tbNewButton.Text))
            {
                this.errorProvider1.SetError(this.tbNewButton, "Action Name can't be empty!");
                tabControl1.SelectedTab = tabPage1;
                return;
            }
            if (btnGroup.Text == "Settings..." || string.IsNullOrEmpty(btnGroup.Text))
            {
                this.errorProvider1.SetError(this.btnGroup, "Please choose a category for the Action!");
                tabControl1.SelectedTab = tabPage1;
                return;
            }
            if (string.IsNullOrEmpty(frmAd_NewButton.iconName))
            {
                this.errorProvider1.SetError(this.btnIcon, "Please choose an icon for the Action!");
                tabControl1.SelectedTab = tabPage1;
                return;
            }
            if (listBox1.Items.Count == 0)
            {
                this.errorProvider1.SetError(this.listBox1, "Output/Save list can't be empty!");
                tabControl1.SelectedTab = tabPage2;
                return;
            }
            if (tbButtonName.SelectedValue == "-1" && ft.ButtonNameExist(cbTemplateName.SelectedValue.ToString(), tbNewButton.Text))
            {
                this.errorProvider1.SetError(this.tbNewButton, "Action Name exist!");
                tabControl1.SelectedTab = tabPage1;
                return;
            }
            try
            {
                if (btnCNB_Create.Text == "Update Action")
                {
                    DoDelete();
                }
                #region insert buttonprocessmacros table in DB
                for (int i = 0; i < listBox1.Items.Count; i++)
                {
                    KeyValuePair<string, string> obj = (KeyValuePair<string, string>)listBox1.Items[i];
                    if (obj.Value.Contains("Output Process") && obj.Key.Contains("Journal Post"))
                    {
                        string ProcessID = cbTemplateName.SelectedValue.ToString() + "," + obj.Key + "," + obj.Value.Substring(0, obj.Value.IndexOf(","));
                        InsertButton(ProcessID, "", i, "1", obj.Value.Substring(0, obj.Value.IndexOf(",")));
                    }
                    else if (obj.Value.Contains("Output Process") && obj.Key.Contains("Journal Update"))
                    {
                        string ProcessID = cbTemplateName.SelectedValue.ToString() + "," + obj.Key + "," + obj.Value.Substring(0, obj.Value.IndexOf(","));
                        InsertButton(ProcessID, "", i, "2", obj.Value.Substring(0, obj.Value.IndexOf(",")));
                    }
                    else if (obj.Value.Contains("Output Process") && obj.Key.Contains("Re-open template"))
                    {
                        InsertButton("", "", i, "6", "");
                    }
                    else if (obj.Value.Contains("Output Process") && !obj.Key.Contains("Journal Update") && !obj.Key.Contains("Journal Post") && !obj.Key.Contains("Re-open template"))
                    {
                        string ProcessID = cbTemplateName.SelectedValue.ToString() + "," + obj.Key + "," + obj.Value.Substring(0, obj.Value.IndexOf(","));
                        InsertButton(ProcessID, "", i, "3", obj.Value.Substring(0, obj.Value.IndexOf(",")));
                    }
                    else if (obj.Value.Contains(",Save PDF"))
                    {
                        InsertButton("", "", i, "7", obj.Value.Substring(0, obj.Value.IndexOf(",")));
                    }
                    else if (obj.Value.Contains(",Save"))
                    {
                        InsertButton("", "", i, "4", obj.Value.Substring(0, obj.Value.IndexOf(",")));
                    }
                    else if (obj.Value == "Output Macro")
                    {
                        InsertButton("", obj.Key, i, "5", "");
                    }
                }
                InsertPermission(this.cbTemplateName.Text, (btnCNB_Create.Text == "Update Action" ? this.tbButtonName.Text : this.tbNewButton.Text), this.cbTemplateName.SelectedValue.ToString());
                #endregion
                MessageBox.Show("Success!", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Ribbon2.InitializeNewButtons(SessionInfo.UserInfo.File_ftid, SessionInfo.UserInfo.FileName);
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LogHelper.WriteLog(typeof(frmAd_NewButton), ex.Message);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="processID"></param>
        /// <param name="macroName"></param>
        /// <param name="ProcessMacroOrder"></param>
        /// <param name="type"></param>
        /// <param name="reference"></param>
        private void InsertButton(string processID, string macroName, int ProcessMacroOrder, string type, string reference)
        {
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsTemplateActions_Ins", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                if (btnCNB_Create.Text == "Update Action")
                    cmd.Parameters.Add(new SqlParameter("@Name", this.tbButtonName.Text));
                else if (btnCNB_Create.Text == "Create New Action")
                    cmd.Parameters.Add(new SqlParameter("@Name", this.tbNewButton.Text));

                cmd.Parameters.Add(new SqlParameter("@Text", this.textBox1.Text.Replace("Display in action's SuperTip ,Click here!", "")));
                cmd.Parameters.Add(new SqlParameter("@ButtonIcon", frmAd_NewButton.iconName));
                cmd.Parameters.Add(new SqlParameter("@ButtonGroup", GroupName));
                cmd.Parameters.Add(new SqlParameter("@ButtonSize", frmAd_NewButton.size));
                cmd.Parameters.Add(new SqlParameter("@ButtonOrder", ButtonOrder));
                cmd.Parameters.Add(new SqlParameter("@GroupOrder", GroupOrder));
                cmd.Parameters.Add(new SqlParameter("@StopOnError", chkStop.Checked));
                cmd.Parameters.Add(new SqlParameter("@templateID", this.cbTemplateName.SelectedValue));
                cmd.Parameters.Add(new SqlParameter("@ProcessID", processID));
                cmd.Parameters.Add(new SqlParameter("@MacroName", macroName));
                cmd.Parameters.Add(new SqlParameter("@ProcessMacroOrder", ProcessMacroOrder));
                cmd.Parameters.Add(new SqlParameter("@Type", type));
                cmd.Parameters.Add(new SqlParameter("@Reference", reference.Replace(" ", "")));
                cmd.Parameters.Add(new SqlParameter("@ShowMsg", cbMsg.Checked));
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
        /// <param name="templateName"></param>
        /// <param name="permissionName"></param>
        /// <param name="templateid"></param>
        private void InsertPermission(string templateName, string permissionName, string templateid)
        {
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();

                SqlCommand cmd = new SqlCommand("rsPermissions_Del", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@templateid", templateid));
                cmd.Parameters.Add(new SqlParameter("@actionID", permissionName));
                cmd.ExecuteNonQuery();

                cmd = new SqlCommand("rsPermissions_Ins", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@PermissionName", templateName + "-" + permissionName + " - Create New Action"));
                cmd.Parameters.Add(new SqlParameter("@TemplateID", templateid));
                cmd.Parameters.Add(new SqlParameter("@ActionID", permissionName));
                cmd.Parameters.Add(new SqlParameter("@Per_Type", "4"));//0.Template - Save/Amend,1.Template - Write,2.Template - Read 3. global .4.create new action
                cmd.Parameters.Add(new SqlParameter("@remark", ""));
                cmd.Parameters.Add(new SqlParameter("@Folder", ""));
                SqlParameter parReturn = new SqlParameter("@ReturnValue", SqlDbType.Int);
                parReturn.Direction = ParameterDirection.ReturnValue;
                cmd.Parameters.Add(parReturn);
                cmd.ExecuteNonQuery();
                SaveGroupPermissions(conn, cmd, parReturn.Value.ToString());
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
        /// <param name="conn"></param>
        /// <param name="cmd"></param>
        /// <param name="permissionid"></param>
        private void SaveGroupPermissions(SqlConnection conn, SqlCommand cmd, string permissionid)
        {
            for (int i = 0; i < clbGroup.CheckedItems.Count; i++)
            {
                DataRowView dv = (DataRowView)clbGroup.CheckedItems[i];
                string groupid = dv["ID"].ToString();
                cmd = new SqlCommand("rsGroupPermissions_Ins", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Clear();
                cmd.Parameters.Add(new SqlParameter("@PermissionID", permissionid));
                cmd.Parameters.Add(new SqlParameter("@GroupID", groupid));
                cmd.Parameters.Add(new SqlParameter("@PermissionGroupName", ""));
                cmd.ExecuteNonQuery();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbFunctionType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbFunctionType.Text == "Output Process")
            {
                tbMacroName.Visible = false;
                cbReference.Visible = true;
                lblRef.Visible = true;
                cbProcess.FormattingEnabled = true;
                cbProcess.Location = new System.Drawing.Point(129, 67);
                cbProcess.Size = new System.Drawing.Size(194, 20);
                cbProcess.Name = "CB";
                cbProcess.DropDownStyle = ComboBoxStyle.DropDownList;
                cbProcess.Visible = true;
                cbProcess.SelectedIndexChanged += new EventHandler(cbProcess_SelectedIndexChanged);
                lbADV_FName.Visible = true;
                this.groupBox1.Controls.Add(cbProcess);
                DataTable dt = ft.GetProcessesFromDB(cbTemplateName.SelectedValue.ToString());
                if (dt.Rows.Count <= 0)
                {
                    dt.Columns.Add("id");
                    dt.Columns.Add("name");
                }
                dt.Rows.Add("-1", "Journal Post");
                dt.Rows.Add("-2", "Journal Update");
                dt.Rows.Add("-3", "Re-open template");
                cbProcess.DisplayMember = "name";
                cbProcess.DataSource = dt;
            }
            else if (cbFunctionType.Text == "Output Macro")
            {
                tbMacroName.Visible = true;
                cbProcess.Visible = false;
                lbADV_FName.Visible = true;
                cbReference.Visible = false;
                lblRef.Visible = false;
            }
            else
            {
                tbMacroName.Visible = false;
                cbProcess.Visible = false;
                lbADV_FName.Visible = false;
                cbReference.Visible = true;
                lblRef.Visible = true;
                DataTable dt = ft.GetSaveRefFromDB(cbTemplateName.SelectedValue.ToString());
                if (dt.Rows.Count <= 0)
                {
                    dt.Columns.Add("ft_id");
                    dt.Columns.Add("Reference");
                }
                cbReference.DisplayMember = "Reference";
                cbReference.DataSource = dt;
                cbReference.ValueMember = "ft_id";
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbProcess_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbProcess.Text == "Re-open template")
            {
                cbReference.Visible = false;
                lblRef.Visible = false;
            }
            else
            {
                DataTable dt = ft.GetProcessesRefFromDB(cbTemplateName.SelectedValue.ToString(), cbProcess.Text);
                cbReference.DataSource = dt;
                this.cbReference.DisplayMember = "reference";
                cbReference.Visible = true;
                lblRef.Visible = true;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listBox1_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                int posindex = listBox1.IndexFromPoint(new Point(e.X, e.Y));
                listBox1.ContextMenuStrip = null;
                if (posindex >= 0 && posindex < listBox1.Items.Count)
                {
                    listBox1.SelectedIndex = posindex;
                    contextMenuStrip1.Show(listBox1, new Point(e.X, e.Y));
                }
            }
            listBox1.Refresh();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAddProcMacro_Click(object sender, EventArgs e)
        {
            if ((tbMacroName.Visible == true) && string.IsNullOrEmpty(tbMacroName.Text))
            {
                this.errorProvider1.SetError(this.tbMacroName, "Macro Name can't be empty!"); return;
            }
            if (tbMacroName.Visible == true)
            {
                listBox1.Items.Add(new KeyValuePair<string, string>(tbMacroName.Text, "Output Macro"));
            }
            else if (cbProcess.Visible == true && cbReference.Visible == true)
            {
                if (string.IsNullOrEmpty(cbReference.Text))
                {
                    this.errorProvider1.SetError(this.cbReference, "Reference can't be empty!"); return;
                }
                else if (!string.IsNullOrEmpty(cbReference.Text))
                    listBox1.Items.Add(new KeyValuePair<string, string>(cbProcess.Text, cbReference.Text.Trim() + ",Output Process"));
            }
            else if (tbMacroName.Visible == false && cbReference.Visible == false)
                listBox1.Items.Add(new KeyValuePair<string, string>(cbProcess.Text, "Output Process"));
            else
            {
                if (string.IsNullOrEmpty(cbReference.Text))
                {
                    this.errorProvider1.SetError(this.cbReference, "Reference can't be empty!"); return;
                }
                if (cbFunctionType.Text == "Save")
                    listBox1.Items.Add(new KeyValuePair<string, string>("Save", cbReference.Text.Trim() + ",Save"));
                else if(cbFunctionType.Text=="Save PDF")
                    listBox1.Items.Add(new KeyValuePair<string, string>("Save PDF", cbReference.Text.Trim() + ",Save PDF"));
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ListBox listbox = contextMenuStrip1.SourceControl as ListBox;
            int i = listbox.SelectedIndex;
            listbox.Items.RemoveAt(i);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnUp_Click(object sender, EventArgs e)
        {
            int lbxLength = this.listBox1.Items.Count;
            int iselect = this.listBox1.SelectedIndex;
            if (lbxLength > iselect && iselect > 0)
            {
                object oTempItem = this.listBox1.SelectedItem;
                this.listBox1.Items.RemoveAt(iselect);
                this.listBox1.Items.Insert(iselect - 1, oTempItem);
                this.listBox1.SelectedIndex = iselect - 1;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDown_Click(object sender, EventArgs e)
        {
            int lbxLength = this.listBox1.Items.Count;
            int iselect = this.listBox1.SelectedIndex;
            if (lbxLength > iselect && iselect < lbxLength - 1)
            {
                object oTempItem = this.listBox1.SelectedItem;
                this.listBox1.Items.RemoveAt(iselect);
                this.listBox1.Items.Insert(iselect + 1, oTempItem);
                this.listBox1.SelectedIndex = iselect + 1;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "Display in action's SuperTip ,Click here!")
                textBox1.Text = textBox1.Text.Replace("Display in action's SuperTip ,Click here!", "");
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
                textBox1.Text = "Display in action's SuperTip ,Click here!";
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbTemplateName_SelectedIndexChanged(object sender, EventArgs e)
        {
            KeyValuePair<string, string> obj = (KeyValuePair<string, string>)this.cbTemplateName.SelectedItem;
            DataTable dt = ft.GetTemplateButtons(obj.Value);
            List<KeyValuePair<string, string>> tbButtons = new List<KeyValuePair<string, string>>();
            tbButtons.Add(new KeyValuePair<string, string>("", "-2"));
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string buttonName = dt.Rows[i]["name"].ToString();
                string buttonGroup = dt.Rows[i]["buttonGroup"].ToString();
                tbButtons.Add(new KeyValuePair<string, string>(buttonName, buttonGroup));
            }
            tbButtons.Add(new KeyValuePair<string, string>("<New Action...>", "-1"));
            tbButtonName.DataSource = tbButtons;
            tbButtonName.SelectedIndexChanged += new EventHandler(tbButtonName_SelectedIndexChanged);
            tbButtonName_SelectedIndexChanged(null, null);
            this.btnGroup.Text = "Settings...";
            this.btnIcon.Text = "Choose an Icon";
            this.btnIcon.Image = null;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tbButtonName_SelectedIndexChanged(object sender, EventArgs e)
        {
            KeyValuePair<string, string> obj = (KeyValuePair<string, string>)this.tbButtonName.SelectedItem;
            listBox1.Items.Clear();
            DataTable dt = new DataTable();
            try
            {
                dt = ft.GetProcessMacroFromDB(cbTemplateName.SelectedValue.ToString(), obj.Key);
            }
            catch { }
            if (obj.Value == "-1")
            {
                #region -1 means 'New Button'
                btnDelete.Visible = false;
                btnCNB_Create.Text = "Create New Action";
                tbButtonName.Visible = false;
                tbNewButton.Location = new System.Drawing.Point(106, 59);
                tbNewButton.Size = new System.Drawing.Size(194, 21);
                tbNewButton.Name = "TB";
                tbNewButton.Visible = true;
                tbNewButton.Text = "";
                this.groupBox3.Controls.Add(tbNewButton);
                textBox1.Text = "Display in action's SuperTip ,Click here!";
                #endregion
            }
            else if (obj.Value == "-2")
            {
                btnDelete.Visible = false;
                btnCNB_Create.Text = "Create New Action";
                tbButtonName.Visible = true;
                tbNewButton.Visible = false;
                textBox1.Text = "Display in action's SuperTip ,Click here!";
            }
            else if (obj.Value != "-1" && obj.Value != "-2")
            {
                #region !=-1 and !=-2 means user has choose an exist button
                btnDelete.Visible = true;
                btnCNB_Create.Text = "Update Action";
                tbButtonName.Visible = true;
                tbNewButton.Visible = false;
                if (dt.Rows.Count > 0)
                {
                    string text = dt.Rows[0]["ButtonText"].ToString(); //ft.GetButtonNameTextFromDB(obj.Key, cbTemplateName.SelectedValue.ToString());//text = text.Substring(text.IndexOf(",,,") + 3, text.Length - text.IndexOf(",,,") - 3);
                    this.textBox1.Text = text;
                    bool? stop = bool.Parse(dt.Rows[0]["StopOnError"].ToString()); //ft.GetButtonStopOnError(obj.Key, cbTemplateName.SelectedValue.ToString());
                    bool? showmsg = bool.Parse(dt.Rows[0]["ShowMsg"].ToString());
                    chkStop.Checked = (bool)stop;
                    cbMsg.Checked = (bool)showmsg;
                    string buttonIcon = dt.Rows[0]["ButtonIcon"].ToString();//ft.GetButtonIcon(obj.Key, cbTemplateName.SelectedValue.ToString());//string buttonIcon = buttonIconAndSize.Substring(0, buttonIconAndSize.IndexOf(",,,"));
                    int buttonSize = int.Parse(dt.Rows[0]["ButtonSize"].ToString());
                    btnIcon.Image = OleCreateConverter.PictureDispToImage(Globals.ThisAddIn.Application.CommandBars.GetImageMso(buttonIcon, buttonSize, buttonSize));
                    this.btnIcon.Text = "";
                    frmAd_NewButton.iconName = buttonIcon;
                    frmAd_NewButton.size = buttonSize;
                    this.btnGroup.Text = dt.Rows[0]["ButtonGroup"].ToString();
                    GroupName = btnGroup.Text;
                    GroupOrder = int.Parse(dt.Rows[0]["GroupOrder"].ToString());
                    ButtonOrder = int.Parse(dt.Rows[0]["ButtonOrder"].ToString());//ft.GetButtonOrder(obj.Key, cbTemplateName.SelectedValue.ToString()).Value;
                }
                #endregion
            }
            #region initialize button's process and
            List<KeyValuePair<string, string>> tbProcessMacros = new List<KeyValuePair<string, string>>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string processMacroNameType = dt.Rows[i]["Type"].ToString();
                string processMacroName = dt.Rows[i]["ProcessID"].ToString();
                string[] sArray = System.Text.RegularExpressions.Regex.Split(processMacroName, ",");
                if (processMacroNameType == "1" || processMacroNameType == "2" || processMacroNameType == "3")
                {
                    listBox1.Items.Add(new KeyValuePair<string, string>(sArray[1], sArray[2] + ",Output Process"));
                }
                else if (processMacroNameType == "6")
                {
                    listBox1.Items.Add(new KeyValuePair<string, string>("Re-open template", "Output Process"));
                }
                else if (processMacroNameType == "4")
                {
                    processMacroName = "Save";
                    string Ref = dt.Rows[i]["Reference"].ToString().Replace(" ", "");
                    listBox1.Items.Add(new KeyValuePair<string, string>(processMacroName, Ref + ",Save"));
                }
                else if (processMacroNameType == "7")
                {
                    processMacroName = "Save PDF";
                    string Ref = dt.Rows[i]["Reference"].ToString().Replace(" ", "");
                    listBox1.Items.Add(new KeyValuePair<string, string>(processMacroName, Ref + ",Save PDF"));
                }
                else if (processMacroNameType == "5")
                {
                    processMacroName = dt.Rows[i]["MacroName"].ToString();
                    listBox1.Items.Add(new KeyValuePair<string, string>(processMacroName, "Output Macro"));
                }
            }
            #endregion
            BindGroups(obj.Key);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDelete_Click(object sender, EventArgs e)
        {
            DoDelete();
            listBox1.Items.Clear();
            cbTemplateName_SelectedIndexChanged(null, null);
        }
        /// <summary>
        /// 
        /// </summary>
        private void DoDelete()
        {
            KeyValuePair<string, string> obj = (KeyValuePair<string, string>)this.tbButtonName.SelectedItem;
            DeleteTemplateButtons(obj.Key, cbTemplateName.SelectedValue.ToString());
            DeletePermissionButtons(obj.Key, cbTemplateName.SelectedValue.ToString());
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="buttonName"></param>
        /// <param name="templateid"></param>
        private void DeletePermissionButtons(string buttonName, string templateid)
        {
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsPermissions_Del", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@templateid", templateid));
                cmd.Parameters.Add(new SqlParameter("@actionID", buttonName));
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
        /// <param name="buttonName"></param>
        /// <param name="templateid"></param>
        private void DeleteTemplateButtons(string buttonName, string templateid)
        {
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsTemplateActions_Del", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@buttonName", buttonName));
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
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnGroup_Click(object sender, EventArgs e)
        {
            this.errorProvider1.Clear();
            if (string.IsNullOrEmpty(tbButtonName.Text))
            {
                this.errorProvider1.SetError(this.tbButtonName, "Action Name can't be empty!"); return;
            }
            if (tbButtonName.SelectedValue == "-1" && string.IsNullOrEmpty(tbNewButton.Text))
            {
                this.errorProvider1.SetError(this.tbNewButton, "Action Name can't be empty!"); return;
            }
            if (tbButtonName.SelectedValue == "-1" && ft.ButtonNameExist(cbTemplateName.SelectedValue.ToString(), tbNewButton.Text))
            {
                this.errorProvider1.SetError(this.tbNewButton, "Action Name exist!"); return;
            }
            FrmAd_NewButton_Group fng = new FrmAd_NewButton_Group(tbButtonName.SelectedValue == "-1" ? tbNewButton.Text : tbButtonName.Text, this.cbTemplateName.Text, this.cbTemplateName.SelectedValue);
            fng.ShowDialog();
            DataTable dt = ft.GetProcessMacroFromDB(cbTemplateName.SelectedValue.ToString(), tbButtonName.SelectedValue == "-1" ? tbNewButton.Text : tbButtonName.Text);
            if (fng.newGroupName != "-1")
            {
                if (dt.Rows.Count > 0)
                {
                    GroupName = dt.Rows[0]["ButtonGroup"].ToString();
                    GroupOrder = int.Parse(dt.Rows[0]["GroupOrder"].ToString());
                }
                else
                {
                    GroupName = fng.newGroupName;
                    GroupOrder = fng.newGroupOrder;
                }
                this.btnGroup.Text = GroupName;
                ButtonOrder = fng.newButtonOrder;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnIcon_Click(object sender, EventArgs e)
        {
            FrmAd_NewButton_Icon fni = new FrmAd_NewButton_Icon();
            fni.ShowDialog();
            if (!string.IsNullOrEmpty(frmAd_NewButton.iconName))
            {
                btnIcon.Image = OleCreateConverter.PictureDispToImage(Globals.ThisAddIn.Application.CommandBars.GetImageMso(frmAd_NewButton.iconName, frmAd_NewButton.size, frmAd_NewButton.size));
                this.btnIcon.Text = "";
            }
        }
    }
}
