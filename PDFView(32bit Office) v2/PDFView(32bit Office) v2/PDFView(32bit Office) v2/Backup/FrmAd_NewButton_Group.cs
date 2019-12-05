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
using System.IO;

namespace ExcelAddIn4
{
    public partial class FrmAd_NewButton_Group : Form
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
        List<KeyValuePair<string, int>> group = new List<KeyValuePair<string, int>>();
        /// <summary>
        /// 
        /// </summary>
        List<KeyValuePair<string, string>> RemovedGroup = new List<KeyValuePair<string, string>>();
        /// <summary>
        /// 
        /// </summary>
        List<KeyValuePair<string, string>> buttonlist = new List<KeyValuePair<string, string>>();
        /// <summary>
        /// 
        /// </summary>
        private string NewButtonName
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        private object TemplateName
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        private object TemplateID
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public string newGroupName
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public int newGroupOrder
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        public int newButtonOrder
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="elem"></param>
        /// <returns></returns>
        private bool EquaWithName(KeyValuePair<string, int> elem)
        {
            if (elem.Key == tbGroupName.Text)
                return true;
            return false;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="elem"></param>
        /// <returns></returns>
        private bool EquaWithButtonName(KeyValuePair<string, string> elem)
        {
            if (elem.Key == NewButtonName)
                return true;
            return false;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="buttonName"></param>
        /// <param name="templateName"></param>
        /// <param name="templateID"></param>
        public FrmAd_NewButton_Group(string buttonName, string templateName, object templateID)
        {
            this.NewButtonName = buttonName;
            this.TemplateName = templateName;
            this.TemplateID = templateID;
            InitializeComponent();
            if (templateID != null)
            {
                DataTable dt = ft.GetTemplateButtons(templateID.ToString());
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string groupName = dt.Rows[i]["ButtonGroup"].ToString();
                    int grouporder = int.Parse(dt.Rows[i]["GroupOrder"].ToString());
                    tbGroupName.Text = groupName;
                    Predicate<KeyValuePair<string, int>> pred = EquaWithName;
                    if (!group.Exists(pred) && !string.IsNullOrEmpty(groupName))
                        group.Add(new KeyValuePair<string, int>(groupName, grouporder));
                }
                group.Add(new KeyValuePair<string, int>("<New Category...>", 1000000));
                IEnumerable<KeyValuePair<string, int>> query = group;
                foreach (KeyValuePair<string, int> item in query)
                    lbGroup.Items.Add(item);

                lbGroup.DisplayMember = "Key";
                lbGroup.ValueMember = "Value";
                lbGroup.SelectedIndexChanged += new EventHandler(lbGroup_SelectedIndexChanged);
                lbGroup_SelectedIndexChanged(null, null);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lbGroup_SelectedIndexChanged(object sender, EventArgs e)
        {
            string groupname = this.lbGroup.Text;
            if (groupname == "<New Category...>")
            {
                this.lblgroupname.Visible = true;
                this.tbGroupName.Visible = true;
                this.tbGroupName.Text = "";
                this.tbGroupName.Focus();
                btnSave.Text = "Add";
            }
            else
            {
                this.lblgroupname.Visible = false;
                this.tbGroupName.Visible = false;
                ShowGroupButtons();
                Predicate<KeyValuePair<string, string>> pred2 = EquaWithButtonName;
                if (buttonlist.Exists(pred2))
                    btnSave.Text = "Save";
                else
                    btnSave.Text = "Add";
            }
        }
        /// <summary>
        /// 
        /// </summary>
        private void ShowGroupButtons()
        {
            if (string.IsNullOrEmpty(lbGroup.Text)) return;
            DataTable dt = ft.GetGroupButtons(lbGroup.Text);
            buttonlist.Clear();
            lbButtons.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string buttonName = dt.Rows[i]["buttonName"].ToString();
                string groupname = dt.Rows[i]["buttonGroup"].ToString();
                if (!buttonlist.Contains(new KeyValuePair<string, string>(buttonName, groupname)) && !RemovedGroup.Contains(new KeyValuePair<string, string>(buttonName, groupname)))
                    buttonlist.Add(new KeyValuePair<string, string>(buttonName, groupname));
            }
            foreach (KeyValuePair<string, string> item in buttonlist)
                lbButtons.Items.Add(item);

            lbButtons.DisplayMember = "Key";
            lbButtons.ValueMember = "Value";
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="elem"></param>
        /// <returns></returns>
        private bool EquaWithTB(KeyValuePair<string, int> elem)
        {
            if (elem.Key == (this.TemplateName.ToString() + "_" + tbGroupName.Text))
                return true;
            return false;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, EventArgs e)
        {
            if (this.btnSave.Text == "Add" && this.tbGroupName.Visible == true)
            {
                if (string.IsNullOrEmpty(this.tbGroupName.Text))
                {
                    this.errorProvider1.SetError(this.tbGroupName, "Category Name can't be empty!");
                    return;
                }
                else
                {
                    this.errorProvider1.Clear();
                    Predicate<KeyValuePair<string, int>> pred = EquaWithTB;
                    if (group.Exists(pred))
                        MessageBox.Show("Category Name exist!", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else
                    {
                        group.Add(new KeyValuePair<string, int>(this.TemplateName.ToString() + "_" + tbGroupName.Text, 99999));
                        group.Remove(new KeyValuePair<string, int>("<New Category...>", 1000000));
                        group.Add(new KeyValuePair<string, int>("<New Category...>", 1000000));
                        lbGroup.Items.Add(new KeyValuePair<string, int>(this.TemplateName.ToString() + "_" + tbGroupName.Text, 99999));
                        lbGroup.Items.Remove(new KeyValuePair<string, int>("<New Category...>", 1000000));
                        lbGroup.Items.Add(new KeyValuePair<string, int>("<New Category...>", 1000000));
                        lbGroup.DisplayMember = "Key";
                        lbGroup.ValueMember = "Value";
                        lbGroup.SelectedIndex = lbGroup.Items.Count - 2;
                        Predicate<KeyValuePair<string, string>> pred2 = EquaWithButtonName;
                        if (buttonlist.Exists(pred2))
                            MessageBox.Show("Action Name exist!", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        else
                        {
                            this.lbButtons.Items.Add(new KeyValuePair<string, string>(this.NewButtonName, this.TemplateName.ToString() + "_" + tbGroupName.Text));
                            newGroupName = lbGroup.Text;
                            newGroupOrder = lbGroup.SelectedIndex;
                            btnSave.Text = "Save";
                        }
                    }
                }
            }
            else if (this.btnSave.Text == "Add" && this.tbGroupName.Visible == false)
            {
                btnSave.Text = "Save";
                Predicate<KeyValuePair<string, string>> pred = EquaWithButtonName;
                if (string.IsNullOrEmpty(lbGroup.Text))
                {
                    this.errorProvider1.SetError(this.lbGroup, "Select a Category!");
                    return;
                }
                else if (buttonlist.Exists(pred))
                    MessageBox.Show("Action Name exist!", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                {
                    this.lbButtons.Items.Add(new KeyValuePair<string, string>(this.NewButtonName, lbGroup.Text));
                    newGroupName = lbGroup.Text;
                    newGroupOrder = lbGroup.SelectedIndex;
                }
            }
            else if (this.btnSave.Text == "Save")
            {
                RemoveDeleteGroups();
                SaveButtons();
                SaveGroups();
                this.Hide();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        private void SaveButtons()
        {
            for (int i = 0; i < lbButtons.Items.Count; i++)
            {
                if (checkButtonExist((KeyValuePair<string, string>)lbButtons.Items[i]))
                    UpdateButtonOrder((KeyValuePair<string, string>)lbButtons.Items[i], i);
                else
                    newButtonOrder = i;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        private void SaveGroups()
        {
            for (int i = 0; i < lbGroup.Items.Count; i++)
            {
                if (checkGroupExist((KeyValuePair<string, int>)lbGroup.Items[i]))
                    UpdateGroupOrder((KeyValuePair<string, int>)lbGroup.Items[i], i);
                if (lbGroup.Items[i] == newGroupName)
                    newGroupOrder = i;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="item"></param>
        /// <param name="j"></param>
        private void UpdateButtonOrder(KeyValuePair<string, string> item, int j)
        {
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsTemplateActions_UpdActionOrder", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@ButtonGroup", item.Value));
                cmd.Parameters.Add(new SqlParameter("@ButtonOrder", j));
                cmd.Parameters.Add(new SqlParameter("@templateid", this.TemplateID));
                cmd.Parameters.Add(new SqlParameter("@buttonName", item.Key));
                cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();
                if (item.Key == NewButtonName)
                    newButtonOrder = j;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog(typeof(FrmAd_NewButton_Group), ex.Message + " - Data Error in Settings... Function, Create New Action Form !" + " Create New Action error");
                throw new Exception(ex.Message + " - Data Error in Settings... Function, Create New Action Form !");
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
        /// <param name="item"></param>
        /// <param name="j"></param>
        private void UpdateGroupOrder(KeyValuePair<string, int> item, int j)
        {
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsTemplateActions_UpdGroupOrder", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                DataTable dt = ft.GetGroupButtons(item.Key);
                if (dt.Rows.Count > 0)
                {
                    cmd.Parameters.Add(new SqlParameter("@ButtonGroup", item.Key));
                    cmd.Parameters.Add(new SqlParameter("@GroupOrder", j));
                    cmd.Parameters.Add(new SqlParameter("@OldButtonGroup", item.Key));
                    cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog(typeof(FrmAd_NewButton_Group), ex.Message + " - Data Error in Settings... Function, Create New Action Form !" + " Create New Action error");
                throw new Exception(ex.Message + " - Data Error in Settings... Function, Create New Action Form !");
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
        /// <param name="item"></param>
        /// <returns></returns>
        private bool checkButtonExist(KeyValuePair<string, string> item)
        {
            DataTable dt = ft.GetProcessMacroFromDB(this.TemplateID.ToString(), item.Key);
            if (dt.Rows.Count > 0)
                return true;
            else
                return false;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        private bool checkGroupExist(KeyValuePair<string, int> item)
        {
            DataTable dt = ft.GetGroupViaGroupName(this.TemplateID.ToString(), item.Key);
            if (dt.Rows.Count > 0)
                return true;
            else
                return false;
        }
        /// <summary>
        /// 
        /// </summary>
        private void RemoveDeleteGroups()
        {
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsTemplateActions_UpdGroupName", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                for (int i = 0; i < RemovedGroup.Count; i++)
                {
                    cmd.Parameters.Add(new SqlParameter("@ButtonGroup", ""));
                    cmd.Parameters.Add(new SqlParameter("@GroupOrder", -1));
                    cmd.Parameters.Add(new SqlParameter("@buttonName", RemovedGroup[i].Key));
                    cmd.Parameters.Add(new SqlParameter("@OldButtonGroup", RemovedGroup[i].Value));
                    cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog(typeof(FrmAd_NewButton_Group), ex.Message + " - Data Error in Settings... Function, Create New Action Form !" + " Create New Action error");
                throw new Exception(ex.Message + " - Data Error in Settings... Function, Create New Action Form !");
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
        private void lbGroup_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.Button == MouseButtons.Right)
                {
                    int posindex = lbGroup.IndexFromPoint(new Point(e.X, e.Y));
                    lbGroup.ContextMenuStrip = null;
                    if (posindex >= 0 && posindex < lbGroup.Items.Count)
                    {
                        lbGroup.SelectedIndex = posindex;
                        contextMenuStrip1.Show(lbGroup, new Point(e.X, e.Y));
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
        private void lbButtons_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.Button == MouseButtons.Right)
                {
                    int posindex = lbButtons.IndexFromPoint(new Point(e.X, e.Y));
                    lbButtons.ContextMenuStrip = null;
                    if (posindex >= 0 && posindex < lbButtons.Items.Count)
                    {
                        lbButtons.SelectedIndex = posindex;
                        contextMenuStrip2.Show(lbButtons, new Point(e.X, e.Y));
                    }
                }
            }
            catch { }
        }
        /// <summary>
        /// delete group
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            ListBox lbGroup = contextMenuStrip1.SourceControl as ListBox;
            if (lbGroup.Text == "<New Category...>")
                return;
            int i = lbGroup.SelectedIndex;
            for (int j = 0; j < lbButtons.Items.Count; j++)
            {
                KeyValuePair<string, string> obj = (KeyValuePair<string, string>)lbButtons.Items[j];
                RemovedGroup.Add((KeyValuePair<string, string>)lbButtons.Items[j]);
            }
            group.RemoveAt(i);
            lbGroup.Items.Remove(lbGroup.Items[i]);
            RemoveDeleteGroups();
            lbButtons.Items.Clear();
            this.btnSave.Text = "Save";
        }
        /// <summary>
        /// move up group
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (lbGroup.Text == "<New Category...>")
                return;
            if (this.lbGroup.SelectedIndices.Count > 0 &&
                  this.lbGroup.SelectedIndices[0] > 0)
            {
                int[] newIndices =
                    this.lbGroup.SelectedIndices.Cast<int>()
                    .Select(index => index - 1).ToArray();

                this.lbGroup.SelectedItems.Clear();

                for (int i = 0; i < newIndices.Length; i++)
                {
                    object obj = this.lbGroup.Items[newIndices[i]];
                    this.lbGroup.Items[newIndices[i]] = this.lbGroup.Items[newIndices[i] + 1];
                    this.lbGroup.Items[newIndices[i] + 1] = obj;
                    this.lbGroup.SelectedItems.Add(this.lbGroup.Items[newIndices[i]]);
                }
                SaveGroups();
                this.btnSave.Text = "Save";
            }
        }
        /// <summary>
        /// move down group
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (lbGroup.Text == "<New Category...>")
                return;
            if ((lbGroup.SelectedIndex + 1) == (lbGroup.Items.Count - 1)) return;
            if (this.lbGroup.SelectedIndices.Count > 0 &&
                     this.lbGroup.SelectedIndices[this.lbGroup.SelectedIndices.Count - 1] <
                     this.lbGroup.Items.Count - 1)
            {
                int[] newIndices =
                    this.lbGroup.SelectedIndices.Cast<int>()
                    .Select(index => index + 1).ToArray();

                this.lbGroup.SelectedItems.Clear();
                for (int i = newIndices.Length; i > 0; i--)
                {
                    object obj = this.lbGroup.Items[newIndices[i - 1]];
                    this.lbGroup.Items[newIndices[i - 1]] = this.lbGroup.Items[newIndices[i - 1] - 1];
                    this.lbGroup.Items[newIndices[i - 1] - 1] = obj;
                    this.lbGroup.SelectedItems.Add(this.lbGroup.Items[newIndices[i - 1]]);
                }
                SaveGroups();
                this.btnSave.Text = "Save";
            }
        }
        /// <summary>
        /// delete button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            ListBox lbButtons = contextMenuStrip2.SourceControl as ListBox;
            int i = lbButtons.SelectedIndex;
            RemovedGroup.Add((KeyValuePair<string, string>)lbButtons.Items[i]);
            lbButtons.Items.Remove(lbButtons.Items[i]);
            RemoveDeleteGroups();
            this.btnSave.Text = "Save";
        }
        /// <summary>
        /// move up button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            if (this.lbButtons.SelectedIndices.Count > 0 &&
               this.lbButtons.SelectedIndices[0] > 0)
            {
                int[] newIndices =
                    this.lbButtons.SelectedIndices.Cast<int>()
                    .Select(index => index - 1).ToArray();

                this.lbButtons.SelectedItems.Clear();
                for (int i = 0; i < newIndices.Length; i++)
                {
                    object obj = this.lbButtons.Items[newIndices[i]];
                    this.lbButtons.Items[newIndices[i]] = this.lbButtons.Items[newIndices[i] + 1];
                    this.lbButtons.Items[newIndices[i] + 1] = obj;
                    this.lbButtons.SelectedItems.Add(this.lbButtons.Items[newIndices[i]]);
                }
                SaveButtons();
                this.btnSave.Text = "Save";
            }
        }
        /// <summary>
        /// move down button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            if (this.lbButtons.SelectedIndices.Count > 0 &&
                     this.lbButtons.SelectedIndices[this.lbButtons.SelectedIndices.Count - 1] <
                     this.lbButtons.Items.Count - 1)
            {
                int[] newIndices =
                    this.lbButtons.SelectedIndices.Cast<int>()
                    .Select(index => index + 1).ToArray();

                this.lbButtons.SelectedItems.Clear();
                for (int i = newIndices.Length; i > 0; i--)
                {
                    object obj = this.lbButtons.Items[newIndices[i - 1]];
                    this.lbButtons.Items[newIndices[i - 1]] = this.lbButtons.Items[newIndices[i - 1] - 1];
                    this.lbButtons.Items[newIndices[i - 1] - 1] = obj;
                    this.lbButtons.SelectedItems.Add(this.lbButtons.Items[newIndices[i - 1]]);
                }
                SaveButtons();
                this.btnSave.Text = "Save";
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnReset_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FrmAd_NewButton_Group_FormClosed(object sender, FormClosedEventArgs e)
        {
            newGroupName = "-1";
        }
    }
}
