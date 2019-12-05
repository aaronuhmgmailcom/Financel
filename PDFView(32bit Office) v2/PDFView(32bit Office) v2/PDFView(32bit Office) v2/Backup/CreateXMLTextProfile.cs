/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<CreateXMLTextProfile>   
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
using System.Xml;
using System.Text.RegularExpressions;
using ExcelAddIn4.Common;

namespace ExcelAddIn4
{
    public partial class CreateXMLTextProfile : Form
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
        public CreateXMLTextProfile()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 
        /// </summary>
        private void InitializeRows()
        {
            TextBox cbsunField;
            TextBox cbFriendName;
            ComboBox cbVisible;
            TextBox cbDefaultValue;
            CheckBox ckMandatory;
            TextBox Separator;
            TextBox txtTextLength;
            ComboBox trimText;
            TextBox txtPrefix;
            TextBox txtSuffix;
            TextBox txtRemoveCharactor;
            for (int i = 0; i < 20; i++)
            {
                cbsunField = new TextBox();
                cbsunField.Location = new System.Drawing.Point(0, 35 * (i));
                cbsunField.Size = new System.Drawing.Size(110, 22);
                cbsunField.Name = "cbsunField";
                cbsunField.Text = "";
                cbFriendName = new TextBox();
                cbFriendName.Location = new System.Drawing.Point(128, 35 * (i));
                cbFriendName.Size = new System.Drawing.Size(120, 22);
                cbFriendName.Name = "cbfriendlyName";
                cbFriendName.Text = "";
                cbVisible = new ComboBox();
                cbVisible.Location = new System.Drawing.Point(264, 35 * (i));
                cbVisible.Size = new System.Drawing.Size(80, 22);
                cbVisible.Name = "cbVisible";
                cbVisible.Items.Add("True");
                cbVisible.Items.Add("False");
                cbVisible.Text = "True";
                cbDefaultValue = new TextBox();
                cbDefaultValue.Location = new System.Drawing.Point(364, 35 * (i));
                cbDefaultValue.Size = new System.Drawing.Size(120, 22);
                cbDefaultValue.Name = "cbDefaultValue";
                ckMandatory = new CheckBox();
                ckMandatory.Location = new System.Drawing.Point(520, 35 * (i));
                ckMandatory.Name = "ckMandatory";
                ckMandatory.Size = new System.Drawing.Size(50, 22);
                Separator = new TextBox();
                Separator.Location = new System.Drawing.Point(580, 35 * (i));
                Separator.Size = new System.Drawing.Size(100, 22);
                Separator.Name = "Separator";
                txtTextLength = new TextBox();
                txtTextLength.Location = new System.Drawing.Point(710, 35 * (i));
                txtTextLength.Size = new System.Drawing.Size(80, 22);
                txtTextLength.Name = "txtTextLength";
                trimText = new ComboBox();
                trimText.Location = new System.Drawing.Point(820, 35 * (i));
                trimText.Size = new System.Drawing.Size(80, 22);
                trimText.Items.Add("Left");
                trimText.Items.Add("Right");
                trimText.Items.Add("None");
                trimText.Name = "trimText";
                txtPrefix = new TextBox();
                txtPrefix.Location = new System.Drawing.Point(920, 35 * (i));
                txtPrefix.Size = new System.Drawing.Size(80, 22);
                txtPrefix.Name = "txtPrefix";
                txtSuffix = new TextBox();
                txtSuffix.Location = new System.Drawing.Point(1030, 35 * (i));
                txtSuffix.Size = new System.Drawing.Size(80, 22);
                txtSuffix.Name = "txtSuffix";
                txtRemoveCharactor = new TextBox();
                txtRemoveCharactor.Location = new System.Drawing.Point(1130, 35 * (i));
                txtRemoveCharactor.Size = new System.Drawing.Size(110, 22);
                txtRemoveCharactor.Name = "txtRemoveCharactor";
                this.panel5.Controls.Add(cbsunField);
                this.panel5.Controls.Add(cbFriendName);
                this.panel5.Controls.Add(cbVisible);
                this.panel5.Controls.Add(cbDefaultValue);
                this.panel5.Controls.Add(ckMandatory);
                this.panel5.Controls.Add(Separator);
                this.panel5.Controls.Add(txtTextLength);
                this.panel5.Controls.Add(trimText);
                this.panel5.Controls.Add(txtPrefix);
                this.panel5.Controls.Add(txtSuffix);
                this.panel5.Controls.Add(txtRemoveCharactor);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbCommand.Text == "Create Manually")
            {
                this.panel5.Controls.Clear();
                lblFilePath.Visible = false;
                txtpath.Visible = false;
                btnBrowse.Visible = false;
                btnUpload.Visible = false;
                lblTextFileName.Visible = true;
                txtTextFileName.Visible = true;
                txtTextFileName.Location = new System.Drawing.Point(319, 12);
                cbTextFileName.Visible = false;
                lblseparator.Visible = false;
                cbseparator.Visible = false;
                txtSeparator.Visible = false;
                InitializeRows();
            }
            else if (cbCommand.Text == "Upload From File")
            {
                this.panel5.Controls.Clear();
                lblFilePath.Visible = true;
                txtpath.Visible = true;
                btnBrowse.Visible = true;
                btnUpload.Visible = true;
                lblTextFileName.Visible = true;
                cbTextFileName.Visible = false;
                txtTextFileName.Visible = true;
                txtTextFileName.Location = new System.Drawing.Point(319, 12);
                lblseparator.Visible = true;
                cbseparator.Visible = true;
                if (cbseparator.Text == "Other")
                    txtSeparator.Visible = true;
                else
                    txtSeparator.Visible = false;
            }
            else if (cbCommand.Text == "Select an Existing Profile")
            {
                this.panel5.Controls.Clear();
                lblTextFileName.Visible = true;
                cbTextFileName.Visible = true;
                txtTextFileName.Visible = false;
                lblFilePath.Visible = false;
                txtpath.Visible = false;
                btnBrowse.Visible = false;
                btnUpload.Visible = false;
                lblseparator.Visible = false;
                cbseparator.Visible = false;
                txtSeparator.Visible = false;
                BindExistTextFile();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        private void BindExistTextFile()
        {
            cbTextFileName.Items.Clear();
            List<string> list = ft.GetTextFiles();
            for (int i = 0; i < list.Count; i++)
                cbTextFileName.Items.Add(list[i]);
        }
        /// <summary>
        /// 
        /// </summary>
        private void BindExistXMLFile()
        {
            cbComponentName.Items.Clear();
            List<string> list = ft.GetXMLFiles();
            for (int i = 0; i < list.Count; i++)
                if (!cbComponentName.Items.Contains(list[i].Substring(0, list[i].LastIndexOf(","))))
                    cbComponentName.Items.Add(list[i].Substring(0, list[i].LastIndexOf(",")));
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbCommand2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbCommand2.Text == "Select an Existing Profile")
            {
                this.panel8.Controls.Clear();
                lblPath2.Visible = false;
                txtPath2.Visible = false;
                btnBrowser2.Visible = false;
                btnUpload2.Visible = false;
                lblComponent2.Visible = true;
                lblMethod2.Visible = true;
                cbComponentName.Visible = true;
                cbMethodName.Visible = true;
                txtComponentName.Visible = false;
                txtMethod.Visible = false;
                BindExistXMLFile();
            }
            else if (cbCommand2.Text == "Upload From File")
            {
                this.panel8.Controls.Clear();
                lblPath2.Visible = true;
                txtPath2.Visible = true;
                btnBrowser2.Visible = true;
                btnUpload2.Visible = true;
                lblComponent2.Visible = true;
                lblMethod2.Visible = true;
                cbComponentName.Visible = false;
                cbMethodName.Visible = false;
                txtComponentName.Visible = true;
                txtMethod.Visible = true;
                txtComponentName.Location = new System.Drawing.Point(345, 19);
                txtMethod.Location = new System.Drawing.Point(591, 19);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnUpload_Click(object sender, EventArgs e)
        {
            if (cbCommand.Text == "Upload From File")
            {
                this.errorProvider1.Clear();
                if (string.IsNullOrEmpty(txtTextFileName.Text))
                {
                    this.errorProvider1.SetError(this.txtTextFileName, "Specify text file name for the file."); return;
                }
                if (string.IsNullOrEmpty(cbseparator.Text))
                {
                    this.errorProvider1.SetError(this.cbseparator, "Specify separator for the file content."); return;
                }
                if (string.IsNullOrEmpty(txtSeparator.Text) && cbseparator.Text == "Other")
                {
                    this.errorProvider1.SetError(this.txtSeparator, "Specify separator for the file content."); return;
                }
                if (string.IsNullOrEmpty(txtpath.Text))
                {
                    this.errorProvider1.SetError(this.txtpath, "Specify the path for the file from which to upload the data. For example, C:\\File.txt, \\Sales\\Northwind\\File.txt. Or, click Browse."); return;
                }
                StreamReader sr = new StreamReader(txtpath.Text);
                string txt = sr.ReadLine();
                sr.Close();
                if (!txt.Contains(txtSeparator.Text))
                {
                    MessageBox.Show("Separator does not exist in the file."); return;
                }
                if (ft.FileExist(txtTextFileName.Text))
                {
                    if (MessageBox.Show("The text file name exists, Are you sure to continue this operation?", "Alert - RSystems FinanceTools", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                == DialogResult.Yes)
                    {
                        DelExistFile(txtTextFileName.Text, 0);
                        InsertContent(txt, txtTextFileName.Text, 0, "");
                        InsertFileFromUpload(txt, txtTextFileName.Text, 0);
                    }
                }
                else
                {
                    InsertContent(txt, txtTextFileName.Text, 0, "");
                    InsertFileFromUpload(txt, txtTextFileName.Text, 0);
                }
                this.panel5.Controls.Clear();
                DataTable dt = ft.GetFileData(txtTextFileName.Text);
                ShowTextFields(dt);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dt"></param>
        private void ShowTextFields(DataTable dt)
        {
            TextBox cbsunField;
            TextBox cbFriendName;
            ComboBox cbVisible;
            TextBox cbDefaultValue;
            CheckBox ckMandatory;
            TextBox Separator;
            TextBox txtTextLength;
            ComboBox trimText;
            TextBox txtPrefix;
            TextBox txtSuffix;
            TextBox txtRemoveCharactor;
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    cbsunField = new TextBox();
                    cbsunField.Location = new System.Drawing.Point(0, 35 * (i));
                    cbsunField.Size = new System.Drawing.Size(110, 22);
                    cbsunField.Name = "cbsunField";
                    cbsunField.Text = dt.Rows[i]["Field"].ToString();
                    cbFriendName = new TextBox();
                    cbFriendName.Location = new System.Drawing.Point(128, 35 * (i));
                    cbFriendName.Size = new System.Drawing.Size(120, 22);
                    cbFriendName.Name = "cbfriendlyName";
                    cbFriendName.Text = dt.Rows[i]["FriendlyName"].ToString();
                    cbVisible = new ComboBox();
                    cbVisible.Location = new System.Drawing.Point(264, 35 * (i));
                    cbVisible.Size = new System.Drawing.Size(80, 22);
                    cbVisible.Name = "cbVisible";
                    cbVisible.Items.Add("True");
                    cbVisible.Items.Add("False");
                    cbVisible.Text = dt.Rows[i]["Visible"].ToString();
                    cbDefaultValue = new TextBox();
                    cbDefaultValue.Location = new System.Drawing.Point(364, 35 * (i));
                    cbDefaultValue.Size = new System.Drawing.Size(120, 22);
                    cbDefaultValue.Name = "cbDefaultValue";
                    cbDefaultValue.Text = dt.Rows[i]["DefaultValue"].ToString();
                    ckMandatory = new CheckBox();
                    ckMandatory.Location = new System.Drawing.Point(520, 35 * (i));
                    ckMandatory.Name = "ckMandatory";
                    ckMandatory.Size = new System.Drawing.Size(50, 22);
                    ckMandatory.Checked = (bool)(dt.Rows[i]["Mandatory"]);
                    Separator = new TextBox();
                    Separator.Location = new System.Drawing.Point(580, 35 * (i));
                    Separator.Size = new System.Drawing.Size(100, 22);
                    Separator.Name = "Separator";
                    Separator.Text = dt.Rows[i]["Separator"].ToString();
                    txtTextLength = new TextBox();
                    txtTextLength.Location = new System.Drawing.Point(710, 35 * (i));
                    txtTextLength.Size = new System.Drawing.Size(80, 22);
                    txtTextLength.Name = "txtTextLength";
                    txtTextLength.Text = dt.Rows[i]["TextLength"].ToString();
                    trimText = new ComboBox();
                    trimText.Location = new System.Drawing.Point(820, 35 * (i));
                    trimText.Size = new System.Drawing.Size(80, 22);
                    trimText.Items.Add("Left");
                    trimText.Items.Add("Right");
                    trimText.Items.Add("None");
                    trimText.Name = "trimText";
                    trimText.Text = dt.Rows[i]["trimText"].ToString();
                    txtPrefix = new TextBox();
                    txtPrefix.Location = new System.Drawing.Point(920, 35 * (i));
                    txtPrefix.Size = new System.Drawing.Size(80, 22);
                    txtPrefix.Name = "txtPrefix";
                    txtPrefix.Text = dt.Rows[i]["Prefix"].ToString();
                    txtSuffix = new TextBox();
                    txtSuffix.Location = new System.Drawing.Point(1030, 35 * (i));
                    txtSuffix.Size = new System.Drawing.Size(80, 22);
                    txtSuffix.Name = "txtSuffix";
                    txtSuffix.Text = dt.Rows[i]["Suffix"].ToString();
                    txtRemoveCharactor = new TextBox();
                    txtRemoveCharactor.Location = new System.Drawing.Point(1130, 35 * (i));
                    txtRemoveCharactor.Size = new System.Drawing.Size(110, 22);
                    txtRemoveCharactor.Name = "txtRemoveCharactor";
                    txtRemoveCharactor.Text = dt.Rows[i]["RemoveCharacters"].ToString();
                    this.panel5.Controls.Add(cbsunField);
                    this.panel5.Controls.Add(cbFriendName);
                    this.panel5.Controls.Add(cbVisible);
                    this.panel5.Controls.Add(cbDefaultValue);
                    this.panel5.Controls.Add(ckMandatory);
                    this.panel5.Controls.Add(Separator);
                    this.panel5.Controls.Add(txtTextLength);
                    this.panel5.Controls.Add(trimText);
                    this.panel5.Controls.Add(txtPrefix);
                    this.panel5.Controls.Add(txtSuffix);
                    this.panel5.Controls.Add(txtRemoveCharactor);
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="filetype"></param>
        private void DelExistFile(string filename, int filetype)
        {
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = null;
                if (filetype == 0)
                {
                    cmd = new SqlCommand("rsTemplateCreateXMLTextProfile_Del", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@textfilename", filename));
                }
                else
                {
                    cmd = new SqlCommand("rsTemplateCreateXMLTextProfile_DelXML", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@SunComponentName", filename.Substring(0, filename.LastIndexOf(","))));
                    cmd.Parameters.Add(new SqlParameter("@SunMethod", filename.Substring(filename.LastIndexOf(",") + 1)));
                }
                cmd.ExecuteNonQuery();
                cmd = new SqlCommand("rsTemplateXMLTEXTFiles_Del", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@RelatedName", filename));
                cmd.Parameters.Add(new SqlParameter("@FileType", filetype));
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
        /// <param name="filecontent"></param>
        /// <param name="relatedname"></param>
        /// <param name="filetype"></param>
        private void InsertFileFromUpload(string filecontent, string relatedname, int filetype)
        {
            SqlConnection conn = null;
            SqlDataReader rdr = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsTemplateCreateXMLTextProfile_Ins", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                if (filetype == 0)
                {
                    string[] sArray = Regex.Split(filecontent, txtSeparator.Text);
                    for (int i = 0; i < sArray.Length; i++)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(new SqlParameter("@Field", sArray[i]));
                        cmd.Parameters.Add(new SqlParameter("@FriendlyName", sArray[i]));
                        cmd.Parameters.Add(new SqlParameter("@Visible", true));
                        cmd.Parameters.Add(new SqlParameter("@DefaultValue", ""));
                        cmd.Parameters.Add(new SqlParameter("@SunComponentName", ""));
                        cmd.Parameters.Add(new SqlParameter("@SunMethod", ""));
                        cmd.Parameters.Add(new SqlParameter("@Mandatory", false));
                        cmd.Parameters.Add(new SqlParameter("@Separator", txtSeparator.Text));
                        cmd.Parameters.Add(new SqlParameter("@TextLength", ""));
                        cmd.Parameters.Add(new SqlParameter("@trimText", "None"));
                        cmd.Parameters.Add(new SqlParameter("@Prefix", ""));
                        cmd.Parameters.Add(new SqlParameter("@Suffix", ""));
                        cmd.Parameters.Add(new SqlParameter("@RemoveCharacters", ""));
                        cmd.Parameters.Add(new SqlParameter("@TextFileName", relatedname));
                        cmd.Parameters.Add(new SqlParameter("@Parent", ""));
                        cmd.Parameters.Add(new SqlParameter("@Section", ""));
                        rdr = cmd.ExecuteReader();
                        rdr.Close();
                    }
                }
                else
                {
                    XmlRWrite xmlRWrite = new XmlRWrite(conn, rdr, cmd, txtComponentName.Text, txtMethod.Text);
                    xmlRWrite.FilePath = txtPath2.Text;
                    xmlRWrite.SaveNodeValue();
                }
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
        /// <param name="filecontent"></param>
        /// <param name="relatedname"></param>
        /// <param name="filetype"></param>
        /// <param name="xmlTemplate"></param>
        private void InsertContent(string filecontent, string relatedname, int filetype, string xmlTemplate)
        {
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsTemplateXMLTEXTFiles_Ins", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@FileContent", filecontent));
                cmd.Parameters.Add(new SqlParameter("@RelatedName", relatedname));
                cmd.Parameters.Add(new SqlParameter("@XMLTemplate", xmlTemplate));
                cmd.Parameters.Add(new SqlParameter("@FileType", filetype));
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
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "TEXT files (*.txt)|*.txt";
            fileDialog.FileName = "";
            fileDialog.Title = "Select a TEXT file to Upload";
            fileDialog.Multiselect = false;
            if (fileDialog.ShowDialog() == DialogResult.OK)
                txtpath.Text = fileDialog.FileName;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnBrowser2_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "XML files (*.xml)|*.xml";
            fileDialog.FileName = "";
            fileDialog.Title = "Select a XML file to Upload";
            fileDialog.Multiselect = false;
            if (fileDialog.ShowDialog() == DialogResult.OK)
                txtPath2.Text = fileDialog.FileName;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnUpload2_Click(object sender, EventArgs e)
        {
            if (cbCommand2.Text == "Upload From File")
            {
                this.errorProvider1.Clear();
                if (string.IsNullOrEmpty(txtComponentName.Text))
                {
                    this.errorProvider1.SetError(this.txtComponentName, "Specify component name for the XML."); return;
                }
                if (string.IsNullOrEmpty(txtMethod.Text))
                {
                    this.errorProvider1.SetError(this.txtMethod, "Specify method name for the XML."); return;
                }
                if (string.IsNullOrEmpty(txtPath2.Text))
                {
                    this.errorProvider1.SetError(this.txtPath2, "Specify the path for the XML from which to upload the data. For example, C:\\File.xml, \\Sales\\Northwind\\File.xml. Or, click Browse."); return;
                }
                StreamReader sr = new StreamReader(txtPath2.Text);
                string xml = sr.ReadToEnd();
                sr.Close();
                if (ft.FileExist(txtComponentName.Text + "," + txtMethod.Text))
                {
                    if (MessageBox.Show("The XML file name exists, Are you sure to continue this operation?", "Alert - RSystems FinanceTools", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                == DialogResult.Yes)
                    {
                        DelExistFile(txtComponentName.Text + "," + txtMethod.Text, 1);
                        InsertFileFromUpload(xml, txtComponentName.Text + "," + txtMethod.Text, 1);
                        InsertContent(xml, txtComponentName.Text + "," + txtMethod.Text, 1, xml);
                    }
                }
                else
                {
                    InsertFileFromUpload(xml, txtComponentName.Text + "," + txtMethod.Text, 1);
                    InsertContent(xml, txtComponentName.Text + "," + txtMethod.Text, 1, xml);
                }
                this.panel8.Controls.Clear();
                DataTable dt = ft.GetXMLData(txtComponentName.Text, txtMethod.Text);
                ShowXMLFields(dt);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbSection_SelectedIndexChanged(object sender, EventArgs e)
        {
            int number = int.Parse(((ComboBox)sender).Tag.ToString());
            string type = ((ComboBox)sender).Text;
            Control[] ctlsSec = this.panel8.Controls.Find("cbSection", true);
            if (type == "Header")
                for (int i = 0; i < number; i++)
                    ((ComboBox)ctlsSec[i]).Text = "Header";
            else if (type == "Footer")
                for (int i = number; i < ctlsSec.Length; i++)
                    ((ComboBox)ctlsSec[i]).Text = "Footer";
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dt"></param>
        private void ShowXMLFields(DataTable dt)
        {
            TextBox cbsunField;
            TextBox cbFriendName;
            TextBox cbParent;
            ComboBox cbSection;
            ComboBox cbVisible;
            TextBox cbDefaultValue;
            CheckBox ckMandatory;
            TextBox txtTextLength;
            ComboBox trimText;
            TextBox txtPrefix;
            TextBox txtSuffix;
            TextBox txtRemoveCharactor;
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    cbsunField = new TextBox();
                    cbsunField.Location = new System.Drawing.Point(0, 35 * (i));
                    cbsunField.Size = new System.Drawing.Size(110, 22);
                    cbsunField.Name = "cbsunField";
                    cbsunField.Text = dt.Rows[i]["Field"].ToString();
                    cbsunField.ReadOnly = true;
                    cbFriendName = new TextBox();
                    cbFriendName.Location = new System.Drawing.Point(120, 35 * (i));
                    cbFriendName.Size = new System.Drawing.Size(120, 22);
                    cbFriendName.Name = "cbfriendlyName";
                    cbFriendName.Text = dt.Rows[i]["FriendlyName"].ToString();
                    cbParent = new TextBox();
                    cbParent.Location = new System.Drawing.Point(260, 35 * (i));
                    cbParent.Size = new System.Drawing.Size(100, 22);
                    cbParent.Name = "cbParent";
                    cbParent.Text = dt.Rows[i]["Parent"].ToString();
                    cbParent.ReadOnly = true;
                    cbSection = new ComboBox();
                    cbSection.Location = new System.Drawing.Point(375, 35 * (i));
                    cbSection.Size = new System.Drawing.Size(60, 22);
                    cbSection.Name = "cbSection";
                    cbSection.Items.Add("Header");
                    cbSection.Items.Add("Line");
                    cbSection.Items.Add("Footer");
                    cbSection.Text = dt.Rows[i]["Section"].ToString();
                    cbSection.Tag = i;
                    cbSection.SelectedIndexChanged += new EventHandler(cbSection_SelectedIndexChanged);
                    cbVisible = new ComboBox();
                    cbVisible.Location = new System.Drawing.Point(455, 35 * (i));
                    cbVisible.Size = new System.Drawing.Size(60, 22);
                    cbVisible.Name = "cbVisible";
                    cbVisible.Items.Add("True");
                    cbVisible.Items.Add("False");
                    cbVisible.Text = dt.Rows[i]["Visible"].ToString();
                    cbDefaultValue = new TextBox();
                    cbDefaultValue.Location = new System.Drawing.Point(525, 35 * (i));
                    cbDefaultValue.Size = new System.Drawing.Size(110, 22);
                    cbDefaultValue.Name = "cbDefaultValue";
                    cbDefaultValue.Text = dt.Rows[i]["DefaultValue"].ToString();
                    ckMandatory = new CheckBox();
                    ckMandatory.Location = new System.Drawing.Point(675, 35 * (i));
                    ckMandatory.Name = "ckMandatory";
                    ckMandatory.Size = new System.Drawing.Size(50, 22);
                    ckMandatory.Checked = (bool)(dt.Rows[i]["Mandatory"]);
                    txtTextLength = new TextBox();
                    txtTextLength.Location = new System.Drawing.Point(730, 35 * (i));
                    txtTextLength.Size = new System.Drawing.Size(60, 22);
                    txtTextLength.Name = "txtTextLength";
                    txtTextLength.Text = dt.Rows[i]["TextLength"].ToString();
                    trimText = new ComboBox();
                    trimText.Location = new System.Drawing.Point(810, 35 * (i));
                    trimText.Size = new System.Drawing.Size(70, 22);
                    trimText.Name = "trimText";
                    trimText.Items.Add("Left");
                    trimText.Items.Add("Right");
                    trimText.Items.Add("None");
                    trimText.Text = dt.Rows[i]["trimText"].ToString();
                    txtPrefix = new TextBox();
                    txtPrefix.Location = new System.Drawing.Point(900, 35 * (i));
                    txtPrefix.Size = new System.Drawing.Size(80, 22);
                    txtPrefix.Name = "txtPrefix";
                    txtPrefix.Text = dt.Rows[i]["Prefix"].ToString();
                    txtSuffix = new TextBox();
                    txtSuffix.Location = new System.Drawing.Point(990, 35 * (i));
                    txtSuffix.Size = new System.Drawing.Size(80, 22);
                    txtSuffix.Name = "txtSuffix";
                    txtSuffix.Text = dt.Rows[i]["Suffix"].ToString();
                    txtRemoveCharactor = new TextBox();
                    txtRemoveCharactor.Location = new System.Drawing.Point(1080, 35 * (i));
                    txtRemoveCharactor.Size = new System.Drawing.Size(150, 22);
                    txtRemoveCharactor.Name = "txtRemoveCharactor";
                    txtRemoveCharactor.Text = dt.Rows[i]["RemoveCharacters"].ToString();
                    this.panel8.Controls.Add(cbsunField);
                    this.panel8.Controls.Add(cbFriendName);
                    this.panel8.Controls.Add(cbParent);
                    this.panel8.Controls.Add(cbSection);
                    this.panel8.Controls.Add(cbVisible);
                    this.panel8.Controls.Add(cbDefaultValue);
                    this.panel8.Controls.Add(ckMandatory);
                    this.panel8.Controls.Add(txtTextLength);
                    this.panel8.Controls.Add(trimText);
                    this.panel8.Controls.Add(txtPrefix);
                    this.panel8.Controls.Add(txtSuffix);
                    this.panel8.Controls.Add(txtRemoveCharactor);
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbseparator_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbseparator.Text == "Other")
                txtSeparator.Visible = true;
            else
            {
                txtSeparator.Visible = false;
                if (cbseparator.Text == "Comma")
                    txtSeparator.Text = ",";
                else if (cbseparator.Text == "Tab")
                    txtSeparator.Text = "\t";
                else if (cbseparator.Text == "Space")
                    txtSeparator.Text = " ";
                else if (cbseparator.Text == "Semicolon")
                    txtSeparator.Text = ";";
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbComponentName_SelectedIndexChanged(object sender, EventArgs e)
        {
            cbMethodName.Items.Clear();
            List<string> list = ft.GetXMLFiles();
            for (int i = 0; i < list.Count; i++)
                if (list[i].Substring(0, list[i].LastIndexOf(",")) == cbComponentName.Text)
                    cbMethodName.Items.Add(list[i].Substring(list[i].LastIndexOf(",") + 1));
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbTextFileName_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.panel5.Controls.Clear();
            DataTable dt = ft.GetFileData(cbTextFileName.Text);
            ShowTextFields(dt);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbMethodName_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.panel8.Controls.Clear();
            DataTable dt = ft.GetXMLData(cbComponentName.Text, cbMethodName.Text);
            ShowXMLFields(dt);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        private bool SaveManully(string filename)
        {
            this.errorProvider1.Clear();
            string content = string.Empty;
            if (string.IsNullOrEmpty(txtTextFileName.Text))
            {
                this.errorProvider1.SetError(this.txtTextFileName, "Specify text file name for the file."); return false;
            }
            if (ft.FileExist(txtTextFileName.Text))
            {
                if (MessageBox.Show("The text file name exists, Are you sure to continue this operation?", "Alert - RSystems FinanceTools", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            == DialogResult.Yes)
                {
                    DelExistFile(filename, 0);
                    InsertFileFromConfig(filename, ref content);
                    InsertContent(content, txtTextFileName.Text, 0, "");
                }
            }
            else
            {
                InsertFileFromConfig(filename, ref content);
                InsertContent(content, txtTextFileName.Text, 0, "");
            }
            return true;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private bool SaveUploadFromFile()
        {
            this.errorProvider1.Clear();
            string content = string.Empty;
            if (string.IsNullOrEmpty(txtTextFileName.Text))
            {
                this.errorProvider1.SetError(this.txtTextFileName, "Specify the text file name for the file."); return false;
            }
            DelExistFile(txtTextFileName.Text, 0);
            InsertFileFromConfig(txtTextFileName.Text, ref content);
            InsertContent(content, txtTextFileName.Text, 0, "");
            return true;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private bool SaveExistProfile()
        {
            string content = string.Empty;
            DelExistFile(cbTextFileName.Text, 0);
            InsertFileFromConfig(cbTextFileName.Text, ref content);
            InsertContent(content, cbTextFileName.Text, 0, "");
            return true;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="content"></param>
        private void InsertFileFromConfig(string filename, ref string content)
        {
            Control[] ctlsSF = this.panel5.Controls.Find("cbsunField", true);
            Control[] ctlsFN = this.panel5.Controls.Find("cbfriendlyName", true);
            Control[] ctlsV = this.panel5.Controls.Find("cbVisible", true);
            Control[] ctlsD = this.panel5.Controls.Find("cbDefaultValue", true);
            Control[] ctlsM = this.panel5.Controls.Find("ckMandatory", true);
            Control[] ctlsS = this.panel5.Controls.Find("Separator", true);
            Control[] ctlsT = this.panel5.Controls.Find("txtTextLength", true);
            Control[] ctlsTrim = this.panel5.Controls.Find("trimText", true);
            Control[] ctlsP = this.panel5.Controls.Find("txtPrefix", true);
            Control[] ctlsSU = this.panel5.Controls.Find("txtSuffix", true);
            Control[] ctlsR = this.panel5.Controls.Find("txtRemoveCharactor", true);
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsTemplateCreateXMLTextProfile_Ins", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                for (int i = 0; i < ctlsSF.Length; i++)
                {
                    string sunfield = ((TextBox)ctlsSF[i]).Text;
                    string friendname = ((TextBox)ctlsFN[i]).Text;
                    string visible = ((ComboBox)ctlsV[i]).Text;
                    string defaultv = ((TextBox)ctlsD[i]).Text;
                    bool mandatory = ((CheckBox)ctlsM[i]).Checked;
                    string separator = ((TextBox)ctlsS[i]).Text;
                    string textlength = ((TextBox)ctlsT[i]).Text;
                    string trimTxt = ((ComboBox)ctlsTrim[i]).Text;
                    string prefix = ((TextBox)ctlsP[i]).Text;
                    string suffix = ((TextBox)ctlsSU[i]).Text;
                    string remove = ((TextBox)ctlsR[i]).Text;
                    content += sunfield + separator;
                    if (string.IsNullOrEmpty(sunfield)) continue;
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add(new SqlParameter("@Field", sunfield));
                    cmd.Parameters.Add(new SqlParameter("@FriendlyName", friendname));
                    cmd.Parameters.Add(new SqlParameter("@Visible", visible));
                    cmd.Parameters.Add(new SqlParameter("@DefaultValue", defaultv));
                    cmd.Parameters.Add(new SqlParameter("@SunComponentName", ""));
                    cmd.Parameters.Add(new SqlParameter("@SunMethod", ""));
                    cmd.Parameters.Add(new SqlParameter("@Mandatory", mandatory));
                    cmd.Parameters.Add(new SqlParameter("@Separator", separator));
                    cmd.Parameters.Add(new SqlParameter("@TextLength", textlength));
                    cmd.Parameters.Add(new SqlParameter("@trimText", trimTxt));
                    cmd.Parameters.Add(new SqlParameter("@Prefix", prefix));
                    cmd.Parameters.Add(new SqlParameter("@Suffix", suffix));
                    cmd.Parameters.Add(new SqlParameter("@RemoveCharacters", remove));
                    cmd.Parameters.Add(new SqlParameter("@TextFileName", filename));
                    cmd.Parameters.Add(new SqlParameter("@Parent", ""));
                    cmd.Parameters.Add(new SqlParameter("@Section", ""));
                    cmd.ExecuteNonQuery();
                }
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
        private void button1_Click(object sender, EventArgs e)
        {
            bool returnValue = false;
            if (tabControl1.SelectedTab.Text == "Create Text File")
            {
                if (cbCommand.Text == "Create Manually")
                    returnValue = SaveManully(txtTextFileName.Text);
                else if (cbCommand.Text == "Upload From File")
                    returnValue = SaveUploadFromFile();
                else if (cbCommand.Text == "Select an Existing Profile")
                    returnValue = SaveExistProfile();
            }
            else
            {
                string returnXML = string.Empty;
                if (cbCommand2.Text == "Select an Existing Profile" && !string.IsNullOrEmpty(cbMethodName.Text))
                {
                    string originalXML = ft.GetXMLFileContent(cbComponentName.Text + "," + cbMethodName.Text);
                    returnValue = UpdateXML(cbComponentName.Text, cbMethodName.Text, ref returnXML);
                    InsertContent(originalXML, cbComponentName.Text + "," + cbMethodName.Text, 1, returnXML);
                }
                else if (cbCommand2.Text == "Upload From File" && !string.IsNullOrEmpty(txtComponentName.Text) && !string.IsNullOrEmpty(txtMethod.Text))
                {
                    string originalXML = ft.GetXMLFileContent(txtComponentName.Text + "," + txtMethod.Text);
                    returnValue = UpdateXML(txtComponentName.Text, txtMethod.Text, ref returnXML);
                    InsertContent(originalXML, txtComponentName.Text + "," + txtMethod.Text, 1, returnXML);
                }
            }
            if (returnValue) this.Close();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="comName"></param>
        /// <param name="methodName"></param>
        /// <param name="returnXML"></param>
        /// <returns></returns>
        private bool UpdateXML(string comName, string methodName, ref string returnXML)
        {
            returnXML = ft.GetXMLFileContent(comName + "," + methodName);
            XmlDocument xmlDoc = new XmlDocument();
            Control[] ctlsSF = this.panel8.Controls.Find("cbsunField", true);
            Control[] ctlsFN = this.panel8.Controls.Find("cbfriendlyName", true);
            Control[] ctlsPar = this.panel8.Controls.Find("cbParent", true);
            Control[] ctlsSec = this.panel8.Controls.Find("cbSection", true);
            Control[] ctlsV = this.panel8.Controls.Find("cbVisible", true);
            Control[] ctlsD = this.panel8.Controls.Find("cbDefaultValue", true);
            Control[] ctlsM = this.panel8.Controls.Find("ckMandatory", true);
            Control[] ctlsT = this.panel8.Controls.Find("txtTextLength", true);
            Control[] ctlsTrim = this.panel8.Controls.Find("trimText", true);
            Control[] ctlsP = this.panel8.Controls.Find("txtPrefix", true);
            Control[] ctlsSU = this.panel8.Controls.Find("txtSuffix", true);
            Control[] ctlsR = this.panel8.Controls.Find("txtRemoveCharactor", true);
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsTemplateCreateXMLTextProfile_DelXML", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@SunComponentName", comName));
                cmd.Parameters.Add(new SqlParameter("@SunMethod", methodName));
                cmd.ExecuteNonQuery();
                cmd = new SqlCommand("rsTemplateXMLTEXTFiles_Del", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@RelatedName", comName + "," + methodName));
                cmd.Parameters.Add(new SqlParameter("@FileType", 1));
                cmd.ExecuteNonQuery();
                cmd = new SqlCommand("rsTemplateCreateXMLTextProfile_Ins", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                for (int i = 0; i < ctlsSF.Length; i++)
                {
                    string sunfield = ((TextBox)ctlsSF[i]).Text;
                    string friendname = ((TextBox)ctlsFN[i]).Text;
                    string parent = ((TextBox)ctlsPar[i]).Text;
                    string section = ((ComboBox)ctlsSec[i]).Text;
                    string visible = ((ComboBox)ctlsV[i]).Text;
                    string defaultv = ((TextBox)ctlsD[i]).Text;
                    bool mandatory = ((CheckBox)ctlsM[i]).Checked;
                    string textlength = ((TextBox)ctlsT[i]).Text;
                    string trimTxt = ((ComboBox)ctlsTrim[i]).Text;
                    string prefix = ((TextBox)ctlsP[i]).Text;
                    string suffix = ((TextBox)ctlsSU[i]).Text;
                    string remove = ((TextBox)ctlsR[i]).Text;
                    xmlDoc.LoadXml(returnXML);
                    int c = ft.GetXMLorTextFileSameFieldCount(comName, methodName, sunfield);
                    XmlNode xn = xmlDoc.GetElementsByTagName(sunfield)[c];
                    if (xn.ChildNodes.Count == 0 || xn.LastChild.NodeType == XmlNodeType.Text)
                        xn.InnerText = defaultv;
                    if (!(xn.ChildNodes.Count >= 1 && (xn.ChildNodes[0].Name != "#text")))
                        if ((visible == "False") && (mandatory == false))//(section == "Line") &&
                            xn.ParentNode.RemoveChild(xn);
                    returnXML = xmlDoc.InnerXml;
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add(new SqlParameter("@Field", sunfield));
                    cmd.Parameters.Add(new SqlParameter("@FriendlyName", friendname));
                    cmd.Parameters.Add(new SqlParameter("@Visible", visible));
                    cmd.Parameters.Add(new SqlParameter("@DefaultValue", defaultv));
                    cmd.Parameters.Add(new SqlParameter("@SunComponentName", comName));
                    cmd.Parameters.Add(new SqlParameter("@SunMethod", methodName));
                    cmd.Parameters.Add(new SqlParameter("@Mandatory", mandatory));
                    cmd.Parameters.Add(new SqlParameter("@Separator", ""));
                    cmd.Parameters.Add(new SqlParameter("@TextLength", textlength));
                    cmd.Parameters.Add(new SqlParameter("@trimText", trimTxt));
                    cmd.Parameters.Add(new SqlParameter("@Prefix", prefix));
                    cmd.Parameters.Add(new SqlParameter("@Suffix", suffix));
                    cmd.Parameters.Add(new SqlParameter("@RemoveCharacters", remove));
                    cmd.Parameters.Add(new SqlParameter("@TextFileName", ""));
                    cmd.Parameters.Add(new SqlParameter("@Parent", parent));
                    cmd.Parameters.Add(new SqlParameter("@Section", section));
                    cmd.ExecuteNonQuery();
                }
                returnXML = ft.FormatXml(xmlDoc.InnerXml);
                return true;
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
