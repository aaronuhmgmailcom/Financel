/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<DataFieldsSetting>   
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
using System.Xml;
using System.Data.SqlClient;
using System.Configuration;

namespace ExcelAddIn4
{
    public partial class DataFieldsSetting : Form
    {
        /// <summary>
        /// 
        /// </summary>
        internal static Finance_Tools ft
        {
            get { return new Finance_Tools(); }
        }
        public string Tag;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Panel_MouseWheel(object sender, MouseEventArgs e)
        {
            if ((panel4.VerticalScroll.Value != panel4.VerticalScroll.Minimum) && (panel4.VerticalScroll.Value != panel4.VerticalScroll.Maximum))
                panel4.VerticalScroll.Value += 5;
            panel4.Refresh();
            panel4.Invalidate();
            panel4.Update();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="tag"></param>
        public DataFieldsSetting(string tag)
        {
            InitializeComponent();
            this.panel4.MouseWheel += new System.Windows.Forms.MouseEventHandler(this.Panel_MouseWheel);
            this.panel4.AutoScroll = true;
            DataTable dt = new DataTable(); ;
            XmlDocument xdoc = new XmlDocument();
            XmlNodeList nodeList = null;
            Tag = tag;
            if (tag == "Gen")
            {
                xdoc.Load("C:\\ProgramData\\RSDataV2\\RSDataConfig\\GenDescFieldsSetting.xml");
                nodeList = xdoc.SelectSingleNode("Nodes").ChildNodes;
                dt = ft.GetUserDataGenFriendlyName();
            }
            else
            {
                xdoc.Load("C:\\ProgramData\\RSDataV2\\RSDataConfig\\DataFieldsSetting.xml");
                nodeList = xdoc.SelectSingleNode("Nodes").ChildNodes;
                dt = ft.GetUserDataFriendlyName();
            }
            TextBox cbsunField;
            TextBox cbFriendName;
            ComboBox cbOutput;
            TextBox cbInput;
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    cbsunField = new TextBox();
                    cbsunField.Location = new System.Drawing.Point(50, 35 * (i));
                    cbsunField.Size = new System.Drawing.Size(110, 22);
                    cbsunField.Name = "cbsunField";
                    cbsunField.Text = dt.Rows[i]["SunField"].ToString();
                    cbsunField.ReadOnly = true;
                    cbFriendName = new TextBox();
                    cbFriendName.Location = new System.Drawing.Point(210, 35 * (i));
                    cbFriendName.Size = new System.Drawing.Size(110, 22);
                    cbFriendName.Name = "cbfriendlyName";
                    cbFriendName.Text = dt.Rows[i]["UserFriendlyName"].ToString();
                    cbOutput = new ComboBox();
                    cbOutput.Location = new System.Drawing.Point(350, 35 * (i));
                    cbOutput.Size = new System.Drawing.Size(110, 22);
                    cbOutput.Name = "cbOutput";
                    cbOutput.Items.Add("True");
                    cbOutput.Items.Add("False");
                    cbOutput.Text = dt.Rows[i]["Output"].ToString();
                    cbInput = new TextBox();
                    cbInput.Location = new System.Drawing.Point(500, 35 * (i));
                    cbInput.Size = new System.Drawing.Size(310, 22);
                    cbInput.Name = "cbinput";
                    cbInput.Text = dt.Rows[i]["XML_Query"].ToString();
                    this.panel4.Controls.Add(cbsunField);
                    this.panel4.Controls.Add(cbFriendName);
                    this.panel4.Controls.Add(cbOutput);
                    this.panel4.Controls.Add(cbInput);
                }
            }
            else
            {
                int i = 0;
                foreach (XmlNode xn in nodeList)
                {
                    XmlElement xe = xn as XmlElement;
                    string sunField = xe.GetAttribute("SunField");
                    string friendlyName = xe.GetAttribute("FriendlyName");
                    string output = xe.GetAttribute("Output");
                    string XML_Query = xe.GetAttribute("XML_Query");
                    cbsunField = new TextBox();
                    cbsunField.Location = new System.Drawing.Point(50, 35 * (i));
                    cbsunField.Size = new System.Drawing.Size(110, 22);
                    cbsunField.Name = "cbsunField";
                    cbsunField.Text = sunField;
                    cbsunField.ReadOnly = true;
                    cbFriendName = new TextBox();
                    cbFriendName.Location = new System.Drawing.Point(210, 35 * (i));
                    cbFriendName.Size = new System.Drawing.Size(110, 22);
                    cbFriendName.Name = "cbfriendlyName";
                    cbFriendName.Text = friendlyName;
                    cbOutput = new ComboBox();
                    cbOutput.Location = new System.Drawing.Point(350, 35 * (i));
                    cbOutput.Size = new System.Drawing.Size(110, 22);
                    cbOutput.Name = "cbOutput";
                    cbOutput.Items.Add("True");
                    cbOutput.Items.Add("False");
                    cbOutput.Text = output;
                    cbInput = new TextBox();
                    cbInput.Location = new System.Drawing.Point(500, 35 * (i));
                    cbInput.Size = new System.Drawing.Size(310, 22);
                    cbInput.Name = "cbinput";
                    cbInput.Text = XML_Query;
                    this.panel4.Controls.Add(cbsunField);
                    this.panel4.Controls.Add(cbFriendName);
                    this.panel4.Controls.Add(cbOutput);
                    this.panel4.Controls.Add(cbInput);
                    i++;
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
            Control[] ctlsSF = this.panel4.Controls.Find("cbsunField", true);
            Control[] ctlsFN = this.panel4.Controls.Find("cbfriendlyName", true);
            Control[] ctlsO = this.panel4.Controls.Find("cbOutput", true);
            Control[] ctlsI = this.panel4.Controls.Find("cbinput", true);
            SqlConnection conn = null;
            if (Tag == "Gen")
            {
                try
                {
                    conn = new
                        SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                    conn.Open();
                    SqlCommand cmd = new SqlCommand("rsTemplateGenDescFields_Del", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                    cmd.ExecuteNonQuery();
                    SqlCommand cmd2 = new SqlCommand("rsTemplateGenDescFields_Ins", conn);
                    cmd2.CommandType = CommandType.StoredProcedure;
                    for (int i = 0; i < ctlsSF.Length; i++)
                    {
                        string sunfield = ((TextBox)ctlsSF[i]).Text;
                        string friendname = ((TextBox)ctlsFN[i]).Text;
                        string output = ((ComboBox)ctlsO[i]).Text;
                        string xmlquery = ((TextBox)ctlsI[i]).Text;
                        cmd2.Parameters.Clear();
                        cmd2.Parameters.Add(new SqlParameter("@FieldGroup", ""));
                        cmd2.Parameters.Add(new SqlParameter("@SunField", sunfield));
                        cmd2.Parameters.Add(new SqlParameter("@UserFriendlyName", friendname));
                        cmd2.Parameters.Add(new SqlParameter("@Output", output));
                        cmd2.Parameters.Add(new SqlParameter("@Input", ""));
                        cmd2.Parameters.Add(new SqlParameter("@XML_Query", xmlquery));
                        cmd2.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                        cmd2.Parameters.Add(new SqlParameter("@version", i));
                        cmd2.ExecuteNonQuery();
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
            else
            {
                try
                {
                    conn = new
                        SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                    conn.Open();
                    SqlCommand cmd = new SqlCommand("rsGlobalFields_Del", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.ExecuteNonQuery();
                    SqlCommand cmd2 = new SqlCommand("rsGlobalFields_Ins", conn);
                    cmd2.CommandType = CommandType.StoredProcedure;
                    for (int i = 0; i < ctlsSF.Length; i++)
                    {
                        string sunfield = ((TextBox)ctlsSF[i]).Text;
                        string friendname = ((TextBox)ctlsFN[i]).Text;
                        string output = ((ComboBox)ctlsO[i]).Text;
                        string xmlquery = ((TextBox)ctlsI[i]).Text;
                        cmd2.Parameters.Clear();
                        cmd2.Parameters.Add(new SqlParameter("@FieldGroup", ""));
                        cmd2.Parameters.Add(new SqlParameter("@SunField", sunfield));
                        cmd2.Parameters.Add(new SqlParameter("@UserFriendlyName", friendname));
                        cmd2.Parameters.Add(new SqlParameter("@Output", output));
                        cmd2.Parameters.Add(new SqlParameter("@Input", ""));
                        cmd2.Parameters.Add(new SqlParameter("@XML_Query", xmlquery));
                        cmd2.Parameters.Add(new SqlParameter("@version", i));
                        cmd2.ExecuteNonQuery();
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
            Finance_Tools.AppSettingSave("isUseFriendlyName", "true");
            this.Close();
        }
    }
}
