using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Diagnostics;
using ExcelAddIn3.Common;

namespace ExcelAddIn3
{
    public partial class ReportCriteriaPane : UserControl
    {
        DataGridView dgv = null;
        /// <summary>
        /// 
        /// </summary>
        internal static Finance_Tools ft
        {
            get { return new Finance_Tools(); }
        }
        public ReportCriteriaPane()
        {
            try
            {
                InitializeComponent();
                dgv = ft.IniGrd();
                BindReports();
                BindCriteria1();
                BindCriteria2();
                BindCriteria3();
                BindCriteria4();
                BindCriteria5();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Report Criteria Panel Error");
                //EventLog.WriteEntry("Finance Tool", ex.Message + "Report Criteria Pane error", EventLogEntryType.Error);
                LogHelper.WriteLog(typeof(ReportCriteriaPane), ex.Message + "Report Criteria Panel Error");
            }
        }
        /// <summary>
        /// 
        /// </summary>
        private void BindCriteria1()
        {
            ft.BindDropdowns(comboBox1, dgv);
        }
        /// <summary>
        /// 
        /// </summary>
        private void BindCriteria2()
        {
            ft.BindDropdowns(comboBox2, dgv);
        }
        /// <summary>
        /// 
        /// </summary>
        private void BindCriteria3()
        {
            ft.BindDropdowns(comboBox3, dgv);
        }
        /// <summary>
        /// 
        /// </summary>
        private void BindCriteria4()
        {
            ft.BindDropdowns(comboBox4, dgv);
        }
        /// <summary>
        /// 
        /// </summary>
        private void BindCriteria5()
        {
            ft.BindDropdowns(comboBox5, dgv);
        }

        /// <summary>
        /// 
        /// </summary>
        private void BindReports()
        {
            this.comboBox6.DisplayMember = "name";
            this.comboBox6.ValueMember = "value";
            this.comboBox6.DataSource = ft.InitialReportTemplates();
        }
        private string getName(int i)
        {
            switch (i)
            {
                case 0:
                    return string.IsNullOrEmpty(this.comboBox1.Text) ? "" : this.comboBox1.Text;
                case 1:
                    return string.IsNullOrEmpty(this.comboBox2.Text) ? "" : this.comboBox2.Text;
                case 2:
                    return string.IsNullOrEmpty(this.comboBox3.Text) ? "" : this.comboBox3.Text;
                case 3:
                    return string.IsNullOrEmpty(this.comboBox4.Text) ? "" : this.comboBox4.Text;
                case 4:
                    return string.IsNullOrEmpty(this.comboBox5.Text) ? "" : this.comboBox5.Text;
                default:
                    return "";
            }
        }
        private void setName(int i, string name)
        {
            switch (i)
            {
                case 0:
                    if (!string.IsNullOrEmpty(name))
                    {
                        this.comboBox1.SelectedText = name;
                    }
                    break;
                case 1:
                    if (!string.IsNullOrEmpty(name))
                    {
                        this.comboBox2.SelectedText = name;
                    }
                    break;
                case 2:
                    if (!string.IsNullOrEmpty(name))
                    {
                        this.comboBox3.SelectedText = name;
                    }
                    break;
                case 3:
                    if (!string.IsNullOrEmpty(name))
                    {
                        this.comboBox4.SelectedText = name;
                    }
                    break;
                case 4:
                    if (!string.IsNullOrEmpty(name))
                    {
                        this.comboBox5.SelectedText = name;
                    }
                    break;
                default:
                    break;
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection conn = null;
            SqlDataReader rdr = null;

            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("FT_Settings_ReportSetting_Del", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@path", this.comboBox6.SelectedValue.ToString()));
                //cmd.Parameters.Add(new SqlParameter("@UserID", SessionInfo.UserInfo.ID));
                rdr = cmd.ExecuteReader();

                for (int i = 0; i < 5; i++)
                {
                    if (!string.IsNullOrEmpty(getName(i)))
                    {
                        // 1. create a command object identifying
                        // the stored procedure
                        SqlCommand cmd2 = new SqlCommand("FT_Settings_ReportSetting_Ins", conn);

                        // 2. set the command object so it knows
                        // to execute a stored procedure
                        cmd2.CommandType = CommandType.StoredProcedure;
                        // 3. add parameter to command, which
                        // will be passed to the stored procedure

                        cmd2.Parameters.Add(new SqlParameter("@TemplatePath", this.comboBox6.SelectedValue.ToString()));
                        cmd2.Parameters.Add(new SqlParameter("@CriteriaName", getName(i)));
                        //cmd2.Parameters.Add(new SqlParameter("@UserID", SessionInfo.UserInfo.ID));

                        // execute the command
                        rdr.Close();
                        rdr = cmd2.ExecuteReader();
                    }
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
            //Ribbon1._MyReportCriteriaCustomTaskPane.Visible = false;
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = ft.GetReportCriteria(this.comboBox6.SelectedValue.ToString());
            comboBox1.Text = "";
            comboBox2.Text = "";
            comboBox3.Text = "";
            comboBox4.Text = "";
            comboBox5.Text = "";

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                setName(i, dt.Rows[i]["CriteriaName"].ToString());
            }
        }
    }
}
