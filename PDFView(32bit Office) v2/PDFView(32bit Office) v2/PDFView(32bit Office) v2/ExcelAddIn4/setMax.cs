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
    public partial class setMax : Form
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
        public setMax()
        {
            InitializeComponent();
            DataTable datatab = ft.GetReportCriteria(SessionInfo.UserInfo.File_ftid);
            cbprefix.DataSource = datatab;
            cbprefix.DisplayMember = "SequencePrefix";
            cbprefix.ValueMember = "SequencePrefix";
            cbprefix_SelectedIndexChanged(null,null);
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
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOK_Click(object sender, EventArgs e)
        {
            this.errorProvider2.Clear();
            int result;
            if (int.TryParse(cbMax.Text, out result))
            {
                if (result < ft.ProcessMaxNumber(cbprefix.SelectedValue.ToString()))
                {
                    this.errorProvider2.SetError(this.cbMax, "Not less than " + ft.ProcessMaxNumber(cbprefix.SelectedValue.ToString()) + " !");
                    return;
                }
                else
                {
                    SqlConnection conn = null;
                    try
                    {
                        conn = new
                            SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                        conn.Open();
                        SqlCommand cmd = new SqlCommand("rsTemplateTransactions_Ins", conn);
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add(new SqlParameter("@TemplateName", SessionInfo.UserInfo.FileName));
                        cmd.Parameters.Add(new SqlParameter("@Criteria1", "SetMax"));
                        cmd.Parameters.Add(new SqlParameter("@Criteria2", "SetMax"));
                        cmd.Parameters.Add(new SqlParameter("@Criteria3", "SetMax"));
                        cmd.Parameters.Add(new SqlParameter("@Criteria4", "SetMax"));
                        cmd.Parameters.Add(new SqlParameter("@Criteria5", "SetMax"));
                        cmd.Parameters.Add(new SqlParameter("@Value1", "SetMax"));
                        cmd.Parameters.Add(new SqlParameter("@Value2", "SetMax"));
                        cmd.Parameters.Add(new SqlParameter("@Value3", "SetMax"));
                        cmd.Parameters.Add(new SqlParameter("@Value4", "SetMax"));
                        cmd.Parameters.Add(new SqlParameter("@Value5", "SetMax"));
                        cmd.Parameters.Add(new SqlParameter("@Data", new byte[0]));
                        cmd.Parameters.Add(new SqlParameter("@DataType", "SetMax"));
                        cmd.Parameters.Add(new SqlParameter("@PDFData", new byte[0]));
                        cmd.Parameters.Add(new SqlParameter("@XMLData", "SetMax"));
                        cmd.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                        cmd.Parameters.Add(new SqlParameter("@maxNum", result));
                        cmd.Parameters.Add(new SqlParameter("@TransactionName", "SetMax"));
                        cmd.Parameters.Add(new SqlParameter("@Prefix", cbprefix.SelectedValue.ToString()));
                        cmd.Parameters.Add(new SqlParameter("@SunJournalNumber", ""));
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex) { throw new Exception(ex.ToString()); }
                    finally
                    {
                        if (conn != null)
                        {
                            conn.Close();
                        }
                    }
                    this.Close();
                }
            }
            else
            {
                this.errorProvider2.SetError(this.cbMax, "Input Error!");
                return;
            }
        }

        private void cbprefix_SelectedIndexChanged(object sender, EventArgs e)
        {
            var maxnum = ft.ProcessMaxNumber(cbprefix.SelectedValue.ToString());
            cbMax.Text = maxnum.ToString();
        }
    }
}
