/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<RunReport>   
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
using ExcelAddIn4.Common;

namespace ExcelAddIn4
{
    public partial class RunReport : Form
    {
        internal static Criterias cs;
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
        public RunReport()
        {
            InitializeComponent();
            InitialReportTemplates();
        }
        /// <summary>
        /// 
        /// </summary>
        private void InitialCriterias()
        {
            try
            {
                this.panel3.Controls.Clear();
                Label l;
                TextBox cb;
                if (!string.IsNullOrEmpty(this.comboBox1.Text))
                {
                    RSFinanceToolsEntities db = new RSFinanceToolsEntities();
                    string p = this.comboBox1.SelectedValue.ToString();
                    var dt = (from FT_sett in db.rsTemplateSettings
                              where FT_sett.TemplateID == p
                              select FT_sett).ToList();
                    DataTable newdt = ft.ToDataTable(dt);
                    var query = from t in newdt.AsEnumerable()
                                group t by new { t1 = t.Field<string>("CriteriaName") } into m
                                select new
                                {
                                    CriteriaName = m.First().Field<string>("CriteriaName"),
                                };
                    int i = 0;
                    foreach (var employee in query)
                    {
                        l = new Label();
                        l.AutoSize = true;
                        l.Location = new System.Drawing.Point(28, 36 * (i));
                        l.Name = "CriteriaLabel";
                        l.Size = new System.Drawing.Size(89, 13);
                        l.Text = employee.CriteriaName;
                        cb = new TextBox();
                        cb.Location = new System.Drawing.Point(199, 36 * (i));
                        cb.Size = new System.Drawing.Size(325, 66);
                        cb.Name = "CriteriaCB";
                        this.panel3.Controls.Add(l);
                        this.panel3.Controls.Add(cb);
                        i++;
                    }
                }
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        private void InitialReportTemplates()
        {
            this.comboBox1.DisplayMember = "name";
            this.comboBox1.ValueMember = "value";
            this.comboBox1.DataSource = ft.InitialReportTemplates();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            InitialCriterias();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            SaveReport();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private Criterias GetCriterias()
        {
            Criterias cris = new Criterias();
            cris.CriteriaN = new List<Criteria>();
            cris.TemplatePath = this.comboBox1.SelectedValue.ToString();

            Control[] ctls = this.panel3.Controls.Find("CriteriaLabel", true);
            for (int i = 0; i < 5; i++)
            {
                Criteria cri = new Criteria();
                cri.CriteriaName = new List<string>();
                cri.CriteriaValue = new List<string>();
                Control[] ctls2 = this.panel3.Controls.Find("CriteriaCB", true);
                try
                {
                    if (string.IsNullOrEmpty(((TextBox)ctls2[i]).Text))
                    {
                        cri.CriteriaName.Add("");
                        cri.CriteriaValue.Add("");
                    }
                    else
                    {
                        cri.CriteriaName.Add(((Label)ctls[i]).Text);
                        cri.CriteriaValue.Add(((TextBox)ctls2[i]).Text);
                    }
                }
                catch
                {
                    cri.CriteriaName.Add("");
                    cri.CriteriaValue.Add("");
                }
                cris.CriteriaN.Add(cri);
            }
            return cris;
        }
        /// <summary>
        /// 
        /// </summary>
        private void SaveReport()
        {
            try
            {
                var xlapp = Globals.ThisAddIn.Application;
                xlapp.Run("'" + SessionInfo.UserInfo.FilePath + "'!RSystems.runProc");
            }
            catch { }
            finally
            {
                cs = GetCriterias();
                SaveBinary();
                this.Close();
                this.Dispose();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        private void SaveBinary()
        {
            try
            {
                Criterias cris = RunReport.cs;
                SqlConnection conn = null;
                List<string> list = new List<string>();
                try
                {
                    conn = new
                        SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                    conn.Open();
                    SqlCommand cmd = new SqlCommand("FT_Reports_Ins", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    for (int j = 0; j < cris.CriteriaN[0].CriteriaName.Count; j++)
                    {
                        for (int k = 0; k < cris.CriteriaN[1].CriteriaName.Count; k++)
                        {
                            for (int l = 0; l < cris.CriteriaN[2].CriteriaName.Count; l++)
                            {
                                for (int m = 0; m < cris.CriteriaN[3].CriteriaName.Count; m++)
                                {
                                    for (int n = 0; n < cris.CriteriaN[4].CriteriaName.Count; n++)
                                    {
                                        if (!ft.CheckFileExist(cris.CriteriaN[0].CriteriaName[j], cris.CriteriaN[0].CriteriaValue[j], cris.CriteriaN[1].CriteriaName[k], cris.CriteriaN[1].CriteriaValue[k], cris.CriteriaN[2].CriteriaName[l], cris.CriteriaN[2].CriteriaValue[l], cris.CriteriaN[3].CriteriaName[m], cris.CriteriaN[3].CriteriaValue[m], cris.CriteriaN[4].CriteriaName[n], cris.CriteriaN[4].CriteriaValue[n], cris.TemplatePath))
                                        {
                                            cmd.Parameters.Add(new SqlParameter("@Criteria1", cris.CriteriaN[0].CriteriaName[j]));
                                            cmd.Parameters.Add(new SqlParameter("@Value1", cris.CriteriaN[0].CriteriaValue[j]));
                                            cmd.Parameters.Add(new SqlParameter("@Criteria2", cris.CriteriaN[1].CriteriaName[k]));
                                            cmd.Parameters.Add(new SqlParameter("@Value2", cris.CriteriaN[1].CriteriaValue[k]));
                                            cmd.Parameters.Add(new SqlParameter("@Criteria3", cris.CriteriaN[2].CriteriaName[l]));
                                            cmd.Parameters.Add(new SqlParameter("@Value3", cris.CriteriaN[2].CriteriaValue[l]));
                                            cmd.Parameters.Add(new SqlParameter("@Criteria4", cris.CriteriaN[3].CriteriaName[m]));
                                            cmd.Parameters.Add(new SqlParameter("@Value4", cris.CriteriaN[3].CriteriaValue[m]));
                                            cmd.Parameters.Add(new SqlParameter("@Criteria5", cris.CriteriaN[4].CriteriaName[n]));
                                            cmd.Parameters.Add(new SqlParameter("@Value5", cris.CriteriaN[4].CriteriaValue[n]));
                                            cmd.Parameters.Add(new SqlParameter("@TemplateName", Path.GetFileNameWithoutExtension(cris.TemplatePath)));
                                            cmd.Parameters.Add(new SqlParameter("@Data", ft.GetData(cris.TemplatePath)));
                                            cmd.Parameters.Add(new SqlParameter("@DataType", Path.GetExtension(cris.TemplatePath)));
                                            cmd.Parameters.Add(new SqlParameter("@PDFData", ft.GetData("")));
                                            cmd.Parameters.Add(new SqlParameter("@XMLData", ""));
                                            cmd.Parameters.Add(new SqlParameter("@TemplatePath", cris.TemplatePath));
                                            cmd.ExecuteNonQuery();
                                            cmd.Parameters.Clear();
                                        }
                                        else
                                        {
                                            list.Add(cris.CriteriaN[0].CriteriaName[j] + " " + cris.CriteriaN[0].CriteriaValue[j] + " " + cris.CriteriaN[1].CriteriaName[k] + " " + cris.CriteriaN[1].CriteriaValue[k] + " " + cris.CriteriaN[2].CriteriaName[l] + " " + cris.CriteriaN[2].CriteriaValue[l] + " " + cris.CriteriaN[3].CriteriaName[m] + " " + cris.CriteriaN[3].CriteriaValue[m] + " " + cris.CriteriaN[4].CriteriaName[n] + " " + cris.CriteriaN[4].CriteriaValue[n]);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if (list.Count != 0)
                    {
                        string str = "Below reports are exist , please check them and change the report certerias.\r\n \r\n";
                        foreach (string s in list)
                        {
                            str += s;
                            str += "\r\n \r\n";
                        }
                        MessageBox.Show(str);
                        LogHelper.WriteLog(typeof(RunReport), str);
                    }
                    else
                    {
                        MessageBox.Show("Success!", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    RunReport.cs.CriteriaN.Clear();
                }
                finally
                {
                    if (conn != null)
                    {
                        conn.Close();
                    }
                }
            }
            catch { }
        }
    }
}
