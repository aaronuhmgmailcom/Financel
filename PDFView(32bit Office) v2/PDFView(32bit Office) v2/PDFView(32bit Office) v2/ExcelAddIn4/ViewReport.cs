/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<ViewReport>   
 * Author：Peter.uhm  (yanb@shinetechchina.com)
 * Modify date：2016.04 
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
using System.Diagnostics;

namespace ExcelAddIn4
{
    public partial class ViewReport : Form
    {
        /// <summary>
        /// 
        /// </summary>
        internal static Finance_Tools ft
        {
            get { return new Finance_Tools(); }
        }
        internal static DataTable dtTransaction;
        /// <summary>
        /// 
        /// </summary>
        public ViewReport()
        {
            try
            {
                InitializeComponent();
                InitializeTemplates();
            }
            catch
            { }
        }
        /// <summary>
        /// 
        /// </summary>
        private void InitializeTemplates()
        {
            dtTransaction = ft.GetUserReports();
            var query = from t in dtTransaction.AsEnumerable()
                        group t by new { t1 = t.Field<string>("TemplateName"), t2 = t.Field<string>("TemplateID") } into m
                        select new
                        {
                            TemplateName = m.First().Field<string>("TemplateName"),
                            TemplatePath = m.First().Field<string>("TemplateID"),
                        };
            DataTable newdt = new DataTable();
            newdt.Columns.Add("value");
            newdt.Columns.Add("name");
            foreach (var employee in query)
            {
                DataRow dr = newdt.NewRow();
                dr["value"] = employee.TemplatePath;
                dr["name"] = employee.TemplateName;
                if (!string.IsNullOrEmpty(ft.getFilePath(employee.TemplatePath)))
                    newdt.Rows.Add(dr);
            }
            DataRow dr2 = newdt.NewRow();
            dr2["value"] = "";
            dr2["name"] = "";
            newdt.Rows.InsertAt(dr2, 0);
            this.comboBox1.DataSource = newdt;
            this.comboBox1.AutoCompleteMode = AutoCompleteMode.Suggest;
            this.comboBox1.AutoCompleteSource = AutoCompleteSource.CustomSource;
            this.comboBox1.DisplayMember = "name";
            this.comboBox1.ValueMember = "value";
        }
        private void InitializePrefixes()
        {
            var query = from t in dtTransaction.AsEnumerable()
                        group t by new { t1 = t.Field<string>("Prefix") } into m
                        select new
                        {
                            Prefix = m.First().Field<string>("Prefix"),
                        };
            DataTable newdt = new DataTable();
            newdt.Columns.Add("value");
            newdt.Columns.Add("name");
            foreach (var employee in query)
            {
                DataRow dr = newdt.NewRow();
                dr["value"] = employee.Prefix;
                dr["name"] = employee.Prefix;
                //if (!string.IsNullOrEmpty(ft.getFilePath(employee.TemplatePath)))
                newdt.Rows.Add(dr);
            }
            DataRow dr2 = newdt.NewRow();
            dr2["value"] = "";
            dr2["name"] = "";
            newdt.Rows.InsertAt(dr2, 0);
            this.comboBox1.DataSource = newdt;
            this.comboBox1.AutoCompleteMode = AutoCompleteMode.Suggest;
            this.comboBox1.AutoCompleteSource = AutoCompleteSource.CustomSource;
            this.comboBox1.DisplayMember = "name";
            this.comboBox1.ValueMember = "value";
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private bool CheckNull(List<string> str)
        {
            StringBuilder sb = new StringBuilder(); for (int i = 0; i < str.Count; i++) { sb.Append(str[i]); }
            String s = sb.ToString();
            if (string.IsNullOrEmpty(s))
                return false;
            else
                return true;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ctl"></param>
        private void OpenFile(Control[] ctl)
        {
            var xlapp = Globals.ThisAddIn.Application;
            for (int i = 0; i < ctl.Length; i++)
            {
                try
                {
                    if (((RadioButton)ctl[i]).Checked)
                    {
                        var data = ft.ReportData(((RadioButton)ctl[i]).Tag.ToString());
                        var fileType = ft.ReportFileType(((RadioButton)ctl[i]).Tag.ToString());
                        string tmp = Guid.NewGuid().ToString();
                        var file = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType, FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
                        var bw = new BinaryWriter(file);
                        bw.Write(data);
                        bw.Close();
                        file.Close();
                        if (fileType == "pdf")
                        {
                            Process.Start(AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType);
                            return;
                        }
                        else
                        {
                            xlapp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType);
                            xlapp.Run("'" + AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType + "'!RSystems.runProc");
                            return;
                        }
                    }
                    else
                    { }
                }
                catch { }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            var xlapp = Globals.ThisAddIn.Application;
            string cellvalue = this.comboBox1.SelectedValue == null ? "" : this.comboBox1.SelectedValue.ToString();
            string transactionNumber = this.cbTransactionName.SelectedValue == null ? "" : this.cbTransactionName.SelectedValue.ToString();
            if (string.IsNullOrEmpty(cellvalue))
            {
                MessageBox.Show("Please choose a report template!", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            List<string> cris = new List<string>();
            List<string> vals = new List<string>();
            Control[] ctls = this.panel3.Controls.Find("CriteriaLabel", true);
            for (int i = 0; i < 5; i++)
            {
                Control[] ctls2 = this.panel3.Controls.Find("CriteriaCB", true);
                try
                {
                    if (!string.IsNullOrEmpty(((Label)ctls[i]).Text))
                    {
                        cris.Add(((Label)ctls[i]).Text);
                        vals.Add(((ComboBox)ctls2[i]).Text.ToString());
                    }
                    else
                    {
                        cris.Add("");
                        vals.Add("");
                    }
                }
                catch
                {
                    cris.Add("");
                    vals.Add("");
                }
            }
            if (!CheckNull(vals))
            {
                MessageBox.Show("Please input criterias values!", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            try
            {
                if (ft.ReportDataCount(cellvalue, cris, vals) == 0)
                {
                    this.panel4.Controls.Clear();
                    Label lbl = new Label();
                    lbl.Text = "No matching records.";
                    lbl.AutoSize = true;
                    lbl.Location = new System.Drawing.Point(38, 15 * 0);
                    lbl.ForeColor = Color.Red;
                    lbl.Size = new System.Drawing.Size(89, 13);
                    this.panel4.Controls.Add(lbl);
                }
                else if (ft.ReportDataCount(cellvalue, cris, vals) >= 1)
                {
                    var data = ft.ReportData(cellvalue, cris, vals, transactionNumber);
                    var fileType = ft.ReportFileType(cellvalue, cris, vals, transactionNumber);
                    SessionInfo.UserInfo.File_ftid = ft.ReportFilePath(cellvalue, cris, vals, transactionNumber);
                    SessionInfo.UserInfo.InvNumber = ft.ReportInvNumber(cellvalue, cris, vals, transactionNumber);
                    SessionInfo.UserInfo.FilePath = ft.getFilePath(SessionInfo.UserInfo.File_ftid);
                    string tmp = Guid.NewGuid().ToString();
                    var file = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType, FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
                    var bw = new BinaryWriter(file);
                    bw.Write(data);
                    bw.Close();
                    file.Close();
                    if (fileType == "pdf")
                    {
                        Process.Start(AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType);
                        return;
                    }
                    else
                    {
                        xlapp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType);
                        return;
                    }
                }
            }
            catch
            { }
            finally
            {
                this.Close();
                this.Dispose();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!cbSearPrefix.Checked)
            {
                string p = this.comboBox1.SelectedValue.ToString();
                DataTable dt = ft.GetReportCriteria(p);
                var query = from t in dt.AsEnumerable()
                            group t by new { t1 = t.Field<string>("Reference") } into m
                            select new
                            {
                                Reference = m.First().Field<string>("Reference"),
                            };
                DataTable newdt = new DataTable();
                newdt.Columns.Add("value");
                foreach (var employee in query)
                {
                    DataRow dr = newdt.NewRow();
                    dr["value"] = employee.Reference;
                    newdt.Rows.Add(dr);
                }
                DataRow dr2 = newdt.NewRow();
                dr2["value"] = "";
                newdt.Rows.InsertAt(dr2, 0);
                this.cbRef.DataSource = newdt;
                this.cbRef.AutoCompleteMode = AutoCompleteMode.Suggest;
                this.cbRef.AutoCompleteSource = AutoCompleteSource.CustomSource;
                this.cbRef.DisplayMember = "Value";
                this.cbRef.ValueMember = "value";
            }
            else
            {
                DataRow[] newdt2 = dtTransaction.Select(" Prefix='" + this.comboBox1.SelectedValue.ToString() + "'").OrderBy(x=>x["maxNum"]).ToArray();

                //DataTable newdt2 = ft.ToDataTable(dt2);
                var query2 = from t in newdt2.AsEnumerable()
                             group t by new { t1 = t.Field<string>("TransactionName") } into m
                             select new
                             {
                                 CriteriaValue = m.First().Field<string>("TransactionName"),
                             };
                DataTable newdt3 = ft.ToDataTable(query2.ToList());
                this.cbTransactionName.DataSource = newdt3;
                this.cbTransactionName.DisplayMember = "CriteriaValue";
                this.cbTransactionName.ValueMember = "CriteriaValue";
            }
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
                if (!string.IsNullOrEmpty(this.comboBox1.Text))
                {
                    string p = this.comboBox1.SelectedValue.ToString();
                    string r = this.cbRef.SelectedValue.ToString();
                    DataTable newdt = ft.GetReportCriteriaByRef(p, r);
                    int i = 0;
                    if (!string.IsNullOrEmpty(newdt.Rows[0]["CriteriaName1"].ToString()))
                    {
                        l = new Label();
                        l.AutoSize = true;
                        l.Location = new System.Drawing.Point(38, 35 * (i));
                        l.Name = "CriteriaLabel";
                        l.Size = new System.Drawing.Size(89, 13);
                        l.Text = newdt.Rows[0]["CriteriaName1"].ToString();
                        this.panel3.Controls.Add(l);
                        i++;
                    }
                    if (!string.IsNullOrEmpty(newdt.Rows[0]["CriteriaName2"].ToString()))
                    {
                        l = new Label();
                        l.AutoSize = true;
                        l.Location = new System.Drawing.Point(38, 35 * (i));
                        l.Name = "CriteriaLabel";
                        l.Size = new System.Drawing.Size(89, 13);
                        l.Text = newdt.Rows[0]["CriteriaName2"].ToString();
                        this.panel3.Controls.Add(l);
                        i++;
                    }
                    if (!string.IsNullOrEmpty(newdt.Rows[0]["CriteriaName3"].ToString()))
                    {
                        l = new Label();
                        l.AutoSize = true;
                        l.Location = new System.Drawing.Point(38, 35 * (i));
                        l.Name = "CriteriaLabel";
                        l.Size = new System.Drawing.Size(89, 13);
                        l.Text = newdt.Rows[0]["CriteriaName3"].ToString();
                        this.panel3.Controls.Add(l);
                        i++;
                    }
                    if (!string.IsNullOrEmpty(newdt.Rows[0]["CriteriaName4"].ToString()))
                    {
                        l = new Label();
                        l.AutoSize = true;
                        l.Location = new System.Drawing.Point(38, 35 * (i));
                        l.Name = "CriteriaLabel";
                        l.Size = new System.Drawing.Size(89, 13);
                        l.Text = newdt.Rows[0]["CriteriaName4"].ToString();
                        this.panel3.Controls.Add(l);
                        i++;
                    }
                    if (!string.IsNullOrEmpty(newdt.Rows[0]["CriteriaName5"].ToString()))
                    {
                        l = new Label();
                        l.AutoSize = true;
                        l.Location = new System.Drawing.Point(38, 35 * (i));
                        l.Name = "CriteriaLabel";
                        l.Size = new System.Drawing.Size(89, 13);
                        l.Text = newdt.Rows[0]["CriteriaName5"].ToString();
                        this.panel3.Controls.Add(l);
                        i++;
                    }
                }
            }
            catch
            { }
        }
        /// <summary>
        /// 
        /// </summary>
        private void CheckFinalDropDown()
        {
            if (!string.IsNullOrEmpty(this.comboBox1.Text))
            {
                RSFinanceToolsEntities db = new RSFinanceToolsEntities();
                string p = this.comboBox1.SelectedValue.ToString();
                Control[] ctlsr = this.panel3.Controls.Find("CriteriaLabel", true);
                Control[] ctlcb = this.panel3.Controls.Find("CriteriaCB", true);
                if (ctlsr.Length > 0)
                {
                    var dt2 = new List<rsTemplateTransaction>();
                    string text5 = string.Empty;
                    string text4 = string.Empty;
                    string text3 = string.Empty;
                    string text2 = string.Empty;
                    string text1 = string.Empty;
                    string cbValue1 = string.Empty;
                    string cbValue2 = string.Empty;
                    string cbValue3 = string.Empty;
                    string cbValue4 = string.Empty;
                    string cbValue5 = string.Empty;
                    if (ctlsr[0] != null) text1 = ((Label)ctlsr[0]).Text;
                    if (ctlsr.Length > 1 && ctlcb.Length > 1)
                    {
                        if (ctlsr[1] != null) text2 = ((Label)ctlsr[1]).Text;
                    }
                    if (ctlsr.Length > 2 && ctlcb.Length > 2)
                    {
                        if (ctlsr[2] != null) text3 = ((Label)ctlsr[2]).Text;
                    }
                    if (ctlsr.Length > 3 && ctlcb.Length > 3)
                    {
                        if (ctlsr[3] != null) text4 = ((Label)ctlsr[3]).Text;
                    }
                    if (ctlsr.Length > 4 && ctlcb.Length > 4)
                    {
                        if (ctlsr[4] != null) text5 = ((Label)ctlsr[4]).Text;
                    }
                    if (ctlcb[0] != null) cbValue1 = ((ComboBox)ctlcb[0]).Text;
                    if (ctlcb.Length > 1)
                    {
                        if (ctlcb[1] != null) cbValue2 = ((ComboBox)ctlcb[1]).Text;
                    }
                    if (ctlcb.Length > 2)
                    {
                        if (ctlcb[2] != null) cbValue3 = ((ComboBox)ctlcb[2]).Text;
                    }
                    if (ctlcb.Length > 3)
                    {
                        if (ctlcb[3] != null) cbValue4 = ((ComboBox)ctlcb[3]).Text;
                    }
                    if (ctlcb.Length > 4)
                    {
                        if (ctlcb[4] != null) cbValue5 = ((ComboBox)ctlcb[4]).Text;
                    }
                    dt2 = (from FT_sett in db.rsTemplateTransactions
                           where FT_sett.TemplateID == p
                            && FT_sett.Criteria5 == text5
                              && FT_sett.Criteria4 == text4
                              && FT_sett.Criteria3 == text3
                           && FT_sett.Criteria2 == text2
                           && FT_sett.Criteria1 == text1
                           && FT_sett.Value1 == cbValue1
                           && FT_sett.Value2 == cbValue2
                           && FT_sett.Value3 == cbValue3
                           && FT_sett.Value4 == cbValue4
                           && FT_sett.Value5 == cbValue5

                           select FT_sett).ToList();

                    DataTable newdt2 = ft.ToDataTable(dt2);
                    var query2 = from t in newdt2.AsEnumerable()
                                 group t by new { t1 = t.Field<string>("TransactionName") } into m
                                 select new
                                 {
                                     CriteriaValue = m.First().Field<string>("TransactionName"),
                                 };
                    DataTable newdt3 = ft.ToDataTable(query2.ToList());
                    this.cbTransactionName.DataSource = newdt3;
                    this.cbTransactionName.DisplayMember = "CriteriaValue";
                    this.cbTransactionName.ValueMember = "CriteriaValue";
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cb0_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                ComboBox cb;
                if (!string.IsNullOrEmpty(this.comboBox1.Text) && !string.IsNullOrEmpty(this.cbRef.Text))
                {
                    RSFinanceToolsEntities db = new RSFinanceToolsEntities();
                    string p = this.comboBox1.SelectedValue.ToString();
                    string Value1 = this.comboBox1.SelectedValue.ToString();
                    Control[] ctlsr = this.panel3.Controls.Find("CriteriaLabel", true);
                    Control[] ctlcb = this.panel3.Controls.Find("CriteriaCB", true);
                    if (ctlsr.Length > 0)
                    {
                        var dt2 = new List<rsTemplateTransaction>();
                        cb = new ComboBox();
                        cb.FormattingEnabled = true;
                        cb.Location = new System.Drawing.Point(198, 35 * (0));
                        cb.Size = new System.Drawing.Size(229, 22);
                        cb.Name = "CriteriaCB";
                        string text1 = ((Label)ctlsr[0]).Text;
                        dt2 = (from FT_sett in db.rsTemplateTransactions
                               where FT_sett.TemplateID == p
                               && FT_sett.Criteria1 == text1
                               select FT_sett).ToList();

                        DataTable newdt2 = ft.ToDataTable(dt2);
                        var query2 = from t in newdt2.AsEnumerable()
                                     group t by new { t1 = t.Field<string>("Value1") } into m
                                     select new
                                     {
                                         CriteriaValue = m.First().Field<string>("Value1"),
                                     };
                        DataTable newdt3 = ft.ToDataTable(query2.ToList());
                        DataRow dr2 = newdt3.NewRow();
                        dr2["CriteriaValue"] = "";
                        newdt3.Rows.InsertAt(dr2, 0);
                        cb.DisplayMember = "CriteriaValue";
                        cb.DataSource = newdt3;
                        cb.ValueMember = "CriteriaValue";
                        cb.SelectedIndexChanged += new EventHandler(cb1_SelectedIndexChanged);
                        if (ctlcb.Length > 0)
                            for (int i = 0; i < ctlcb.Length; i++)
                                this.panel3.Controls.Remove(ctlcb[i]);

                        this.panel3.Controls.Add(cb);
                        cb1_SelectedIndexChanged(null, null);
                    }
                }
                CheckFinalDropDown();
            }
            catch (Exception ex)
            {
                string s = ex.Message;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cb1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                ComboBox cb;
                if (!string.IsNullOrEmpty(this.comboBox1.Text))
                {
                    RSFinanceToolsEntities db = new RSFinanceToolsEntities();
                    string p = this.comboBox1.SelectedValue.ToString();
                    string Value1 = this.comboBox1.SelectedValue.ToString();
                    Control[] ctlsr = this.panel3.Controls.Find("CriteriaLabel", true);
                    Control[] ctlcb = this.panel3.Controls.Find("CriteriaCB", true);
                    if (ctlsr.Length > 1)
                    {
                        var dt2 = new List<rsTemplateTransaction>();
                        cb = new ComboBox();
                        cb.FormattingEnabled = true;
                        cb.Location = new System.Drawing.Point(198, 35 * (1));
                        cb.Size = new System.Drawing.Size(229, 22);
                        cb.Name = "CriteriaCB";
                        string text2 = ((Label)ctlsr[1]).Text;
                        string text1 = ((Label)ctlsr[0]).Text;
                        string cbValue1 = ((ComboBox)ctlcb[0]).Text;
                        dt2 = (from FT_sett in db.rsTemplateTransactions
                               where FT_sett.TemplateID == p
                               && FT_sett.Criteria2 == text2
                               && FT_sett.Criteria1 == text1
                               && FT_sett.Value1 == cbValue1
                               select FT_sett).ToList();

                        DataTable newdt2 = ft.ToDataTable(dt2);
                        var query2 = from t in newdt2.AsEnumerable()
                                     group t by new { t1 = t.Field<string>("Value2") } into m
                                     select new
                                     {
                                         CriteriaValue = m.First().Field<string>("Value2"),
                                     };
                        DataTable newdt3 = ft.ToDataTable(query2.ToList());

                        DataRow dr2 = newdt3.NewRow();
                        dr2["CriteriaValue"] = "";
                        newdt3.Rows.InsertAt(dr2, 0);

                        cb.DisplayMember = "CriteriaValue";
                        cb.DataSource = newdt3;
                        cb.ValueMember = "CriteriaValue";
                        cb.SelectedIndexChanged += new EventHandler(cb2_SelectedIndexChanged);
                        if (ctlcb.Length > 1)
                            for (int i = 1; i < ctlcb.Length; i++)
                                this.panel3.Controls.Remove(ctlcb[i]);

                        this.panel3.Controls.Add(cb);
                        cb2_SelectedIndexChanged(null, null);
                    }
                }
                CheckFinalDropDown();
            }
            catch (Exception ex)
            {
                string s = ex.Message;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cb2_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                ComboBox cb;
                if (!string.IsNullOrEmpty(this.comboBox1.Text))
                {
                    RSFinanceToolsEntities db = new RSFinanceToolsEntities();
                    string p = this.comboBox1.SelectedValue.ToString();
                    Control[] ctlsr = this.panel3.Controls.Find("CriteriaLabel", true);
                    Control[] ctlcb = this.panel3.Controls.Find("CriteriaCB", true);
                    if (ctlsr.Length > 2)
                    {
                        var dt2 = new List<rsTemplateTransaction>();
                        cb = new ComboBox();
                        cb.FormattingEnabled = true;
                        cb.Location = new System.Drawing.Point(198, 35 * (2));
                        cb.Size = new System.Drawing.Size(229, 22);
                        cb.Name = "CriteriaCB";
                        string text3 = ((Label)ctlsr[2]).Text;
                        string text2 = ((Label)ctlsr[1]).Text;
                        string text1 = ((Label)ctlsr[0]).Text;
                        string cbValue1 = ((ComboBox)ctlcb[0]).Text;
                        string cbValue2 = ((ComboBox)ctlcb[1]).Text;
                        dt2 = (from FT_sett in db.rsTemplateTransactions
                               where FT_sett.TemplateID == p
                               && FT_sett.Criteria3 == text3
                               && FT_sett.Criteria2 == text2
                               && FT_sett.Criteria1 == text1
                               && FT_sett.Value1 == cbValue1
                               && FT_sett.Value2 == cbValue2
                               select FT_sett).ToList();

                        DataTable newdt2 = ft.ToDataTable(dt2);
                        var query2 = from t in newdt2.AsEnumerable()
                                     group t by new { t1 = t.Field<string>("Value3") } into m
                                     select new
                                     {
                                         CriteriaValue = m.First().Field<string>("Value3"),
                                     };
                        DataTable newdt3 = ft.ToDataTable(query2.ToList());
                        DataRow dr2 = newdt3.NewRow();
                        dr2["CriteriaValue"] = "";
                        newdt3.Rows.InsertAt(dr2, 0);
                        cb.DisplayMember = "CriteriaValue";
                        cb.DataSource = newdt3;
                        cb.ValueMember = "CriteriaValue";
                        cb.SelectedIndexChanged += new EventHandler(cb3_SelectedIndexChanged);
                        if (ctlcb.Length > 2)
                            for (int i = 2; i < ctlcb.Length; i++)
                                this.panel3.Controls.Remove(ctlcb[i]);

                        this.panel3.Controls.Add(cb);
                        cb3_SelectedIndexChanged(null, null);
                    }
                }
                CheckFinalDropDown();
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cb3_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                ComboBox cb;
                if (!string.IsNullOrEmpty(this.comboBox1.Text))
                {
                    RSFinanceToolsEntities db = new RSFinanceToolsEntities();
                    string p = this.comboBox1.SelectedValue.ToString();
                    Control[] ctlsr = this.panel3.Controls.Find("CriteriaLabel", true);
                    Control[] ctlcb = this.panel3.Controls.Find("CriteriaCB", true);
                    if (ctlsr.Length > 3)
                    {
                        var dt2 = new List<rsTemplateTransaction>();
                        cb = new ComboBox();
                        cb.FormattingEnabled = true;
                        cb.Location = new System.Drawing.Point(198, 35 * (3));
                        cb.Size = new System.Drawing.Size(229, 22);
                        cb.Name = "CriteriaCB";
                        string text4 = ((Label)ctlsr[3]).Text;
                        string text3 = ((Label)ctlsr[2]).Text;
                        string text2 = ((Label)ctlsr[1]).Text;
                        string text1 = ((Label)ctlsr[0]).Text;
                        string cbValue1 = ((ComboBox)ctlcb[0]).Text;
                        string cbValue2 = ((ComboBox)ctlcb[1]).Text;
                        string cbValue3 = ((ComboBox)ctlcb[2]).Text;
                        dt2 = (from FT_sett in db.rsTemplateTransactions
                               where FT_sett.TemplateID == p
                                 && FT_sett.Criteria4 == text4
                                  && FT_sett.Criteria3 == text3
                               && FT_sett.Criteria2 == text2
                               && FT_sett.Criteria1 == text1
                               && FT_sett.Value1 == cbValue1
                               && FT_sett.Value2 == cbValue2
                               && FT_sett.Value3 == cbValue3
                               select FT_sett).ToList();
                        DataTable newdt2 = ft.ToDataTable(dt2);
                        var query2 = from t in newdt2.AsEnumerable()
                                     group t by new { t1 = t.Field<string>("Value4") } into m
                                     select new
                                     {
                                         CriteriaValue = m.First().Field<string>("Value4"),
                                     };
                        DataTable newdt3 = ft.ToDataTable(query2.ToList());
                        DataRow dr2 = newdt3.NewRow();
                        dr2["CriteriaValue"] = "";
                        newdt3.Rows.InsertAt(dr2, 0);
                        cb.DisplayMember = "CriteriaValue";
                        cb.DataSource = newdt3;
                        cb.ValueMember = "CriteriaValue";
                        cb.SelectedIndexChanged += new EventHandler(cb4_SelectedIndexChanged);
                        if (ctlcb.Length > 3)
                            for (int i = 3; i < ctlcb.Length; i++)
                                this.panel3.Controls.Remove(ctlcb[i]);

                        this.panel3.Controls.Add(cb);
                        cb4_SelectedIndexChanged(null, null);
                    }
                }
                CheckFinalDropDown();
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cb4_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                ComboBox cb;
                if (!string.IsNullOrEmpty(this.comboBox1.Text))
                {
                    RSFinanceToolsEntities db = new RSFinanceToolsEntities();
                    string p = this.comboBox1.SelectedValue.ToString();
                    Control[] ctlsr = this.panel3.Controls.Find("CriteriaLabel", true);
                    Control[] ctlcb = this.panel3.Controls.Find("CriteriaCB", true);
                    if (ctlsr.Length > 4)
                    {
                        var dt2 = new List<rsTemplateTransaction>();
                        cb = new ComboBox();
                        cb.FormattingEnabled = true;
                        cb.Location = new System.Drawing.Point(198, 35 * (4));
                        cb.Size = new System.Drawing.Size(229, 22);
                        cb.Name = "CriteriaCB";
                        string text5 = ((Label)ctlsr[4]).Text;
                        string text4 = ((Label)ctlsr[3]).Text;
                        string text3 = ((Label)ctlsr[2]).Text;
                        string text2 = ((Label)ctlsr[1]).Text;
                        string text1 = ((Label)ctlsr[0]).Text;
                        string cbValue1 = ((ComboBox)ctlcb[0]).Text;
                        string cbValue2 = ((ComboBox)ctlcb[1]).Text;
                        string cbValue3 = ((ComboBox)ctlcb[2]).Text;
                        string cbValue4 = ((ComboBox)ctlcb[3]).Text;
                        dt2 = (from FT_sett in db.rsTemplateTransactions
                               where FT_sett.TemplateID == p
                                && FT_sett.Criteria5 == text5
                                  && FT_sett.Criteria4 == text4
                                  && FT_sett.Criteria3 == text3
                               && FT_sett.Criteria2 == text2
                               && FT_sett.Criteria1 == text1
                               && FT_sett.Value1 == cbValue1
                               && FT_sett.Value2 == cbValue2
                               && FT_sett.Value3 == cbValue3
                               && FT_sett.Value4 == cbValue4
                               select FT_sett).ToList();
                        DataTable newdt2 = ft.ToDataTable(dt2);
                        var query2 = from t in newdt2.AsEnumerable()
                                     group t by new { t1 = t.Field<string>("Value5") } into m
                                     select new
                                     {
                                         CriteriaValue = m.First().Field<string>("Value5"),
                                     };
                        DataTable newdt3 = ft.ToDataTable(query2.ToList());
                        DataRow dr2 = newdt3.NewRow();
                        dr2["CriteriaValue"] = "";
                        newdt3.Rows.InsertAt(dr2, 0);
                        cb.DisplayMember = "CriteriaValue";
                        cb.DataSource = newdt3;
                        cb.ValueMember = "CriteriaValue";
                        cb.SelectedIndexChanged += new EventHandler(cb5_SelectedIndexChanged);
                        if (ctlcb.Length > 4)
                            for (int i = 4; i < ctlcb.Length; i++)
                                this.panel3.Controls.Remove(ctlcb[i]);

                        this.panel3.Controls.Add(cb);
                        cb5_SelectedIndexChanged(null, null);
                    }
                }
                CheckFinalDropDown();
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cb5_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                ComboBox cb;
                if (!string.IsNullOrEmpty(this.comboBox1.Text))
                {
                    RSFinanceToolsEntities db = new RSFinanceToolsEntities();
                    string p = this.comboBox1.SelectedValue.ToString();
                    Control[] ctlsr = this.panel3.Controls.Find("CriteriaLabel", true);
                    Control[] ctlcb = this.panel3.Controls.Find("CriteriaCB", true);
                    if (ctlsr.Length > 5)
                    {
                        var dt2 = new List<rsTemplateTransaction>();
                        cb = new ComboBox();
                        cb.FormattingEnabled = true;
                        cb.Location = new System.Drawing.Point(198, 35 * (5));
                        cb.Size = new System.Drawing.Size(229, 22);
                        cb.Name = "CriteriaCB";
                        string text5 = ((Label)ctlsr[4]).Text;
                        string text4 = ((Label)ctlsr[3]).Text;
                        string text3 = ((Label)ctlsr[2]).Text;
                        string text2 = ((Label)ctlsr[1]).Text;
                        string text1 = ((Label)ctlsr[0]).Text;
                        string cbValue1 = ((ComboBox)ctlcb[0]).Text;
                        string cbValue2 = ((ComboBox)ctlcb[1]).Text;
                        string cbValue3 = ((ComboBox)ctlcb[2]).Text;
                        string cbValue4 = ((ComboBox)ctlcb[3]).Text;
                        string cbValue5 = ((ComboBox)ctlcb[4]).Text;
                        dt2 = (from FT_sett in db.rsTemplateTransactions
                               where FT_sett.TemplateID == p
                                && FT_sett.Criteria5 == text5
                                  && FT_sett.Criteria4 == text4
                                  && FT_sett.Criteria3 == text3
                               && FT_sett.Criteria2 == text2
                               && FT_sett.Criteria1 == text1
                               && FT_sett.Value1 == cbValue1
                               && FT_sett.Value2 == cbValue2
                               && FT_sett.Value3 == cbValue3
                               && FT_sett.Value4 == cbValue4
                               && FT_sett.Value5 == cbValue5
                               select FT_sett).ToList();
                        DataTable newdt2 = ft.ToDataTable(dt2);
                        var query2 = from t in newdt2.AsEnumerable()
                                     group t by new { t1 = t.Field<string>("Value5") } into m
                                     select new
                                     {
                                         CriteriaValue = m.First().Field<string>("Value5"),
                                     };
                        DataTable newdt3 = ft.ToDataTable(query2.ToList());
                        DataRow dr2 = newdt3.NewRow();
                        dr2["CriteriaValue"] = "";
                        newdt3.Rows.InsertAt(dr2, 0);

                        cb.DisplayMember = "CriteriaValue";
                        cb.DataSource = newdt3;
                        cb.ValueMember = "CriteriaValue";//cb.SelectedIndexChanged += new EventHandler(cb5_SelectedIndexChanged);
                        if (ctlcb.Length > 5)
                            for (int i = 5; i < ctlcb.Length; i++)
                                this.panel3.Controls.Remove(ctlcb[i]);

                        this.panel3.Controls.Add(cb);
                    }
                }
                CheckFinalDropDown();
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbRef_SelectedIndexChanged(object sender, EventArgs e)
        {
            InitialCriterias();
            cb0_SelectedIndexChanged(null, null);
        }

        private void cbSearPrefix_CheckedChanged(object sender, EventArgs e)
        {
            if (cbSearPrefix.Checked)
            {
                label1.Text = "Prefix";
                label3.Visible = false;
                cbRef.Visible = false;
                this.panel3.Controls.Clear();
                InitializePrefixes();
            }
            else
            {
                label1.Text = "Template";
                label3.Visible = true;
                cbRef.Visible = true;
                InitializeTemplates();
            }
        }
    }
}
