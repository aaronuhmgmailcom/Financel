/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<TransUpdPostFrm>   
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
using System.IO;
using System.Reflection;
using System.Globalization;
using System.Threading;
using System.Net;
using System.Web.Services.Description;
using System.Xml.Serialization;
using System.CodeDom;
using System.CodeDom.Compiler;
using System.Xml;
using System.Text.RegularExpressions;
using System.Diagnostics;
using ExcelAddIn4.Common;

namespace ExcelAddIn4
{
    public partial class TransUpdPostFrm : Form
    {
        /// <summary>
        /// 
        /// </summary>
        DataGridView dgv = null;
        /// <summary>
        /// 
        /// </summary>
        string journalNumber = string.Empty;
        /// <summary>
        /// 
        /// </summary>
        string sequenceNumbering = string.Empty;
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
        public TransUpdPostFrm()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 
        /// </summary>
        public void bddata()
        {
            try
            {
                this.Text = "Journal Update (" + SessionInfo.UserInfo.CurrentRef + ") - RSystems FinanceTools v2";
                dgv = ft.IniXMLFormGrdForTransUpd();
                dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
                dgv.AutoGenerateColumns = false;
                dgv.ColumnHeadersHeight = 40;
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dgv.Dock = DockStyle.Fill;
                dgv.Visible = true;
                dgv.ColumnHeaderMouseClick += new DataGridViewCellMouseEventHandler(DataGridView1_CellMouseDown);
                dgv.RowPostPaint += new DataGridViewRowPostPaintEventHandler(dgv_RowPostPaint);
                dgv.DataBindingComplete += new DataGridViewBindingCompleteEventHandler(ft.XMLFormdataGridView_DataBindingComplete);
                BindData();
                DataTable tb1 = ft.ToDataTable((List<Common2.Specialist>)this.dgv.DataSource);
                for (int i = 0; i < tb1.Columns.Count; i++)
                {
                    string str = string.Empty;
                    for (int j = 0; j < tb1.Rows.Count; j++)
                    {
                        str += tb1.Rows[j][i].ToString();
                        if (tb1.Rows[j][i].ToString().ToUpper() == "[SEQUENCE]")
                            sequenceNumbering = tb1.Columns[i].ColumnName + "," + sequenceNumbering;
                    }
                    if ((string.IsNullOrEmpty(str) && tb1.Columns[i].ColumnName != "DebitCredit" && tb1.Columns[i].ColumnName != "DetailLad") || ((string.IsNullOrEmpty(str.Replace("0", "").Replace(".", "")) ? true : false) && tb1.Columns[i].ColumnName != "DebitCredit" && tb1.Columns[i].ColumnName != "DetailLad"))
                        dgv.Columns.Remove(tb1.Columns[i].ColumnName.Replace("GenDesc", "GeneralDescription"));

                    if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(tb1.Columns[i].ColumnName))
                        dgv.Columns.Remove(tb1.Columns[i].ColumnName.Replace("GenDesc", "GeneralDescription"));
                }
                for (int i = 0; i < dgv.Columns.Count; i++)
                {
                    if (OutputContainer.updateStatus.Contains(dgv.Columns[i].DataPropertyName) || OutputContainer.updateStatusForPost.Contains(dgv.Columns[i].DataPropertyName))
                        dgv.Columns[i].DefaultCellStyle.BackColor = Color.Aqua;
                }
                this.tabPage1.Controls.Add(dgv);
            }
            catch (Exception ex)
            { MessageBox.Show(ex.Message); }
            tabControl1_SelectedIndexChanged(null, null);
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
        private int currentColumnIndex = 0;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DataGridView1_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                if (e.Button == MouseButtons.Right)
                {
                    if (e.ColumnIndex >= 0)
                    {
                        if (dgv.Columns[e.ColumnIndex].Selected == false)
                        {
                            dgv.ClearSelection();
                            dgv.Columns[e.ColumnIndex].Selected = true;
                            currentColumnIndex = e.ColumnIndex;
                        }
                        this.contextMenuStrip1.Show(MousePosition.X, MousePosition.Y);
                    }
                }
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        private void BindData()
        {
            try
            {
                this.dgv.DataSource = Ribbon2.outputPane.TransUpdFinallist;
            }
            catch (Exception ex)
            {
                TransUpdPostFrm.richTextBox1.Text += ex.Message;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void btnPost_Click(object sender, EventArgs e)
        {
            if (DoPost(sender))
            {
            }
            else
            {
                throw new Exception();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <returns></returns>
        private bool DoPost(object sender)
        {
            string sSSCVoucher = "";
            SunSystems.Connect.Client.SecurityManager oSecMan = null;
            try
            {
                oSecMan = new SunSystems.Connect.Client.SecurityManager(SessionInfo.UserInfo.SunUserIP);
                ////http://95.138.187.185:81/SecurityWebServer/Login.aspx?redirect=http://95.138.187.185:8080/ssc/login.jsp
                oSecMan.Login(SessionInfo.UserInfo.SunUserID, SessionInfo.UserInfo.SunUserPass);
                if (oSecMan.Authorised)
                {
                    sSSCVoucher = oSecMan.Voucher;
                }
                else
                {
                    this.textBox1.Text = "SunSystems Server is not exist or Password for user is incorrect.";
                    if (Ribbon2.tupf.Visible == true)
                        MessageBox.Show(this.textBox1.Text, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else
                        SessionInfo.UserInfo.GlobalError += "Process:Journal Update(" + SessionInfo.UserInfo.CurrentRef + ") - Fail! " + this.textBox1.Text;

                    return false;
                }
            }
            catch (Exception ex)
            {
                this.textBox1.Text = "An error occurred in validation " + (oSecMan == null ? "" : oSecMan.ErrorMessage) + ex;
                if (Ribbon2.tupf.Visible == true)
                    MessageBox.Show(this.textBox1.Text, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    SessionInfo.UserInfo.GlobalError += "Process:Journal Update(" + SessionInfo.UserInfo.CurrentRef + ") - Fail! " + this.textBox1.Text;

                LogHelper.WriteLog(typeof(TransUpdPostFrm), this.textBox1.Text);
                return false;
            }
            finally
            {
                oSecMan = null;
            }
            try
            {
                WebClient web = new WebClient();
                Stream stream = web.OpenRead("http://" + SessionInfo.UserInfo.SunUserIP + ":8080/connect/wsdl/ComponentExecutor?wsdl");
                ServiceDescription description = ServiceDescription.Read(stream);
                ServiceDescriptionImporter importer = new ServiceDescriptionImporter();
                importer.ProtocolName = "Soap"; // The specified access protocol.
                importer.Style = ServiceDescriptionImportStyle.Client; // To generate client proxy.
                importer.CodeGenerationOptions = CodeGenerationOptions.GenerateProperties | CodeGenerationOptions.GenerateNewAsync;
                importer.AddServiceDescription(description, null, null); // Add a WSDL document.
                CodeNamespace nmspace = new CodeNamespace(); // Add a namespace for the proxy class, the default is the global space.
                CodeCompileUnit unit = new CodeCompileUnit();
                unit.Namespaces.Add(nmspace);
                ServiceDescriptionImportWarnings warning = importer.Import(nmspace, unit);
                CodeDomProvider provider = CodeDomProvider.CreateProvider("CSharp");
                CompilerParameters parameter = new CompilerParameters();
                parameter.GenerateExecutable = false;
                parameter.GenerateInMemory = true;
                parameter.ReferencedAssemblies.Add("System.dll");
                parameter.ReferencedAssemblies.Add("System.XML.dll");
                parameter.ReferencedAssemblies.Add("System.Web.Services.dll");
                parameter.ReferencedAssemblies.Add("System.Data.dll");
                CompilerResults result = provider.CompileAssemblyFromDom(parameter, unit);
                if (!result.Errors.HasErrors)
                {
                    Assembly asm = result.CompiledAssembly;
                    Type t = asm.GetType("ComponentExecutor", true, true); // If in front of adding a namespace for the proxy class, here need to be added to the front of the namespace type.
                    object o = Activator.CreateInstance(t);
                    MethodInfo method = t.GetMethod("Execute");
                    string sInputPayload;
                    sInputPayload = this.txtXML.Text.Replace("\r\n", "").Replace("\n", ""); ;
                    string[] sArray1 = sInputPayload.Split(new char[3] { '*', '*', '*' });
                    object strResu = null;
                    foreach (string ss in sArray1)
                    {
                        if (!string.IsNullOrEmpty(ss))
                        {
                            strResu = method.Invoke(o, new object[] { sSSCVoucher, null, "LedgerTransaction", "LedgerAnalysisUpdate", null, ss });
                            this.textBox1.Text += GetErrorLines(strResu.ToString()) + "\r\n***\r\n";
                        }
                    }
                    //PostErrorFrm pef = new PostErrorFrm(this.textBox1.Text.Replace("This line has been rejected due to errors in other lines or posting options", ""));
                    //pef.ShowDialog();
                    if (!string.IsNullOrEmpty(this.textBox1.Text.Trim().Replace("***", "")))
                    {
                        SessionInfo.UserInfo.GlobalError += "Process:Journal Update(" + SessionInfo.UserInfo.CurrentRef + ") - Fail! " + this.textBox1.Text;
                        return false;
                    }
                    else
                    {
                        if (Ribbon2.tupf.Visible == true)
                            MessageBox.Show("Update Successful", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        else
                            SessionInfo.UserInfo.GlobalError += "Process:Journal Update(" + SessionInfo.UserInfo.CurrentRef + ") - Success! ";
                    }
                    if (strResu != null)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                if (Ribbon2.tupf.Visible == true)
                    MessageBox.Show(ex.ToString(), "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    SessionInfo.UserInfo.GlobalError += "Process:Journal Update(" + SessionInfo.UserInfo.CurrentRef + ") - Fail! " + ex.Message;

                LogHelper.WriteLog(typeof(Ribbon2), ex.Message + "Do Post error");
                return false;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private string GetErrorLines(string str)
        {
            string error = string.Empty;
            string reference = string.Empty;
            string usertext = string.Empty;
            string lineInfo = string.Empty;
            //XmlOperator Xo = new XmlOperator("<?xml version='1.0' encoding='UTF-8' ?><SSC><ErrorContext><CompatibilityMode>0</CompatibilityMode><ErrorOutput>1</ErrorOutput><ErrorThreshold>1</ErrorThreshold></ErrorContext><User><Name>SSC</Name></User><SunSystemsContext><BusinessUnit>PK1</BusinessUnit><BudgetCode>A</BudgetCode></SunSystemsContext><Payload><LedgerUpdate status='fail'><AccountCode>11000</AccountCode><AllocationMarker>NOT_ALLOCATED</AllocationMarker><AnalysisCode1>               </AnalysisCode1><AnalysisCode10>               </AnalysisCode10><AnalysisCode2>13             </AnalysisCode2><AnalysisCode3>               </AnalysisCode3><AnalysisCode4>               </AnalysisCode4><AnalysisCode5>               </AnalysisCode5><AnalysisCode6>V              </AnalysisCode6><AnalysisCode7>               </AnalysisCode7><AnalysisCode8>               </AnalysisCode8><AnalysisCode9>               </AnalysisCode9><BaseAmount>247.500</BaseAmount><DebitCredit>C</DebitCredit><JournalLineNumber>1</JournalLineNumber><JournalNumber>24</JournalNumber><JournalType>SIMC</JournalType><Actions status='fail'><AnalysisCode1 Reference='1'>TEST</AnalysisCode1></Actions><Messages><Message Level='error' Reference='1'><UserText>Analysis Dimension is set to validated but the Analysis Code Value does not exist.  Analysis data not updated.</UserText><Application><Component>LedgerTransaction</Component><Method>LedgerAnalysisUpdate</Method><Version>42</Version></Application></Message></Messages></LedgerUpdate><LedgerUpdate status='fail'><AccountCode>64005</AccountCode><AllocationMarker>NOT_ALLOCATED</AllocationMarker><AnalysisCode1>               </AnalysisCode1><AnalysisCode10>               </AnalysisCode10><AnalysisCode2>               </AnalysisCode2><AnalysisCode3>               </AnalysisCode3><AnalysisCode4>               </AnalysisCode4><AnalysisCode5>               </AnalysisCode5><AnalysisCode6>               </AnalysisCode6><AnalysisCode7>               </AnalysisCode7><AnalysisCode8>               </AnalysisCode8><AnalysisCode9>               </AnalysisCode9><BaseAmount>-290.810</BaseAmount><DebitCredit>D</DebitCredit><JournalLineNumber>2</JournalLineNumber><JournalNumber>24</JournalNumber><JournalType>SIMC</JournalType><Actions status='fail'><AnalysisCode1 Reference='2'>10001</AnalysisCode1></Actions><Messages><Message Level='error' Reference='2'><UserText>Amendment of this Analysis Dimension is prohibited</UserText><Application><Component>LedgerTransaction</Component><Method>LedgerAnalysisUpdate</Method><Version>42</Version></Application></Message></Messages></LedgerUpdate><LedgerUpdate status='fail'><AccountCode>11000</AccountCode><AllocationMarker>NOT_ALLOCATED</AllocationMarker><AnalysisCode1>               </AnalysisCode1><AnalysisCode10>               </AnalysisCode10><AnalysisCode2>13             </AnalysisCode2><AnalysisCode3>               </AnalysisCode3><AnalysisCode4>               </AnalysisCode4><AnalysisCode5>               </AnalysisCode5><AnalysisCode6>V              </AnalysisCode6><AnalysisCode7>               </AnalysisCode7><AnalysisCode8>               </AnalysisCode8><AnalysisCode9>               </AnalysisCode9><BaseAmount>247.500</BaseAmount><DebitCredit>C</DebitCredit><JournalLineNumber>1</JournalLineNumber><JournalNumber>24</JournalNumber><JournalType>SIMC</JournalType><Actions status='fail'><AnalysisCode1 Reference='3'>TEST</AnalysisCode1></Actions><Messages><Message Level='error' Reference='3'><UserText>Analysis Dimension is set to validated but the Analysis Code Value does not exist.  Analysis data not updated.</UserText><Application><Component>LedgerTransaction</Component><Method>LedgerAnalysisUpdate</Method><Version>42</Version></Application></Message></Messages></LedgerUpdate><LedgerUpdate status='fail'><AccountCode>84050</AccountCode><AllocationMarker>NOT_ALLOCATED</AllocationMarker><AnalysisCode1>               </AnalysisCode1><AnalysisCode10>               </AnalysisCode10><AnalysisCode2>               </AnalysisCode2><AnalysisCode3>               </AnalysisCode3><AnalysisCode4>               </AnalysisCode4><AnalysisCode5>               </AnalysisCode5><AnalysisCode6>               </AnalysisCode6><AnalysisCode7>               </AnalysisCode7><AnalysisCode8>               </AnalysisCode8><AnalysisCode9>               </AnalysisCode9><BaseAmount>485.000</BaseAmount><DebitCredit>C</DebitCredit><JournalLineNumber>2</JournalLineNumber><JournalNumber>1</JournalNumber><JournalType>PACC</JournalType><Actions status='fail'><AnalysisCode1 Reference='4'>TEST</AnalysisCode1></Actions><Messages><Message Level='error' Reference='4'><UserText>Amendment of this Analysis Dimension is prohibited</UserText><Application><Component>LedgerTransaction</Component><Method>LedgerAnalysisUpdate</Method><Version>42</Version></Application></Message></Messages></LedgerUpdate><LedgerUpdate Reference='5' status='fail'><JournalLineNumber>3</JournalLineNumber><JournalNumber>2</JournalNumber><Actions><AnalysisCode1>10001</AnalysisCode1></Actions><Messages><Message><Exception>.</Exception><UserText>.</UserText><Application><Component>.</Component><DataItem>.</DataItem><Driver>.</Driver><Item>.</Item><LastMethod>.</LastMethod><Message>.</Message><MessageNumber>.</MessageNumber><Method>.</Method><Type>.</Type><Value>.</Value><Version>.</Version></Application></Message><Message Level='error' Reference='5'><UserText>There are no transactions that satisfy the payload selection criteria</UserText><Application><Component>LedgerTransaction</Component><Method>LedgerAnalysisUpdate</Method><Version>42</Version></Application></Message></Messages></LedgerUpdate><LedgerUpdate status='fail'><AccountCode>11000</AccountCode><AllocationMarker>NOT_ALLOCATED</AllocationMarker><AnalysisCode1>               </AnalysisCode1><AnalysisCode10>               </AnalysisCode10><AnalysisCode2>13             </AnalysisCode2><AnalysisCode3>               </AnalysisCode3><AnalysisCode4>               </AnalysisCode4><AnalysisCode5>               </AnalysisCode5><AnalysisCode6>V              </AnalysisCode6><AnalysisCode7>               </AnalysisCode7><AnalysisCode8>               </AnalysisCode8><AnalysisCode9>               </AnalysisCode9><BaseAmount>247.500</BaseAmount><DebitCredit>C</DebitCredit><JournalLineNumber>1</JournalLineNumber><JournalNumber>24</JournalNumber><JournalType>SIMC</JournalType><Actions status='fail'><AnalysisCode1 Reference='6'>10002</AnalysisCode1></Actions><Messages><Message Level='error' Reference='6'><UserText>Analysis Dimension is set to validated but the Analysis Code Value does not exist.  Analysis data not updated.</UserText><Application><Component>LedgerTransaction</Component><Method>LedgerAnalysisUpdate</Method><Version>42</Version></Application></Message></Messages></LedgerUpdate></Payload></SSC>");
            XmlOperator Xo = new XmlOperator(str);
            //XmlNodeList userCollection = Xo.XmlDoc.GetElementsByTagName("LedgerUpdate");
            XmlNodeList xnl = Xo.XmlDoc.SelectNodes("//LedgerUpdate/Messages/Message");
            XmlNodeList xnl2 = Xo.XmlDoc.SelectNodes("//LedgerUpdate/Messages/Message/UserText");
            XmlNode xnl3 = Xo.XmlDoc.SelectSingleNode("//BudgetCode");
            error += "Ledger : " + xnl3.InnerText + "\r\n";
            int i = 0;
            foreach (XmlNode xn in xnl)
            {
                if (xn.ParentNode.ParentNode.Attributes["status"].Value == "fail")
                {
                    if (xnl[i].Attributes["Reference"] != null && xnl2[i] != null)
                    {
                        reference = "Line " + xnl[i].Attributes["Reference"].Value;
                        string JournalNumber = string.Empty;
                        string JournalLineNumber = string.Empty;
                        for (int j = 0; j < xn.ParentNode.ParentNode.ChildNodes.Count; j++)
                        {
                            XmlNode tmp = xn.ParentNode.ParentNode.ChildNodes[j];
                            if (tmp.Name == "JournalNumber")
                            {
                                JournalNumber = tmp.InnerText;
                            }
                            if (tmp.Name == "JournalLineNumber")
                            {
                                JournalLineNumber = tmp.InnerText;
                            }
                            lineInfo = "(JournalNumber " + JournalNumber + ", JournalLineNumber " + JournalLineNumber + ")";
                        }
                        usertext = xnl2[i].InnerText;
                    }
                    if (usertext != "This line has been rejected due to errors in other lines or posting options")
                        error += reference + lineInfo + " : " + usertext + "\r\n";
                }
                i++;
            }
            return error;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void contextMenuScript1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (currentColumnIndex >= 0)
            {
                dgv.Columns.RemoveAt(currentColumnIndex);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.tabControl1.SelectedIndex == 1 || this.tabControl1.SelectedIndex == 0)
            {
                try
                {
                    if (dgv != null)
                    {
                        List<ExcelAddIn4.Common2.LedgerUpdate> newlist = new List<ExcelAddIn4.Common2.LedgerUpdate>();
                        for (int i = 0; i < dgv.Rows.Count; i++)
                        {
                            ExcelAddIn4.Common2.LedgerUpdate re = new ExcelAddIn4.Common2.LedgerUpdate();
                            re.AccountRange = new Common2.AccountRange();
                            re.Actions = new ExcelAddIn4.Common2.Actions();
                            if ((dgv.Columns["JournalNumber"] != null) && (dgv.Columns["JournalNumber"].Visible == true))
                                re.JournalNumber = dgv.Rows[i].Cells["JournalNumber"] == null ? "" : dgv.Rows[i].Cells["JournalNumber"].Value.ToString();
                            if ((dgv.Columns["JournalLineNumber"] != null) && (dgv.Columns["JournalLineNumber"].Visible == true))
                                re.JournalLineNumber = dgv.Rows[i].Cells["JournalLineNumber"] == null ? "" : dgv.Rows[i].Cells["JournalLineNumber"].Value.ToString();
                            if ((dgv.Columns["Ledger"] != null) && (dgv.Columns["Ledger"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("Ledger,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Ledger = invName;
                                    dgv.Rows[i].Cells["Ledger"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("Ledger,"))
                                {
                                    re.Ledger = null;
                                    dgv.Rows[i].Cells["Ledger"].Value = null;
                                }
                                else
                                    re.Ledger = dgv.Rows[i].Cells["Ledger"] == null ? "" : dgv.Rows[i].Cells["Ledger"].Value.ToString();
                            }
                            if ((dgv.Columns["AccountCode"] != null) && dgv.Columns["AccountCode"].Visible == true)
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("AccountCode,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.AccountRange.AccountCodeFrom = invName;
                                    re.AccountRange.AccountCodeTo = invName;
                                    re.AccountCode = invName;
                                    dgv.Rows[i].Cells["AccountCode"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AccountCode,"))
                                {
                                    re.AccountRange.AccountCodeFrom = null;
                                    re.AccountRange.AccountCodeTo = null;
                                    re.AccountCode = null;
                                    dgv.Rows[i].Cells["AccountCode"].Value = null;
                                }
                                else
                                {
                                    re.AccountCode = dgv.Rows[i].Cells["AccountCode"] == null ? "" : dgv.Rows[i].Cells["AccountCode"].Value.ToString();
                                    re.AccountRange.AccountCodeFrom = dgv.Rows[i].Cells["AccountCode"] == null ? "" : dgv.Rows[i].Cells["AccountCode"].Value.ToString();
                                    re.AccountRange.AccountCodeTo = re.AccountRange.AccountCodeFrom;
                                }
                            }
                            if ((dgv.Columns["AccountingPeriod"] != null) && (dgv.Columns["AccountingPeriod"].Visible == true))
                            {
                                re.AccountingPeriod = dgv.Rows[i].Cells["AccountingPeriod"] == null ? "" : dgv.Rows[i].Cells["AccountingPeriod"].Value.ToString();
                                //bool result;
                                //string y = string.Empty;
                                //string m = string.Empty;
                                //if (dgv.Rows[i].Cells["AccountingPeriod"].Value != "")
                                //{
                                //    result = Finance_Tools.IsPeriodString(dgv.Rows[i].Cells["AccountingPeriod"].Value.ToString());
                                //    if (result)
                                //    {
                                //        //Format(Right(2012/001,3)&Left(2012/001,4),"0000000")
                                //        y = dgv.Rows[i].Cells["AccountingPeriod"].Value.ToString().Replace("/", "").Substring(0, 4);
                                //        m = dgv.Rows[i].Cells["AccountingPeriod"].Value.ToString().Replace("/", "").Substring(4, 3);
                                //        int a = 0;
                                //        if ((int.TryParse(y, out a) == false) || (int.TryParse(m, out a) == false))
                                //        {
                                //            throw new Exception("AccountingPeriod format is not correct");
                                //        }
                                //        re.AccountingPeriod = m.PadLeft(3, '0') + y;//;
                                //    }
                                //    else
                                //    {
                                //        throw new Exception("AccountingPeriod format is not correct");
                                //    }
                                //}
                            }
                            if ((dgv.Columns["TransactionDate"] != null) && (dgv.Columns["TransactionDate"].Visible == true))
                            {
                                bool result;
                                DateTime r;
                                result = DateTime.TryParse(dgv.Rows[i].Cells["TransactionDate"].Value.ToString(), out r);
                                if (result)
                                {
                                    re.TransactionDate = r.ToString("yyyyMMdd").Substring(6, 2) + r.ToString("yyyyMMdd").Substring(4, 2) + r.ToString("yyyyMMdd").Substring(0, 4);
                                }
                            }
                            if ((dgv.Columns["DueDate"] != null) && (dgv.Columns["DueDate"].Visible == true))
                            {
                                bool result;
                                DateTime r;
                                result = DateTime.TryParse(dgv.Rows[i].Cells["DueDate"].Value.ToString(), out r);
                                if (result)
                                {
                                    re.DueDate = r.ToString("yyyyMMdd").Substring(6, 2) + r.ToString("yyyyMMdd").Substring(4, 2) + r.ToString("yyyyMMdd").Substring(0, 4);
                                }
                                //re.DueDate = dgv.Rows[i].Cells["DueDate"] == null ? "" : dgv.Rows[i].Cells["DueDate"].Value.ToString();//"ddMMyyyy"
                            }
                            if ((dgv.Columns["JournalType"] != null) && (dgv.Columns["JournalType"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("JournalType,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.JournalType = invName;
                                    dgv.Rows[i].Cells["JournalType"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("JournalType,"))
                                {
                                    re.JournalType = null;
                                    dgv.Rows[i].Cells["JournalType"].Value = null;
                                }
                                else
                                    re.JournalType = dgv.Rows[i].Cells["JournalType"] == null ? "" : dgv.Rows[i].Cells["JournalType"].Value.ToString();
                            }
                            if ((dgv.Columns["JournalSource"] != null) && (dgv.Columns["JournalSource"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("JournalSource,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.JournalSource = invName;
                                    dgv.Rows[i].Cells["JournalSource"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("JournalSource,"))
                                {
                                    re.JournalSource = null;
                                    dgv.Rows[i].Cells["JournalSource"].Value = null;
                                }
                                else
                                    re.JournalSource = dgv.Rows[i].Cells["JournalSource"] == null ? "" : dgv.Rows[i].Cells["JournalSource"].Value.ToString();
                            }
                            if ((dgv.Columns["TransactionReference"] != null) && (dgv.Columns["TransactionReference"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("TransactionReference,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.TransactionReference = invName;
                                    dgv.Rows[i].Cells["TransactionReference"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("TransactionReference,"))
                                {
                                    re.TransactionReference = null;
                                    dgv.Rows[i].Cells["TransactionReference"].Value = null;
                                }
                                else
                                    re.TransactionReference = dgv.Rows[i].Cells["TransactionReference"] == null ? "" : dgv.Rows[i].Cells["TransactionReference"].Value.ToString();
                            }
                            if ((dgv.Columns["Description"] != null) && (dgv.Columns["Description"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("Description,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Description = invName;
                                    dgv.Rows[i].Cells["Description"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("Description,"))
                                {
                                    re.Description = null;
                                    dgv.Rows[i].Cells["Description"].Value = null;
                                }
                                else
                                    re.Description = dgv.Rows[i].Cells["Description"] == null ? "" : dgv.Rows[i].Cells["Description"].Value.ToString();
                            }
                            if ((dgv.Columns["AllocationMarker"] != null) && (dgv.Columns["AllocationMarker"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("AllocationMarker,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.AllocationMarker = invName;
                                    dgv.Rows[i].Cells["AllocationMarker"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AllocationMarker,"))
                                {
                                    re.AllocationMarker = null;
                                    dgv.Rows[i].Cells["AllocationMarker"].Value = null;
                                }
                                else
                                    re.AllocationMarker = dgv.Rows[i].Cells["AllocationMarker"] == null ? "" : dgv.Rows[i].Cells["AllocationMarker"].Value.ToString();
                            }
                            if ((dgv.Columns["AnalysisCode1"] != null) && (dgv.Columns["AnalysisCode1"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("AnalysisCode1,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.AnalysisCode1 = invName;
                                    dgv.Rows[i].Cells["AnalysisCode1"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AnalysisCode1,"))
                                {
                                    re.AnalysisCode1 = null;
                                    dgv.Rows[i].Cells["AnalysisCode1"].Value = null;
                                }
                                else
                                    re.AnalysisCode1 = dgv.Rows[i].Cells["AnalysisCode1"] == null ? "" : dgv.Rows[i].Cells["AnalysisCode1"].Value.ToString();

                                if (dgv.Columns["AnalysisCode1"].DefaultCellStyle.BackColor == Color.Aqua)
                                {
                                    re.Actions.AnalysisCode1 = re.AnalysisCode1;
                                    re.AnalysisCode1 = "";
                                }
                            }
                            if ((dgv.Columns["AnalysisCode2"] != null) && (dgv.Columns["AnalysisCode2"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("AnalysisCode2,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.AnalysisCode2 = invName;
                                    dgv.Rows[i].Cells["AnalysisCode2"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AnalysisCode2,"))
                                {
                                    re.AnalysisCode2 = null;
                                    dgv.Rows[i].Cells["AnalysisCode2"].Value = null;
                                }
                                else
                                    re.AnalysisCode2 = dgv.Rows[i].Cells["AnalysisCode2"] == null ? "" : dgv.Rows[i].Cells["AnalysisCode2"].Value.ToString();

                                if (dgv.Columns["AnalysisCode2"].DefaultCellStyle.BackColor == Color.Aqua)
                                {
                                    re.Actions.AnalysisCode2 = re.AnalysisCode2;
                                    re.AnalysisCode2 = "";
                                }
                            }
                            if ((dgv.Columns["AnalysisCode3"] != null) && (dgv.Columns["AnalysisCode3"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("AnalysisCode3,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.AnalysisCode3 = invName;
                                    dgv.Rows[i].Cells["AnalysisCode3"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AnalysisCode3,"))
                                {
                                    re.AnalysisCode3 = null;
                                    dgv.Rows[i].Cells["AnalysisCode3"].Value = null;
                                }
                                else
                                    re.AnalysisCode3 = dgv.Rows[i].Cells["AnalysisCode3"] == null ? "" : dgv.Rows[i].Cells["AnalysisCode3"].Value.ToString();

                                if (dgv.Columns["AnalysisCode3"].DefaultCellStyle.BackColor == Color.Aqua)
                                {
                                    re.Actions.AnalysisCode3 = re.AnalysisCode3;
                                    re.AnalysisCode3 = "";
                                }
                            }
                            if ((dgv.Columns["AnalysisCode4"] != null) && (dgv.Columns["AnalysisCode4"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("AnalysisCode4,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.AnalysisCode4 = invName;
                                    dgv.Rows[i].Cells["AnalysisCode4"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AnalysisCode4,"))
                                {
                                    re.AnalysisCode4 = null;
                                    dgv.Rows[i].Cells["AnalysisCode4"].Value = null;
                                }
                                else
                                    re.AnalysisCode4 = dgv.Rows[i].Cells["AnalysisCode4"] == null ? "" : dgv.Rows[i].Cells["AnalysisCode4"].Value.ToString();

                                if (dgv.Columns["AnalysisCode4"].DefaultCellStyle.BackColor == Color.Aqua)
                                {
                                    re.Actions.AnalysisCode4 = re.AnalysisCode4;
                                    re.AnalysisCode4 = "";
                                }
                            }
                            if ((dgv.Columns["AnalysisCode5"] != null) && (dgv.Columns["AnalysisCode5"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("AnalysisCode5,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.AnalysisCode5 = invName;
                                    dgv.Rows[i].Cells["AnalysisCode5"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AnalysisCode5,"))
                                {
                                    re.AnalysisCode5 = null;
                                    dgv.Rows[i].Cells["AnalysisCode5"].Value = null;
                                }
                                else
                                    re.AnalysisCode5 = dgv.Rows[i].Cells["AnalysisCode5"] == null ? "" : dgv.Rows[i].Cells["AnalysisCode5"].Value.ToString();

                                if (dgv.Columns["AnalysisCode5"].DefaultCellStyle.BackColor == Color.Aqua)
                                {
                                    re.Actions.AnalysisCode5 = re.AnalysisCode5;
                                    re.AnalysisCode5 = "";
                                }
                            }
                            if ((dgv.Columns["AnalysisCode6"] != null) && (dgv.Columns["AnalysisCode6"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("AnalysisCode6,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.AnalysisCode6 = invName;
                                    dgv.Rows[i].Cells["AnalysisCode6"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AnalysisCode6,"))
                                {
                                    re.AnalysisCode6 = null;
                                    dgv.Rows[i].Cells["AnalysisCode6"].Value = null;
                                }
                                else
                                    re.AnalysisCode6 = dgv.Rows[i].Cells["AnalysisCode6"] == null ? "" : dgv.Rows[i].Cells["AnalysisCode6"].Value.ToString();

                                if (dgv.Columns["AnalysisCode6"].DefaultCellStyle.BackColor == Color.Aqua)
                                {
                                    re.Actions.AnalysisCode6 = re.AnalysisCode6;
                                    re.AnalysisCode6 = "";
                                }
                            }
                            if ((dgv.Columns["AnalysisCode7"] != null) && (dgv.Columns["AnalysisCode7"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("AnalysisCode7,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.AnalysisCode7 = invName;
                                    dgv.Rows[i].Cells["AnalysisCode7"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AnalysisCode7,"))
                                {
                                    re.AnalysisCode7 = null;
                                    dgv.Rows[i].Cells["AnalysisCode7"].Value = null;
                                }
                                else
                                    re.AnalysisCode7 = dgv.Rows[i].Cells["AnalysisCode7"] == null ? "" : dgv.Rows[i].Cells["AnalysisCode7"].Value.ToString();

                                if (dgv.Columns["AnalysisCode7"].DefaultCellStyle.BackColor == Color.Aqua)
                                {
                                    re.Actions.AnalysisCode7 = re.AnalysisCode7;
                                    re.AnalysisCode7 = "";
                                }
                            }
                            if ((dgv.Columns["AnalysisCode8"] != null) && (dgv.Columns["AnalysisCode8"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("AnalysisCode8,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.AnalysisCode8 = invName;
                                    dgv.Rows[i].Cells["AnalysisCode8"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AnalysisCode8,"))
                                {
                                    re.AnalysisCode8 = null;
                                    dgv.Rows[i].Cells["AnalysisCode8"].Value = null;
                                }
                                else
                                    re.AnalysisCode8 = dgv.Rows[i].Cells["AnalysisCode8"] == null ? "" : dgv.Rows[i].Cells["AnalysisCode8"].Value.ToString();

                                if (dgv.Columns["AnalysisCode8"].DefaultCellStyle.BackColor == Color.Aqua)
                                {
                                    re.Actions.AnalysisCode8 = re.AnalysisCode8;
                                    re.AnalysisCode8 = "";
                                }
                            }
                            if ((dgv.Columns["AnalysisCode9"] != null) && (dgv.Columns["AnalysisCode9"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("AnalysisCode9,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.AnalysisCode9 = invName;
                                    dgv.Rows[i].Cells["AnalysisCode9"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AnalysisCode9,"))
                                {
                                    re.AnalysisCode9 = null;
                                    dgv.Rows[i].Cells["AnalysisCode9"].Value = null;
                                }
                                else
                                    re.AnalysisCode9 = dgv.Rows[i].Cells["AnalysisCode9"] == null ? "" : dgv.Rows[i].Cells["AnalysisCode9"].Value.ToString();

                                if (dgv.Columns["AnalysisCode9"].DefaultCellStyle.BackColor == Color.Aqua)
                                {
                                    re.Actions.AnalysisCode9 = re.AnalysisCode9;
                                    re.AnalysisCode9 = "";
                                }
                            }
                            if ((dgv.Columns["AnalysisCode10"] != null) && (dgv.Columns["AnalysisCode10"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("AnalysisCode10,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.AnalysisCode10 = invName;
                                    dgv.Rows[i].Cells["AnalysisCode10"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AnalysisCode10,"))
                                {
                                    re.AnalysisCode10 = null;
                                    dgv.Rows[i].Cells["AnalysisCode10"].Value = null;
                                }
                                else
                                    re.AnalysisCode10 = dgv.Rows[i].Cells["AnalysisCode10"] == null ? "" : dgv.Rows[i].Cells["AnalysisCode10"].Value.ToString();

                                if (dgv.Columns["AnalysisCode10"].DefaultCellStyle.BackColor == Color.Aqua)
                                {
                                    re.Actions.AnalysisCode10 = re.AnalysisCode10;
                                    re.AnalysisCode10 = "";
                                }
                            }
                            if ((dgv.Columns["AccountCode"] != null) && (dgv.Columns["AccountCode"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("AccountCode,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.AccountCode = invName;
                                    dgv.Rows[i].Cells["AccountCode"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AccountCode,"))
                                {
                                    re.AccountCode = null;
                                    dgv.Rows[i].Cells["AccountCode"].Value = null;
                                }
                                else
                                    re.AccountCode = dgv.Rows[i].Cells["AccountCode"] == null ? "" : dgv.Rows[i].Cells["AccountCode"].Value.ToString();
                            }
                            if ((dgv.Columns["AccountingPeriod"] != null) && (dgv.Columns["AccountingPeriod"].Visible == true))
                                re.AccountingPeriod = re.AccountingPeriod;
                            if ((dgv.Columns["GeneralDescription1"] != null) && (dgv.Columns["GeneralDescription1"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc1,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription1 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription1"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc1,"))
                                {
                                    re.Actions.GeneralDescription1 = null;
                                    dgv.Rows[i].Cells["GeneralDescription1"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription1 = dgv.Rows[i].Cells["GeneralDescription1"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription1"].Value.ToString();
                            }
                            if ((dgv.Columns["GeneralDescription2"] != null) && (dgv.Columns["GeneralDescription2"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc2,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription2 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription2"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc2,"))
                                {
                                    re.Actions.GeneralDescription2 = null;
                                    dgv.Rows[i].Cells["GeneralDescription2"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription2 = dgv.Rows[i].Cells["GeneralDescription2"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription2"].Value.ToString();
                            }
                            if ((dgv.Columns["GeneralDescription3"] != null) && (dgv.Columns["GeneralDescription3"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc3,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription3 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription3"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc3,"))
                                {
                                    re.Actions.GeneralDescription3 = null;
                                    dgv.Rows[i].Cells["GeneralDescription3"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription3 = dgv.Rows[i].Cells["GeneralDescription3"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription3"].Value.ToString();
                            }
                            if ((dgv.Columns["GeneralDescription4"] != null) && (dgv.Columns["GeneralDescription4"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc4,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription4 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription4"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc4,"))
                                {
                                    re.Actions.GeneralDescription4 = null;
                                    dgv.Rows[i].Cells["GeneralDescription4"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription4 = dgv.Rows[i].Cells["GeneralDescription4"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription4"].Value.ToString();
                            }
                            if ((dgv.Columns["GeneralDescription5"] != null) && (dgv.Columns["GeneralDescription5"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc5,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription5 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription5"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc5,"))
                                {
                                    re.Actions.GeneralDescription5 = null;
                                    dgv.Rows[i].Cells["GeneralDescription5"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription5 = dgv.Rows[i].Cells["GeneralDescription5"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription5"].Value.ToString();
                            }
                            if ((dgv.Columns["GeneralDescription6"] != null) && (dgv.Columns["GeneralDescription6"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc6,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription6 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription6"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc6,"))
                                {
                                    re.Actions.GeneralDescription6 = null;
                                    dgv.Rows[i].Cells["GeneralDescription6"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription6 = dgv.Rows[i].Cells["GeneralDescription6"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription6"].Value.ToString();
                            }
                            if ((dgv.Columns["GeneralDescription7"] != null) && (dgv.Columns["GeneralDescription7"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc7,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription7 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription7"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc7,"))
                                {
                                    re.Actions.GeneralDescription7 = null;
                                    dgv.Rows[i].Cells["GeneralDescription7"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription7 = dgv.Rows[i].Cells["GeneralDescription7"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription7"].Value.ToString();
                            }
                            if ((dgv.Columns["GeneralDescription8"] != null) && (dgv.Columns["GeneralDescription8"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc8,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription8 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription8"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc8,"))
                                {
                                    re.Actions.GeneralDescription8 = null;
                                    dgv.Rows[i].Cells["GeneralDescription8"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription8 = dgv.Rows[i].Cells["GeneralDescription8"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription8"].Value.ToString();
                            }
                            if ((dgv.Columns["GeneralDescription9"] != null) && (dgv.Columns["GeneralDescription9"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc9,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription9 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription9"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc9,"))
                                {
                                    re.Actions.GeneralDescription9 = null;
                                    dgv.Rows[i].Cells["GeneralDescription9"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription9 = dgv.Rows[i].Cells["GeneralDescription9"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription9"].Value.ToString();
                            }
                            if ((dgv.Columns["GeneralDescription10"] != null) && (dgv.Columns["GeneralDescription10"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc10,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription10 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription10"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc10,"))
                                {
                                    re.Actions.GeneralDescription10 = null;
                                    dgv.Rows[i].Cells["GeneralDescription10"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription10 = dgv.Rows[i].Cells["GeneralDescription10"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription10"].Value.ToString();
                            }
                            if ((dgv.Columns["GeneralDescription11"] != null) && (dgv.Columns["GeneralDescription11"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc11,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription11 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription11"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc11,"))
                                {
                                    re.Actions.GeneralDescription11 = null;
                                    dgv.Rows[i].Cells["GeneralDescription11"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription11 = dgv.Rows[i].Cells["GeneralDescription11"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription11"].Value.ToString();
                            }
                            if ((dgv.Columns["GeneralDescription12"] != null) && (dgv.Columns["GeneralDescription12"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc12,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription12 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription12"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc12,"))
                                {
                                    re.Actions.GeneralDescription12 = null;
                                    dgv.Rows[i].Cells["GeneralDescription12"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription12 = dgv.Rows[i].Cells["GeneralDescription12"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription12"].Value.ToString();
                            }
                            if ((dgv.Columns["GeneralDescription13"] != null) && (dgv.Columns["GeneralDescription13"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc13,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription13 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription13"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc13,"))
                                {
                                    re.Actions.GeneralDescription13 = null;
                                    dgv.Rows[i].Cells["GeneralDescription13"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription13 = dgv.Rows[i].Cells["GeneralDescription13"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription13"].Value.ToString();
                            }
                            if ((dgv.Columns["GeneralDescription14"] != null) && (dgv.Columns["GeneralDescription14"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc14,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription14 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription14"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc14,"))
                                {
                                    re.Actions.GeneralDescription14 = null;
                                    dgv.Rows[i].Cells["GeneralDescription14"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription14 = dgv.Rows[i].Cells["GeneralDescription14"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription14"].Value.ToString();
                            }
                            if ((dgv.Columns["GeneralDescription15"] != null) && (dgv.Columns["GeneralDescription15"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc15,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription15 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription15"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc15,"))
                                {
                                    re.Actions.GeneralDescription15 = null;
                                    dgv.Rows[i].Cells["GeneralDescription15"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription15 = dgv.Rows[i].Cells["GeneralDescription15"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription15"].Value.ToString();
                            }
                            if ((dgv.Columns["GeneralDescription16"] != null) && (dgv.Columns["GeneralDescription16"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc16,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription16 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription16"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc16,"))
                                {
                                    re.Actions.GeneralDescription16 = null;
                                    dgv.Rows[i].Cells["GeneralDescription16"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription16 = dgv.Rows[i].Cells["GeneralDescription16"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription16"].Value.ToString();
                            }
                            if ((dgv.Columns["GeneralDescription17"] != null) && (dgv.Columns["GeneralDescription17"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc17,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription17 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription17"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc17,"))
                                {
                                    re.Actions.GeneralDescription17 = null;
                                    dgv.Rows[i].Cells["GeneralDescription17"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription17 = dgv.Rows[i].Cells["GeneralDescription17"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription17"].Value.ToString();
                            }
                            if ((dgv.Columns["GeneralDescription18"] != null) && (dgv.Columns["GeneralDescription18"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc18,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription18 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription18"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc18,"))
                                {
                                    re.Actions.GeneralDescription18 = null;
                                    dgv.Rows[i].Cells["GeneralDescription18"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription18 = dgv.Rows[i].Cells["GeneralDescription18"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription18"].Value.ToString();
                            }
                            if ((dgv.Columns["GeneralDescription19"] != null) && (dgv.Columns["GeneralDescription19"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc19,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription19 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription19"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc19,"))
                                {
                                    re.Actions.GeneralDescription19 = null;
                                    dgv.Rows[i].Cells["GeneralDescription19"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription19 = dgv.Rows[i].Cells["GeneralDescription19"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription19"].Value.ToString();
                            }
                            if ((dgv.Columns["GeneralDescription20"] != null) && (dgv.Columns["GeneralDescription20"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc20,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription20 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription20"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc20,"))
                                {
                                    re.Actions.GeneralDescription20 = null;
                                    dgv.Rows[i].Cells["GeneralDescription20"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription20 = dgv.Rows[i].Cells["GeneralDescription20"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription20"].Value.ToString();
                            }
                            if ((dgv.Columns["GeneralDescription21"] != null) && (dgv.Columns["GeneralDescription21"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc21,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription21 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription21"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc21,"))
                                {
                                    re.Actions.GeneralDescription21 = null;
                                    dgv.Rows[i].Cells["GeneralDescription21"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription21 = dgv.Rows[i].Cells["GeneralDescription21"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription21"].Value.ToString();
                            }
                            if ((dgv.Columns["GeneralDescription22"] != null) && (dgv.Columns["GeneralDescription22"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc22,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription22 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription22"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc22,"))
                                {
                                    re.Actions.GeneralDescription22 = null;
                                    dgv.Rows[i].Cells["GeneralDescription22"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription22 = dgv.Rows[i].Cells["GeneralDescription22"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription22"].Value.ToString();
                            }
                            if ((dgv.Columns["GeneralDescription23"] != null) && (dgv.Columns["GeneralDescription23"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc23,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription23 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription23"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc23,"))
                                {
                                    re.Actions.GeneralDescription23 = null;
                                    dgv.Rows[i].Cells["GeneralDescription23"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription23 = dgv.Rows[i].Cells["GeneralDescription23"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription23"].Value.ToString();
                            }
                            if ((dgv.Columns["GeneralDescription24"] != null) && (dgv.Columns["GeneralDescription24"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc24,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription24 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription24"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc24,"))
                                {
                                    re.Actions.GeneralDescription24 = null;
                                    dgv.Rows[i].Cells["GeneralDescription24"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription24 = dgv.Rows[i].Cells["GeneralDescription24"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription24"].Value.ToString();
                            }
                            if ((dgv.Columns["GeneralDescription25"] != null) && (dgv.Columns["GeneralDescription25"].Visible == true))
                            {
                                if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains("GenDesc25,"))
                                {
                                    int? ii = 0;
                                    string invName = string.Empty;
                                    string prefix = string.Empty;
                                    try
                                    {
                                        ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                    }
                                    catch
                                    {
                                    }
                                    re.Actions.GeneralDescription25 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription25"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc25,"))
                                {
                                    re.Actions.GeneralDescription25 = null;
                                    dgv.Rows[i].Cells["GeneralDescription25"].Value = null;
                                }
                                else
                                    re.Actions.GeneralDescription25 = dgv.Rows[i].Cells["GeneralDescription25"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription25"].Value.ToString();
                            }
                            if ((dgv.Columns["TransactionDate"] != null) && (dgv.Columns["TransactionDate"].Visible == true))
                            {
                                //re.Actions.TransactionDate = re.TransactionDate;//"ddMMyyyy"
                            }
                            if ((dgv.Columns["TransactionAmount"] != null) && (dgv.Columns["TransactionAmount"].Visible == true))
                                re.TransactionAmount = double.Parse(string.IsNullOrEmpty(dgv.Rows[i].Cells["TransactionAmount"] == null ? "" : dgv.Rows[i].Cells["TransactionAmount"].Value.ToString()) ? "0" : dgv.Rows[i].Cells["TransactionAmount"] == null ? "" : dgv.Rows[i].Cells["TransactionAmount"].Value.ToString()).ToString("0.000");
                            if ((dgv.Columns["CurrencyCode"] != null) && (dgv.Columns["CurrencyCode"].Visible == true))
                                re.CurrencyCode = dgv.Rows[i].Cells["CurrencyCode"] == null ? "" : dgv.Rows[i].Cells["CurrencyCode"].Value.ToString();
                            //if (dgv.Columns["DebitCredit"] != null)

                            if ((dgv.Columns["BaseAmount"] != null) && (dgv.Columns["BaseAmount"].Visible == true))
                                re.BaseAmount = dgv.Rows[i].Cells["BaseAmount"] == null ? "" : dgv.Rows[i].Cells["BaseAmount"].Value.ToString();
                            if ((dgv.Columns["Base2ReportingAmount"] != null) && (dgv.Columns["Base2ReportingAmount"].Visible == true))
                                re.Base2ReportingAmount = dgv.Rows[i].Cells["Base2ReportingAmount"] == null ? "" : dgv.Rows[i].Cells["Base2ReportingAmount"].Value.ToString();
                            if ((dgv.Columns["Value4Amount"] != null) && (dgv.Columns["Value4Amount"].Visible == true))
                                re.Value4Amount = dgv.Rows[i].Cells["Value4Amount"] == null ? "" : dgv.Rows[i].Cells["Value4Amount"].Value.ToString();

                            if (!string.IsNullOrEmpty(re.Base2ReportingAmount))
                                re.Base2ReportingAmount = re.Base2ReportingAmount.Replace("-", "");
                            if (!string.IsNullOrEmpty(re.TransactionAmount))
                                re.TransactionAmount = re.TransactionAmount.Replace("-", "");
                            if (!string.IsNullOrEmpty(re.Value4Amount))
                                re.Value4Amount = re.Value4Amount.Replace("-", "");
                            if (!string.IsNullOrEmpty(re.BaseAmount))
                                re.BaseAmount = re.BaseAmount.Replace("-", "");

                            re.Messages = new Common2.Messages();
                            re.Messages.Message = new Common2.Message();
                            re.Messages.Message.Exception = ".";
                            re.Messages.Message.UserText = ".";
                            re.Messages.Message.Application = new Common2.Application();
                            re.Messages.Message.Application.Component = ".";
                            re.Messages.Message.Application.DataItem = ".";
                            re.Messages.Message.Application.Driver = ".";
                            re.Messages.Message.Application.Item = ".";
                            re.Messages.Message.Application.LastMethod = ".";
                            re.Messages.Message.Application.Message = ".";
                            re.Messages.Message.Application.MessageNumber = ".";
                            re.Messages.Message.Application.Method = ".";
                            re.Messages.Message.Application.Type = ".";
                            re.Messages.Message.Application.Value = ".";
                            re.Messages.Message.Application.Version = ".";
                            newlist.Add(re);
                        }
                        string script = string.Empty;
                        IEnumerable<IGrouping<string, ExcelAddIn4.Common2.LedgerUpdate>> query = newlist.GroupBy(pet => pet.Ledger, pet => pet);
                        foreach (IGrouping<string, ExcelAddIn4.Common2.LedgerUpdate> info in query)
                        {
                            List<ExcelAddIn4.Common2.LedgerUpdate> sl = info.ToList<ExcelAddIn4.Common2.LedgerUpdate>();
                            script += ft.GetTransUpdXMLScript(sl) + "\r\n***\r\n";
                        }
                        this.txtXML.Text = script;
                    }
                }
                catch (Exception ex)
                {
                    this.txtXML.Text = ex.Message;
                }
            }
        }
    }
}
