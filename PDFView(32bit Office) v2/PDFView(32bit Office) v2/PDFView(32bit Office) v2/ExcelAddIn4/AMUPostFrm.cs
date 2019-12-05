/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<AMUPostFrm>   
 * Author：Peter.uhm  (yanb@shinetechchina.com)
 * Modify date：2015.10
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
    public partial class AMUPostFrm : Form
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
        public AMUPostFrm()
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
                dgv = ft.IniXMLFormGrdForAllocationMakerUpd();
                dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
                dgv.AutoGenerateColumns = false;
                dgv.ColumnHeadersHeight = 40;
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dgv.Dock = DockStyle.Fill;
                dgv.Visible = true;
                dgv.ColumnHeaderMouseClick += new DataGridViewCellMouseEventHandler(DataGridView1_CellMouseDown);
                dgv.RowPostPaint += new DataGridViewRowPostPaintEventHandler(dgv_RowPostPaint);
                //dgv.DataBindingComplete += new DataGridViewBindingCompleteEventHandler(ft.XMLFormdataGridView_DataBindingComplete);
                BindData();
                //Remove the cloumns doesn't contain data
                DataTable tb1 = ft.ToDataTable((List<AMUEntityForSave>)this.dgv.DataSource);
                for (int i = 0; i < tb1.Columns.Count; i++)
                {
                    string str = string.Empty;
                    for (int j = 0; j < tb1.Rows.Count; j++)
                    {
                        str += tb1.Rows[j][i].ToString();
                        if (tb1.Rows[j][i].ToString().ToUpper() == "[SEQUENCE]")
                            sequenceNumbering = tb1.Columns[i].ColumnName + "," + sequenceNumbering;
                    }
                    // this place need to confirm
                    if ((string.IsNullOrEmpty(str) && tb1.Columns[i].ColumnName != "DebitCredit" && tb1.Columns[i].ColumnName != "DetailLad") || ((string.IsNullOrEmpty(str.Replace("0", "").Replace(".", "")) ? true : false) && tb1.Columns[i].ColumnName != "DebitCredit" && tb1.Columns[i].ColumnName != "DetailLad"))
                        dgv.Columns.Remove(tb1.Columns[i].ColumnName);

                    if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(tb1.Columns[i].ColumnName))
                        dgv.Columns.Remove(tb1.Columns[i].ColumnName);
                }
                try
                {
                    dgv.Columns["AllocationMarker"].DefaultCellStyle.BackColor = Color.Aqua;
                }
                catch
                {
                    MessageBox.Show("Update(Allocate) criteria can't be empty!", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                this.tabPage1.Controls.Add(dgv);
            }
            catch { }
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
                this.dgv.DataSource = Ribbon2.outputPane.AMUFinallist;
            }
            catch (Exception ex)
            {
                AMUPostFrm.richTextBox1.Text += ex.Message;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void btnPost_Click(object sender, EventArgs e)
        {
            DoPost(sender);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <returns></returns>
        private bool DoPost(object sender)
        {
            //tabControl1_SelectedIndexChanged(null, null);
            string sSSCVoucher = "";
            SunSystems.Connect.Client.SecurityManager oSecMan = new SunSystems.Connect.Client.SecurityManager(SessionInfo.UserInfo.SunUserIP);
            ////http://95.138.187.185:81/SecurityWebServer/Login.aspx?redirect=http://95.138.187.185:8080/ssc/login.jsp
            try
            {
                oSecMan.Login(SessionInfo.UserInfo.SunUserID, SessionInfo.UserInfo.SunUserPass);
                if (oSecMan.Authorised)
                {
                    sSSCVoucher = oSecMan.Voucher;
                    //label1.Content = sSSCVoucher;RSI, RicoSimp2
                }
                else
                {
                    this.textBox1.Text = "SunSystems Server is not exist or Password for user is incorrect.";
                    MessageBox.Show(this.textBox1.Text, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }
            }
            catch (Exception ex)
            {
                this.textBox1.Text = "An error occurred in validation " + oSecMan.ErrorMessage + ex;
                MessageBox.Show(this.textBox1.Text, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LogHelper.WriteLog(typeof(AMUPostFrm), this.textBox1.Text);
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
                            if (SessionInfo.UserInfo.AllocationMakerType == "Allocate")
                                strResu = method.Invoke(o, new object[] { sSSCVoucher, null, "AccountAllocations", "Allocate", null, ss });
                            else
                                strResu = method.Invoke(o, new object[] { sSSCVoucher, null, "AllocationMarkerUpdate", "AmendMarker", null, ss });

                            this.textBox1.Text += GetErrorLines(strResu.ToString()) + "\r\n***\r\n";
                        }
                    }
                    PostErrorFrm pef = new PostErrorFrm(this.textBox1.Text.Replace("This line has been rejected due to errors in other lines or posting options", ""));
                    pef.ShowDialog();
                    if (!string.IsNullOrEmpty(this.textBox1.Text.Trim()))
                    {
                        return false;
                    }
                    else
                    {
                        MessageBox.Show("Update Successful", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                MessageBox.Show(ex.ToString(), "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LogHelper.WriteLog(typeof(Ribbon2), ex.Message + "AllocationMarker Update error");
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
            XmlNodeList xnl;
            XmlNodeList xnl2;
            if (SessionInfo.UserInfo.AllocationMakerType == "Allocate")
            {
                xnl = Xo.XmlDoc.SelectNodes("//AccountAllocations/Messages/Message");
                xnl2 = Xo.XmlDoc.SelectNodes("//AccountAllocations/Messages/Message/UserText");
            }
            else
            {
                xnl = Xo.XmlDoc.SelectNodes("//AllocationMarkers/Messages/Message");
                xnl2 = Xo.XmlDoc.SelectNodes("//AllocationMarkers/Messages/Message/UserText");
                XmlNode xnl3 = Xo.XmlDoc.SelectSingleNode("//BudgetCode");
                error += "Ledger : " + xnl3.InnerText + "\r\n";
            }
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
        /// <returns></returns>
        private string GetCriteriaStr()
        {
            string str = string.Empty;
            if (!string.IsNullOrEmpty(SessionInfo.UserInfo.Criteria1))
                str += SessionInfo.UserInfo.Criteria1 + " : " + SessionInfo.UserInfo.CellReference1 + "\r\n";
            if (!string.IsNullOrEmpty(SessionInfo.UserInfo.Criteria2))
                str += SessionInfo.UserInfo.Criteria2 + " : " + SessionInfo.UserInfo.CellReference2 + "\r\n";
            if (!string.IsNullOrEmpty(SessionInfo.UserInfo.Criteria3))
                str += SessionInfo.UserInfo.Criteria3 + " : " + SessionInfo.UserInfo.CellReference3 + "\r\n";
            if (!string.IsNullOrEmpty(SessionInfo.UserInfo.Criteria4))
                str += SessionInfo.UserInfo.Criteria4 + " : " + SessionInfo.UserInfo.CellReference4 + "\r\n";
            if (!string.IsNullOrEmpty(SessionInfo.UserInfo.Criteria5))
                str += SessionInfo.UserInfo.Criteria5 + " : " + SessionInfo.UserInfo.CellReference5 + "\r\n";

            return str;
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
                        if (SessionInfo.UserInfo.AllocationMakerType == "Allocate")
                        {
                            List<ExcelAddIn4.Common3.AccountAllocations> newlist = new List<ExcelAddIn4.Common3.AccountAllocations>();
                            for (int i = 0; i < dgv.Rows.Count; i++)
                            {
                                ExcelAddIn4.Common3.AccountAllocations re = new ExcelAddIn4.Common3.AccountAllocations();
                                re.NewSettings = new ExcelAddIn4.Common3.NewSettings();
                                re.SelectionCriteria = new Common3.SelectionCriteria();
                                if ((dgv.Columns["JournalNumber"] != null) && (dgv.Columns["JournalNumber"].Visible == true))
                                {
                                    re.SelectionCriteria.JournalNumberFrom = dgv.Rows[i].Cells["JournalNumber"] == null ? "" : dgv.Rows[i].Cells["JournalNumber"].Value.ToString();
                                    re.SelectionCriteria.JournalNumberTo = re.SelectionCriteria.JournalNumberFrom;
                                }
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
                                        re.SelectionCriteria.Ledger = invName;
                                        dgv.Rows[i].Cells["Ledger"].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("Ledger,"))
                                    {
                                        re.SelectionCriteria.Ledger = null;
                                        dgv.Rows[i].Cells["Ledger"].Value = null;
                                    }
                                    else
                                        re.SelectionCriteria.Ledger = dgv.Rows[i].Cells["Ledger"] == null ? "" : dgv.Rows[i].Cells["Ledger"].Value.ToString();
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
                                        re.SelectionCriteria.AccountCode = invName;
                                        dgv.Rows[i].Cells["AccountCode"].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AccountCode,"))
                                    {
                                        re.SelectionCriteria.AccountCode = null;
                                        dgv.Rows[i].Cells["AccountCode"].Value = null;
                                    }
                                    else
                                        re.SelectionCriteria.AccountCode = dgv.Rows[i].Cells["AccountCode"] == null ? "" : dgv.Rows[i].Cells["AccountCode"].Value.ToString();
                                }

                                if ((dgv.Columns["AccountingPeriod"] != null) && (dgv.Columns["AccountingPeriod"].Visible == true))
                                {
                                    re.SelectionCriteria.AccountingPeriodFrom = dgv.Rows[i].Cells["AccountingPeriod"] == null ? "" : dgv.Rows[i].Cells["AccountingPeriod"].Value.ToString();
                                    re.SelectionCriteria.AccountingPeriodTo = re.SelectionCriteria.AccountingPeriodFrom;
                                }

                                if ((dgv.Columns["TransactionDate"] != null) && (dgv.Columns["TransactionDate"].Visible == true))
                                {
                                    bool result;
                                    DateTime r;
                                    result = DateTime.TryParse(dgv.Rows[i].Cells["TransactionDate"].Value.ToString(), out r);
                                    if (result)
                                    {
                                        re.SelectionCriteria.TransactionDateFrom = r.ToString("yyyyMMdd").Substring(6, 2) + r.ToString("yyyyMMdd").Substring(4, 2) + r.ToString("yyyyMMdd").Substring(0, 4);
                                        re.SelectionCriteria.TransactionDateTo = re.SelectionCriteria.TransactionDateFrom;
                                    }
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
                                        re.SelectionCriteria.JournalTypeFrom = invName;
                                        dgv.Rows[i].Cells["JournalType"].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("JournalType,"))
                                    {
                                        re.SelectionCriteria.JournalTypeFrom = null;
                                        dgv.Rows[i].Cells["JournalType"].Value = null;
                                    }
                                    else
                                        re.SelectionCriteria.JournalTypeFrom = dgv.Rows[i].Cells["JournalType"] == null ? "" : dgv.Rows[i].Cells["JournalType"].Value.ToString();

                                    re.SelectionCriteria.JournalTypeTo = re.SelectionCriteria.JournalTypeFrom;
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
                                        re.SelectionCriteria.TransactionReferenceFrom = invName;
                                        dgv.Rows[i].Cells["TransactionReference"].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("TransactionReference,"))
                                    {
                                        re.SelectionCriteria.TransactionReferenceFrom = null;
                                        dgv.Rows[i].Cells["TransactionReference"].Value = null;
                                    }
                                    else
                                        re.SelectionCriteria.TransactionReferenceFrom = dgv.Rows[i].Cells["TransactionReference"] == null ? "" : dgv.Rows[i].Cells["TransactionReference"].Value.ToString();

                                    re.SelectionCriteria.TransactionReferenceTo = re.SelectionCriteria.TransactionReferenceFrom;
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
                                        re.NewSettings.AllocationMarker = invName;
                                        dgv.Rows[i].Cells["AllocationMarker"].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AllocationMarker,"))
                                    {
                                        re.NewSettings.AllocationMarker = null;
                                        dgv.Rows[i].Cells["AllocationMarker"].Value = null;
                                    }
                                    else
                                        re.NewSettings.AllocationMarker = dgv.Rows[i].Cells["AllocationMarker"] == null ? "" : dgv.Rows[i].Cells["AllocationMarker"].Value.ToString();

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
                                        re.SelectionCriteria.TransactionAnalysis1From = invName;
                                        dgv.Rows[i].Cells["AnalysisCode1"].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AnalysisCode1,"))
                                    {
                                        re.SelectionCriteria.TransactionAnalysis1From = null;
                                        dgv.Rows[i].Cells["AnalysisCode1"].Value = null;
                                    }
                                    else
                                        re.SelectionCriteria.TransactionAnalysis1From = dgv.Rows[i].Cells["AnalysisCode1"] == null ? "" : dgv.Rows[i].Cells["AnalysisCode1"].Value.ToString();

                                    re.SelectionCriteria.TransactionAnalysis1To = re.SelectionCriteria.TransactionAnalysis1From;
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
                                        re.SelectionCriteria.TransactionAnalysis2From = invName;
                                        dgv.Rows[i].Cells["AnalysisCode2"].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AnalysisCode2,"))
                                    {
                                        re.SelectionCriteria.TransactionAnalysis2From = null;
                                        dgv.Rows[i].Cells["AnalysisCode2"].Value = null;
                                    }
                                    else
                                        re.SelectionCriteria.TransactionAnalysis2From = dgv.Rows[i].Cells["AnalysisCode2"] == null ? "" : dgv.Rows[i].Cells["AnalysisCode2"].Value.ToString();

                                    re.SelectionCriteria.TransactionAnalysis2To = re.SelectionCriteria.TransactionAnalysis2From;
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
                                        re.SelectionCriteria.TransactionAnalysis3From = invName;
                                        dgv.Rows[i].Cells["AnalysisCode3"].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AnalysisCode3,"))
                                    {
                                        re.SelectionCriteria.TransactionAnalysis3From = null;
                                        dgv.Rows[i].Cells["AnalysisCode3"].Value = null;
                                    }
                                    else
                                        re.SelectionCriteria.TransactionAnalysis3From = dgv.Rows[i].Cells["AnalysisCode3"] == null ? "" : dgv.Rows[i].Cells["AnalysisCode3"].Value.ToString();

                                    re.SelectionCriteria.TransactionAnalysis3To = re.SelectionCriteria.TransactionAnalysis3From;
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
                                        re.SelectionCriteria.TransactionAnalysis4From = invName;
                                        dgv.Rows[i].Cells["AnalysisCode4"].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AnalysisCode4,"))
                                    {
                                        re.SelectionCriteria.TransactionAnalysis4From = null;
                                        dgv.Rows[i].Cells["AnalysisCode4"].Value = null;
                                    }
                                    else
                                        re.SelectionCriteria.TransactionAnalysis4From = dgv.Rows[i].Cells["AnalysisCode4"] == null ? "" : dgv.Rows[i].Cells["AnalysisCode4"].Value.ToString();

                                    re.SelectionCriteria.TransactionAnalysis4To = re.SelectionCriteria.TransactionAnalysis4From;
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
                                        re.SelectionCriteria.TransactionAnalysis5From = invName;
                                        dgv.Rows[i].Cells["AnalysisCode5"].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AnalysisCode5,"))
                                    {
                                        re.SelectionCriteria.TransactionAnalysis5From = null;
                                        dgv.Rows[i].Cells["AnalysisCode5"].Value = null;
                                    }
                                    else
                                        re.SelectionCriteria.TransactionAnalysis5From = dgv.Rows[i].Cells["AnalysisCode5"] == null ? "" : dgv.Rows[i].Cells["AnalysisCode5"].Value.ToString();

                                    re.SelectionCriteria.TransactionAnalysis5To = re.SelectionCriteria.TransactionAnalysis5From;
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
                                        re.SelectionCriteria.TransactionAnalysis6From = invName;
                                        dgv.Rows[i].Cells["AnalysisCode6"].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AnalysisCode6,"))
                                    {
                                        re.SelectionCriteria.TransactionAnalysis6From = null;
                                        dgv.Rows[i].Cells["AnalysisCode6"].Value = null;
                                    }
                                    else
                                        re.SelectionCriteria.TransactionAnalysis6From = dgv.Rows[i].Cells["AnalysisCode6"] == null ? "" : dgv.Rows[i].Cells["AnalysisCode6"].Value.ToString();

                                    re.SelectionCriteria.TransactionAnalysis6To = re.SelectionCriteria.TransactionAnalysis6From;
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
                                        re.SelectionCriteria.TransactionAnalysis7From = invName;
                                        dgv.Rows[i].Cells["AnalysisCode7"].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AnalysisCode7,"))
                                    {
                                        re.SelectionCriteria.TransactionAnalysis7From = null;
                                        dgv.Rows[i].Cells["AnalysisCode7"].Value = null;
                                    }
                                    else
                                        re.SelectionCriteria.TransactionAnalysis7From = dgv.Rows[i].Cells["AnalysisCode7"] == null ? "" : dgv.Rows[i].Cells["AnalysisCode7"].Value.ToString();

                                    re.SelectionCriteria.TransactionAnalysis7To = re.SelectionCriteria.TransactionAnalysis7From;
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
                                        re.SelectionCriteria.TransactionAnalysis8From = invName;
                                        dgv.Rows[i].Cells["AnalysisCode8"].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AnalysisCode8,"))
                                    {
                                        re.SelectionCriteria.TransactionAnalysis8From = null;
                                        dgv.Rows[i].Cells["AnalysisCode8"].Value = null;
                                    }
                                    else
                                        re.SelectionCriteria.TransactionAnalysis8From = dgv.Rows[i].Cells["AnalysisCode8"] == null ? "" : dgv.Rows[i].Cells["AnalysisCode8"].Value.ToString();

                                    re.SelectionCriteria.TransactionAnalysis8To = re.SelectionCriteria.TransactionAnalysis8From;
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
                                        re.SelectionCriteria.TransactionAnalysis9From = invName;
                                        dgv.Rows[i].Cells["AnalysisCode9"].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AnalysisCode9,"))
                                    {
                                        re.SelectionCriteria.TransactionAnalysis9From = null;
                                        dgv.Rows[i].Cells["AnalysisCode9"].Value = null;
                                    }
                                    else
                                        re.SelectionCriteria.TransactionAnalysis9From = dgv.Rows[i].Cells["AnalysisCode9"] == null ? "" : dgv.Rows[i].Cells["AnalysisCode9"].Value.ToString();

                                    re.SelectionCriteria.TransactionAnalysis9To = re.SelectionCriteria.TransactionAnalysis9From;
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
                                        re.SelectionCriteria.TransactionAnalysis10From = invName;
                                        dgv.Rows[i].Cells["AnalysisCode10"].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AnalysisCode10,"))
                                    {
                                        re.SelectionCriteria.TransactionAnalysis10From = null;
                                        dgv.Rows[i].Cells["AnalysisCode10"].Value = null;
                                    }
                                    else
                                        re.SelectionCriteria.TransactionAnalysis10From = dgv.Rows[i].Cells["AnalysisCode10"] == null ? "" : dgv.Rows[i].Cells["AnalysisCode10"].Value.ToString();

                                    re.SelectionCriteria.TransactionAnalysis10To = re.SelectionCriteria.TransactionAnalysis10From;
                                }

                                re.FlagOptions = new Common3.FlagOptions();
                                re.FlagOptions.AllowClosedOrSuspendedAccountCode = "Y";
                                re.FlagOptions.OverrideAllocations = "Y";

                                re.ActionAllSettings = new Common3.ActionAllSettings();
                                re.ActionAllSettings.AllocationMarkerFrom = ".";
                                re.ActionAllSettings.AllocationMarkerTo = ".";

                                re.Messages = new Common3.Messages();
                                re.Messages.Message = new Common3.Message();
                                re.Messages.Message.Exception = ".";
                                re.Messages.Message.UserText = ".";
                                re.Messages.Message.Application = new Common3.Application();
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
                            IEnumerable<IGrouping<string, ExcelAddIn4.Common3.AccountAllocations>> query = newlist.GroupBy(pet => pet.SelectionCriteria.Ledger, pet => pet);
                            foreach (IGrouping<string, ExcelAddIn4.Common3.AccountAllocations> info in query)
                            {
                                List<ExcelAddIn4.Common3.AccountAllocations> sl = info.ToList<ExcelAddIn4.Common3.AccountAllocations>();
                                script += ft.GetAMUXMLScript(sl) + "\r\n***\r\n";
                            }
                            this.txtXML.Text = script;
                        }
                        else
                        {
                            List<ExcelAddIn4.Common3.AllocationMarkers> newlist = new List<ExcelAddIn4.Common3.AllocationMarkers>();
                            for (int i = 0; i < dgv.Rows.Count; i++)
                            {
                                ExcelAddIn4.Common3.AllocationMarkers re = new ExcelAddIn4.Common3.AllocationMarkers();
                                if ((dgv.Columns["JournalNumber"] != null) && (dgv.Columns["JournalNumber"].Visible == true))
                                {
                                    re.JournalNumber = dgv.Rows[i].Cells["JournalNumber"] == null ? "" : dgv.Rows[i].Cells["JournalNumber"].Value.ToString();
                                }
                                if ((dgv.Columns["JournalLineNumber"] != null) && (dgv.Columns["JournalLineNumber"].Visible == true))
                                {
                                    re.JournalLineNumber = dgv.Rows[i].Cells["JournalLineNumber"] == null ? "" : dgv.Rows[i].Cells["JournalLineNumber"].Value.ToString();
                                }
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
                                        dgv.Rows[i].Cells["AccountCode"].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AccountCode,"))
                                    {
                                        re.AccountRange.AccountCodeFrom = null;
                                        re.AccountRange.AccountCodeTo = null;
                                        dgv.Rows[i].Cells["AccountCode"].Value = null;
                                    }
                                    else
                                    {
                                        re.AccountRange.AccountCodeFrom = dgv.Rows[i].Cells["AccountCode"] == null ? "" : dgv.Rows[i].Cells["AccountCode"].Value.ToString();
                                        re.AccountRange.AccountCodeTo = re.AccountRange.AccountCodeFrom;
                                    }
                                }

                                if ((dgv.Columns["AccountingPeriod"] != null) && (dgv.Columns["AccountingPeriod"].Visible == true))
                                {
                                    re.AccountingPeriod = dgv.Rows[i].Cells["AccountingPeriod"] == null ? "" : dgv.Rows[i].Cells["AccountingPeriod"].Value.ToString();
                                }

                                if ((dgv.Columns["TransactionDate"] != null) && (dgv.Columns["TransactionDate"].Visible == true))
                                {
                                    bool result;
                                    DateTime r;
                                    result = DateTime.TryParse(dgv.Rows[i].Cells["TransactionDate"].Value.ToString(), out r);
                                    if (result)
                                    {
                                        re.TransactionDateRange.TransactionDateFrom = r.ToString("yyyyMMdd").Substring(6, 2) + r.ToString("yyyyMMdd").Substring(4, 2) + r.ToString("yyyyMMdd").Substring(0, 4);
                                        re.TransactionDateRange.TransactionDateTo = re.TransactionDateRange.TransactionDateFrom;
                                    }
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
                                        re.Actions.AllocationMarker = invName;
                                        dgv.Rows[i].Cells["AllocationMarker"].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AllocationMarker,"))
                                    {
                                        re.Actions.AllocationMarker = null;
                                        dgv.Rows[i].Cells["AllocationMarker"].Value = null;
                                    }
                                    else
                                        re.Actions.AllocationMarker = dgv.Rows[i].Cells["AllocationMarker"] == null ? "" : dgv.Rows[i].Cells["AllocationMarker"].Value.ToString();

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
                                }
                                re.Messages = new Common3.Messages();
                                re.Messages.Message = new Common3.Message();
                                re.Messages.Message.Exception = ".";
                                re.Messages.Message.UserText = ".";
                                re.Messages.Message.Application = new Common3.Application();
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
                            IEnumerable<IGrouping<string, ExcelAddIn4.Common3.AllocationMarkers>> query = newlist.GroupBy(pet => pet.Ledger, pet => pet);
                            foreach (IGrouping<string, ExcelAddIn4.Common3.AllocationMarkers> info in query)
                            {
                                List<ExcelAddIn4.Common3.AllocationMarkers> sl = info.ToList<ExcelAddIn4.Common3.AllocationMarkers>();
                                script += ft.GetAMUUpdateXMLScript(sl) + "\r\n***\r\n";
                            }
                            this.txtXML.Text = script;
                        }
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
