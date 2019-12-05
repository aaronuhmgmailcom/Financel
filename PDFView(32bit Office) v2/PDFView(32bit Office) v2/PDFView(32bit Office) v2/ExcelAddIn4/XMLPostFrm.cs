/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<XMLPostFrm>   
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
    public partial class XMLPostFrm : Form
    {
        DataGridView dgv = null;
        string journalNumber = string.Empty;
        string sequenceNumbering = string.Empty;
        /// <summary>
        /// 
        /// </summary>
        internal static Finance_Tools ft
        {
            get { return new Finance_Tools(); }
        }
        public XMLPostFrm()
        {
            InitializeComponent();
        }
        public void bddata()
        {
            try
            {
                this.Text = "Post Journal (" + SessionInfo.UserInfo.CurrentRef + ") - RSystems FinanceTools v2";
                dgv = ft.IniXMLFormGrd();
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
                DataTable tb1 = ft.ToDataTable((List<Specialist>)this.dgv.DataSource);
                for (int i = 0; i < tb1.Columns.Count; i++)
                {
                    string str = string.Empty;
                    for (int j = 0; j < tb1.Rows.Count; j++)
                    {
                        str += tb1.Rows[j][i].ToString();
                        if (tb1.Rows[j][i].ToString().ToUpper() == "[SEQUENCE]")
                            sequenceNumbering = tb1.Columns[i].ColumnName + "," + sequenceNumbering;
                    }
                    if ((string.IsNullOrEmpty(str) && tb1.Columns[i].ColumnName != "DebitCredit" && tb1.Columns[i].ColumnName != "DetailLad"))//|| ((string.IsNullOrEmpty(str.Replace("0", "").Replace(".", "")) ? true : false) && tb1.Columns[i].ColumnName != "DebitCredit" && tb1.Columns[i].ColumnName != "DetailLad")
                        dgv.Columns.Remove(tb1.Columns[i].ColumnName.Replace("GenDesc", "GeneralDescription"));

                    if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(tb1.Columns[i].ColumnName))
                        dgv.Columns.Remove(tb1.Columns[i].ColumnName.Replace("GenDesc", "GeneralDescription"));
                }
                this.tabPage1.Controls.Add(dgv);
            }
            catch
            { }
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
        private int currentColumnIndex = 0;
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
                this.dgv.DataSource = Ribbon2.outputPane.finallist;
            }
            catch (Exception ex)
            {
                XMLPostFrm.richTextBox1.Text += ex.Message;
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
                try
                {
                    SaveHistory();
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            else
            {
                throw new Exception();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="invName"></param>
        private void PopulateCell(string invName)
        {
            if (!string.IsNullOrEmpty(SessionInfo.UserInfo.PopulateCell))
            {
                Globals.ThisAddIn.Application.Rows.get_Range(SessionInfo.UserInfo.PopulateCell).Value = invName;
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
                    //label1.Content = sSSCVoucher;RSI, RicoSimp2
                }
                else
                {
                    this.textBox1.Text = "SunSystems Server is not exist or Password for user is incorrect.";
                    if (Ribbon2.xpf.Visible == true)
                        MessageBox.Show(this.textBox1.Text, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else
                        SessionInfo.UserInfo.GlobalError += "Process:Journal Post(" + SessionInfo.UserInfo.CurrentRef + ") - Fail! " + this.textBox1.Text;

                    return false;
                }
            }
            catch (Exception ex)
            {
                this.textBox1.Text = "An error occurred in validation " + (oSecMan == null ? "" : oSecMan.ErrorMessage) + ex;
                if (Ribbon2.xpf.Visible == true)
                    MessageBox.Show(this.textBox1.Text, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    SessionInfo.UserInfo.GlobalError += "Process:Journal Post(" + SessionInfo.UserInfo.CurrentRef + ") - Fail! " + this.textBox1.Text;
                LogHelper.WriteLog(typeof(XMLPostFrm), this.textBox1.Text);
                return false;
            }
            finally
            {
                oSecMan = null;
            }

            try
            {
                WebClient web = new WebClient();
                Stream stream = web.OpenRead("http://" + SessionInfo.UserInfo.SunUserIP + ":8080/connect/wsdl/ComponentExecutor?wsdl");// 2. Create and format of WSDL document.
                ServiceDescription description = ServiceDescription.Read(stream);// 3. Create a client proxy proxy class.
                ServiceDescriptionImporter importer = new ServiceDescriptionImporter();
                importer.ProtocolName = "Soap"; // The specified access protocol.
                importer.Style = ServiceDescriptionImportStyle.Client; // To generate client proxy.
                importer.CodeGenerationOptions = CodeGenerationOptions.GenerateProperties | CodeGenerationOptions.GenerateNewAsync;
                importer.AddServiceDescription(description, null, null); // Add a WSDL document.// 4. Compile the client proxy classes using CodeDom.
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
                CompilerResults result = provider.CompileAssemblyFromDom(parameter, unit);// 5. Using Reflection to call WebService.
                if (!result.Errors.HasErrors)
                {
                    Assembly asm = result.CompiledAssembly;
                    Type t = asm.GetType("ComponentExecutor", true, true); // If in front of adding a namespace for the proxy class, here need to be added to the front of the namespace type.
                    object o = Activator.CreateInstance(t);
                    MethodInfo method = t.GetMethod("Execute");
                    string sInputPayload;
                    sInputPayload = this.txtXML.Text.Replace("\r\n", "").Replace("\n", ""); ;
                    object strResu = method.Invoke(o, new object[] { sSSCVoucher, null, "Journal", "Import", null, sInputPayload });
                    //MessageBox.Show(strResu.ToString(), "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.textBox1.Text = GetErrorLines(strResu.ToString());
                    //if (Ribbon2.xpf.Visible == false)//(!Ribbon2.outputPane.chkShowForm.Checked && sender != null) || sender == null || 
                    //{
                    //    PostErrorFrm pef = new PostErrorFrm(this.textBox1.Text);
                    //    pef.ShowDialog();
                    //    //MessageBox.Show(this.textBox1.Text.Replace("This line has been rejected due to errors in other lines or posting options", ""), "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //}
                    if (!string.IsNullOrEmpty(this.textBox1.Text.Trim()))
                    {
                        SessionInfo.UserInfo.GlobalError += "Process:Journal Post(" + SessionInfo.UserInfo.CurrentRef + ") - Fail! " + this.textBox1.Text;
                        return false;
                    }
                    else
                    {
                        try
                        {//Save journal number
                            XmlDocument xdoc = new XmlDocument();
                            xdoc.LoadXml(strResu.ToString());//XmlNode node = xdoc.DocumentElement;//XmlNode node = xdoc.SelectSingleNode("Nodes");
                            journalNumber = xdoc.GetElementsByTagName("JournalNumber").Item(1).InnerText;
                            if (!string.IsNullOrEmpty(journalNumber))
                            {
                                SessionInfo.UserInfo.SunJournalNumber = journalNumber;
                                PopulateCellWithJnNumber();
                                if (Ribbon2.xpf.Visible == true)
                                    MessageBox.Show("journalNumber:" + journalNumber, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                else
                                    SessionInfo.UserInfo.GlobalError += "Process:Journal Post(" + SessionInfo.UserInfo.CurrentRef + ") - Success! journalNumber:" + journalNumber;
                            }//present journal number
                        }
                        catch (Exception ex)
                        { throw ex; }
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
                if (Ribbon2.xpf.Visible == true)
                    MessageBox.Show(ex.ToString(), "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    SessionInfo.UserInfo.GlobalError += "Process:Journal Post(" + SessionInfo.UserInfo.CurrentRef + ") - Fail! " + ex.Message;
                LogHelper.WriteLog(typeof(Ribbon2), ex.Message + "Do Post error");
                return false;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        private void PopulateCellWithJnNumber()
        {
            if (!string.IsNullOrEmpty(SessionInfo.UserInfo.PopulateCellWithJnNumber))
                Globals.ThisAddIn.Application.Rows.get_Range(SessionInfo.UserInfo.PopulateCellWithJnNumber).Value = SessionInfo.UserInfo.SunJournalNumber;
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
            try
            {
                XmlOperator Xo = new XmlOperator(str);
                XmlNodeList userCollection = Xo.XmlDoc.SelectNodes("//ErrorMessages/Message");
                foreach (XmlNode xn in userCollection)
                {
                    try
                    {
                        reference = "Line " + xn.SelectSingleNode("PayloadRef").InnerText;
                    }
                    catch
                    {
                        reference = "";
                    }
                    try
                    {
                        usertext = xn.SelectSingleNode("UserText").InnerText;
                    }
                    catch
                    {
                        usertext = "";
                    }
                    if (usertext != "This line has been rejected due to errors in other lines or posting options")
                        error += reference + " : " + usertext + "\r\n";
                }
                return error;
            }
            catch
            {
                XmlOperator Xo = new XmlOperator(str);
                XmlNodeList userCollection = Xo.XmlDoc.GetElementsByTagName("Line");
                foreach (XmlNode xn in userCollection)
                {
                    if (xn.Attributes["status"].Value == "fail")
                    {
                        //XmlOperator Xo = new XmlOperator(xn.InnerXml);
                        XmlNode xn2 = xn.SelectSingleNode("//Messages/Message/UserText");
                        try
                        {
                            reference = "Line " + xn.Attributes["Reference"].Value;
                        }
                        catch
                        {
                            reference = "";
                        }
                        try
                        {
                            usertext = xn2.InnerText;
                        }
                        catch
                        {
                            usertext = "";
                        }
                        if (usertext != "This line has been rejected due to errors in other lines or posting options")
                            error += reference + " : " + usertext + "\r\n";
                    }
                }
                return error;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        public void SaveHistory()
        {
            if ((SessionInfo.UserInfo.UseSequenceNumbering == "1") || SessionInfo.UserInfo.UseCriteria)
            {
                int? i = 0;
                string invName = string.Empty;
                string prefix = string.Empty;
                try
                {
                    ft.GetInvoiceInfo(ref prefix, ref invName, ref i);
                }
                catch
                {
                }
                if (!string.IsNullOrEmpty(invName))
                {
                    PopulateCell(invName);
                }
                string tmp = Guid.NewGuid().ToString();
                //Globals.ThisAddIn.Application.ActiveWorkbook.SaveCopyAs(AppDomain.CurrentDomain.BaseDirectory + "SaveHistoryTmpFile" + Path.GetExtension(SessionInfo.UserInfo.FilePath));
                SqlConnection conn = null;
                try
                {
                    conn = new
                        SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                    conn.Open();
                    string IDvalue = string.Empty;
                    if (SessionInfo.UserInfo.Dictionary.dict.Count != 0 && SessionInfo.UserInfo.Dictionary.dict.ContainsKey(SessionInfo.UserInfo.CachePath))
                    {
                        IDvalue = SessionInfo.UserInfo.Dictionary.dict[SessionInfo.UserInfo.CachePath];
                        string[] sArray = Regex.Split(IDvalue, ",");
                        IDvalue = sArray[0];
                    }
                    if (!ft.IsGUID(Path.GetFileNameWithoutExtension(SessionInfo.UserInfo.CachePath)) || string.IsNullOrEmpty(IDvalue))
                    {
                        if (SessionInfo.UserInfo.OpentransuponSave)
                        {
                            if (!ft.IsGUID(Path.GetFileNameWithoutExtension(SessionInfo.UserInfo.CachePath)))
                            {
                                Globals.ThisAddIn.Application.ActiveWorkbook.SaveAs(AppDomain.CurrentDomain.BaseDirectory.EndsWith("\\") ? AppDomain.CurrentDomain.BaseDirectory + tmp + Path.GetExtension(SessionInfo.UserInfo.FilePath) : AppDomain.CurrentDomain.BaseDirectory + "\\" + tmp + Path.GetExtension(SessionInfo.UserInfo.FilePath));
                                SessionInfo.UserInfo.CachePath = Globals.ThisAddIn.Application.ActiveWorkbook.FullName;
                                if (!SessionInfo.UserInfo.Dictionary.dict.ContainsKey(SessionInfo.UserInfo.CachePath))
                                    SessionInfo.UserInfo.Dictionary.dict.Add(SessionInfo.UserInfo.CachePath, SessionInfo.UserInfo.File_ftid + "," + i);
                            }
                        }
                        else
                            Globals.ThisAddIn.Application.ActiveWorkbook.SaveCopyAs(AppDomain.CurrentDomain.BaseDirectory.EndsWith("\\") ? AppDomain.CurrentDomain.BaseDirectory + tmp + Path.GetExtension(SessionInfo.UserInfo.FilePath) : AppDomain.CurrentDomain.BaseDirectory + "\\" + tmp + Path.GetExtension(SessionInfo.UserInfo.FilePath));
                        SqlCommand cmd = new SqlCommand("rsTemplateTransactions_Ins", conn);
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add(new SqlParameter("@TemplateName", SessionInfo.UserInfo.FileName));
                        cmd.Parameters.Add(new SqlParameter("@Criteria1", SessionInfo.UserInfo.Criteria1));
                        cmd.Parameters.Add(new SqlParameter("@Criteria2", SessionInfo.UserInfo.Criteria2));
                        cmd.Parameters.Add(new SqlParameter("@Criteria3", SessionInfo.UserInfo.Criteria3));
                        cmd.Parameters.Add(new SqlParameter("@Criteria4", SessionInfo.UserInfo.Criteria4));
                        cmd.Parameters.Add(new SqlParameter("@Criteria5", SessionInfo.UserInfo.Criteria5));
                        cmd.Parameters.Add(new SqlParameter("@Value1", SessionInfo.UserInfo.CellReference1));
                        cmd.Parameters.Add(new SqlParameter("@Value2", SessionInfo.UserInfo.CellReference2));
                        cmd.Parameters.Add(new SqlParameter("@Value3", SessionInfo.UserInfo.CellReference3));
                        cmd.Parameters.Add(new SqlParameter("@Value4", SessionInfo.UserInfo.CellReference4));
                        cmd.Parameters.Add(new SqlParameter("@Value5", SessionInfo.UserInfo.CellReference5));
                        //if (SessionInfo.UserInfo.CachePath != SessionInfo.UserInfo.FilePath)
                        //    cmd.Parameters.Add(new SqlParameter("@Data", ft.GetData(SessionInfo.UserInfo.CachePath)));
                        //else
                        cmd.Parameters.Add(new SqlParameter("@Data", ft.GetData(AppDomain.CurrentDomain.BaseDirectory.EndsWith("\\") ? AppDomain.CurrentDomain.BaseDirectory + tmp + Path.GetExtension(SessionInfo.UserInfo.FilePath) : AppDomain.CurrentDomain.BaseDirectory + "\\" + tmp + Path.GetExtension(SessionInfo.UserInfo.FilePath))));
                        cmd.Parameters.Add(new SqlParameter("@DataType", Path.GetExtension(SessionInfo.UserInfo.FilePath)));
                        cmd.Parameters.Add(new SqlParameter("@PDFData", ft.GetData(SessionInfo.UserInfo.Containerpath)));
                        cmd.Parameters.Add(new SqlParameter("@XMLData", this.txtXML.Text));
                        //cmd.Parameters.Add(new SqlParameter("@OwnUserID", SessionInfo.UserInfo.ID));
                        cmd.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                        cmd.Parameters.Add(new SqlParameter("@maxNum", i));
                        cmd.Parameters.Add(new SqlParameter("@TransactionName", invName));
                        cmd.Parameters.Add(new SqlParameter("@Prefix", prefix));
                        cmd.Parameters.Add(new SqlParameter("@SunJournalNumber", SessionInfo.UserInfo.SunJournalNumber));
                        cmd.ExecuteNonQuery();
                    }
                    else
                    {
                        SqlCommand cmd = new SqlCommand("rsTemplateTransactions_Del", conn);
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                        cmd.Parameters.Add(new SqlParameter("@TransactionName", invName));
                        cmd.ExecuteNonQuery();
                        SqlCommand cmd2 = new SqlCommand("rsTemplateTransactions_Ins", conn);
                        cmd2.CommandType = CommandType.StoredProcedure;
                        cmd2.Parameters.Add(new SqlParameter("@TemplateName", SessionInfo.UserInfo.FileName));
                        cmd2.Parameters.Add(new SqlParameter("@Criteria1", SessionInfo.UserInfo.Criteria1));
                        cmd2.Parameters.Add(new SqlParameter("@Criteria2", SessionInfo.UserInfo.Criteria2));
                        cmd2.Parameters.Add(new SqlParameter("@Criteria3", SessionInfo.UserInfo.Criteria3));
                        cmd2.Parameters.Add(new SqlParameter("@Criteria4", SessionInfo.UserInfo.Criteria4));
                        cmd2.Parameters.Add(new SqlParameter("@Criteria5", SessionInfo.UserInfo.Criteria5));
                        cmd2.Parameters.Add(new SqlParameter("@Value1", SessionInfo.UserInfo.CellReference1));
                        cmd2.Parameters.Add(new SqlParameter("@Value2", SessionInfo.UserInfo.CellReference2));
                        cmd2.Parameters.Add(new SqlParameter("@Value3", SessionInfo.UserInfo.CellReference3));
                        cmd2.Parameters.Add(new SqlParameter("@Value4", SessionInfo.UserInfo.CellReference4));
                        cmd2.Parameters.Add(new SqlParameter("@Value5", SessionInfo.UserInfo.CellReference5));
                        Globals.ThisAddIn.Application.ActiveWorkbook.Save();
                        cmd2.Parameters.Add(new SqlParameter("@Data", ft.GetData(SessionInfo.UserInfo.CachePath)));
                        cmd2.Parameters.Add(new SqlParameter("@DataType", Path.GetExtension(SessionInfo.UserInfo.FilePath)));
                        cmd2.Parameters.Add(new SqlParameter("@PDFData", ft.GetData(SessionInfo.UserInfo.Containerpath)));
                        cmd2.Parameters.Add(new SqlParameter("@XMLData", this.txtXML.Text));
                        //cmd.Parameters.Add(new SqlParameter("@OwnUserID", SessionInfo.UserInfo.ID));
                        cmd2.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                        cmd2.Parameters.Add(new SqlParameter("@maxNum", i));
                        cmd2.Parameters.Add(new SqlParameter("@TransactionName", invName));
                        cmd2.Parameters.Add(new SqlParameter("@Prefix", prefix));
                        cmd2.Parameters.Add(new SqlParameter("@SunJournalNumber", SessionInfo.UserInfo.SunJournalNumber));
                        cmd2.ExecuteNonQuery();
                    }

                    if (!string.IsNullOrEmpty(invName))
                    {
                        if (Ribbon2.xpf.Visible == true)
                            MessageBox.Show("Your new transaction name is " + invName, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        else
                            SessionInfo.UserInfo.GlobalError += "Process:Save(" + SessionInfo.UserInfo.CurrentSaveRef + ") - Success! Your new transaction name is " + invName;
                    }
                    string str = GetCriteriaStr();
                    if (!string.IsNullOrEmpty(str))
                    {
                        if (Ribbon2.xpf.Visible == true)
                            MessageBox.Show("Your new transaction criterias are: \r\n" + str, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        else
                            SessionInfo.UserInfo.GlobalError += ".Your new transaction criterias are: \r\n" + str;

                        Ribbon2.outputPane.BindSaveOptions();
                    }
                    SessionInfo.UserInfo.InvNumber = i;
                    //if ((!ft.IsGUID(Path.GetFileNameWithoutExtension(SessionInfo.UserInfo.CachePath)) || string.IsNullOrEmpty(IDvalue)) && SessionInfo.UserInfo.OpentransuponSave)
                    //{
                    //    OpenTransaction(invName);
                    //}
                }
                catch (Exception ex) { throw new Exception(ex.ToString()); }
                finally
                {
                    if (conn != null)
                    {
                        conn.Close();
                    }
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Invnumber"></param>
        //private void OpenTransaction(string Invnumber)
        //{
        //    var xlapp = Globals.ThisAddIn.Application;
        //    Globals.ThisAddIn.Application.DisplayAlerts = false;
        //    xlapp.ActiveWorkbook.Close();
        //    Globals.ThisAddIn.Application.DisplayAlerts = true;
        //    var data = ft.ProcessData(Invnumber);
        //    var fileType = ft.ProcessFileType(Invnumber);
        //    string tmp = Guid.NewGuid().ToString();
        //    var file = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType, FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
        //    var bw = new BinaryWriter(file);
        //    bw.Write(data);
        //    bw.Close();
        //    file.Close();
        //    if (fileType == "pdf")
        //    {
        //        Process.Start(AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType);
        //        return;
        //    }
        //    else
        //    {
        //        xlapp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType);//xlapp.Run("'" + AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType + "'!RSystems.runProc");
        //        return;
        //    }
        //}
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
                        List<Line> newlist = new List<Line>();
                        for (int i = 0; i < dgv.Rows.Count; i++)
                        {
                            Line re = new Line();
                            re.DetailLad = new DetailLad();
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
                                    re.DetailLad.AccountCode = invName;
                                    dgv.Rows[i].Cells["AccountCode"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("AccountCode,"))
                                {
                                    re.DetailLad.AccountCode = null;
                                    dgv.Rows[i].Cells["AccountCode"].Value = null;
                                }
                                else
                                    re.DetailLad.AccountCode = dgv.Rows[i].Cells["AccountCode"] == null ? "" : dgv.Rows[i].Cells["AccountCode"].Value.ToString();
                            }
                            if ((dgv.Columns["AccountingPeriod"] != null) && (dgv.Columns["AccountingPeriod"].Visible == true))
                                re.DetailLad.AccountingPeriod = re.AccountingPeriod;
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
                                    re.DetailLad.GeneralDescription1 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription1"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc1,"))
                                {
                                    re.DetailLad.GeneralDescription1 = null;
                                    dgv.Rows[i].Cells["GeneralDescription1"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription1 = dgv.Rows[i].Cells["GeneralDescription1"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription1"].Value.ToString();
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
                                    re.DetailLad.GeneralDescription2 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription2"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc2,"))
                                {
                                    re.DetailLad.GeneralDescription2 = null;
                                    dgv.Rows[i].Cells["GeneralDescription2"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription2 = dgv.Rows[i].Cells["GeneralDescription2"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription2"].Value.ToString();
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
                                    re.DetailLad.GeneralDescription3 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription3"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc3,"))
                                {
                                    re.DetailLad.GeneralDescription3 = null;
                                    dgv.Rows[i].Cells["GeneralDescription3"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription3 = dgv.Rows[i].Cells["GeneralDescription3"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription3"].Value.ToString();
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
                                    re.DetailLad.GeneralDescription4 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription4"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc4,"))
                                {
                                    re.DetailLad.GeneralDescription4 = null;
                                    dgv.Rows[i].Cells["GeneralDescription4"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription4 = dgv.Rows[i].Cells["GeneralDescription4"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription4"].Value.ToString();
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
                                    re.DetailLad.GeneralDescription5 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription5"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc5,"))
                                {
                                    re.DetailLad.GeneralDescription5 = null;
                                    dgv.Rows[i].Cells["GeneralDescription5"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription5 = dgv.Rows[i].Cells["GeneralDescription5"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription5"].Value.ToString();
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
                                    re.DetailLad.GeneralDescription6 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription6"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc6,"))
                                {
                                    re.DetailLad.GeneralDescription6 = null;
                                    dgv.Rows[i].Cells["GeneralDescription6"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription6 = dgv.Rows[i].Cells["GeneralDescription6"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription6"].Value.ToString();
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
                                    re.DetailLad.GeneralDescription7 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription7"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc7,"))
                                {
                                    re.DetailLad.GeneralDescription7 = null;
                                    dgv.Rows[i].Cells["GeneralDescription7"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription7 = dgv.Rows[i].Cells["GeneralDescription7"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription7"].Value.ToString();
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
                                    re.DetailLad.GeneralDescription8 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription8"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc8,"))
                                {
                                    re.DetailLad.GeneralDescription8 = null;
                                    dgv.Rows[i].Cells["GeneralDescription8"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription8 = dgv.Rows[i].Cells["GeneralDescription8"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription8"].Value.ToString();
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
                                    re.DetailLad.GeneralDescription9 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription9"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc9,"))
                                {
                                    re.DetailLad.GeneralDescription9 = null;
                                    dgv.Rows[i].Cells["GeneralDescription9"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription9 = dgv.Rows[i].Cells["GeneralDescription9"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription9"].Value.ToString();
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
                                    re.DetailLad.GeneralDescription10 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription10"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc10,"))
                                {
                                    re.DetailLad.GeneralDescription10 = null;
                                    dgv.Rows[i].Cells["GeneralDescription10"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription10 = dgv.Rows[i].Cells["GeneralDescription10"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription10"].Value.ToString();
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
                                    re.DetailLad.GeneralDescription11 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription11"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc11,"))
                                {
                                    re.DetailLad.GeneralDescription11 = null;
                                    dgv.Rows[i].Cells["GeneralDescription11"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription11 = dgv.Rows[i].Cells["GeneralDescription11"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription11"].Value.ToString();
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
                                    re.DetailLad.GeneralDescription12 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription12"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc12,"))
                                {
                                    re.DetailLad.GeneralDescription12 = null;
                                    dgv.Rows[i].Cells["GeneralDescription12"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription12 = dgv.Rows[i].Cells["GeneralDescription12"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription12"].Value.ToString();
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
                                    re.DetailLad.GeneralDescription13 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription13"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc13,"))
                                {
                                    re.DetailLad.GeneralDescription13 = null;
                                    dgv.Rows[i].Cells["GeneralDescription13"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription13 = dgv.Rows[i].Cells["GeneralDescription13"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription13"].Value.ToString();
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
                                    re.DetailLad.GeneralDescription14 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription14"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc14,"))
                                {
                                    re.DetailLad.GeneralDescription14 = null;
                                    dgv.Rows[i].Cells["GeneralDescription14"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription14 = dgv.Rows[i].Cells["GeneralDescription14"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription14"].Value.ToString();
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
                                    re.DetailLad.GeneralDescription15 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription15"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc15,"))
                                {
                                    re.DetailLad.GeneralDescription15 = null;
                                    dgv.Rows[i].Cells["GeneralDescription15"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription15 = dgv.Rows[i].Cells["GeneralDescription15"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription15"].Value.ToString();
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
                                    re.DetailLad.GeneralDescription16 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription16"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc16,"))
                                {
                                    re.DetailLad.GeneralDescription16 = null;
                                    dgv.Rows[i].Cells["GeneralDescription16"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription16 = dgv.Rows[i].Cells["GeneralDescription16"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription16"].Value.ToString();
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
                                    re.DetailLad.GeneralDescription17 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription17"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc17,"))
                                {
                                    re.DetailLad.GeneralDescription17 = null;
                                    dgv.Rows[i].Cells["GeneralDescription17"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription17 = dgv.Rows[i].Cells["GeneralDescription17"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription17"].Value.ToString();
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
                                    re.DetailLad.GeneralDescription18 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription18"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc18,"))
                                {
                                    re.DetailLad.GeneralDescription18 = null;
                                    dgv.Rows[i].Cells["GeneralDescription18"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription18 = dgv.Rows[i].Cells["GeneralDescription18"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription18"].Value.ToString();
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
                                    re.DetailLad.GeneralDescription19 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription19"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc19,"))
                                {
                                    re.DetailLad.GeneralDescription19 = null;
                                    dgv.Rows[i].Cells["GeneralDescription19"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription19 = dgv.Rows[i].Cells["GeneralDescription19"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription19"].Value.ToString();
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
                                    re.DetailLad.GeneralDescription20 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription20"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc20,"))
                                {
                                    re.DetailLad.GeneralDescription20 = null;
                                    dgv.Rows[i].Cells["GeneralDescription20"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription20 = dgv.Rows[i].Cells["GeneralDescription20"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription20"].Value.ToString();
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
                                    re.DetailLad.GeneralDescription21 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription21"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc21,"))
                                {
                                    re.DetailLad.GeneralDescription21 = null;
                                    dgv.Rows[i].Cells["GeneralDescription21"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription21 = dgv.Rows[i].Cells["GeneralDescription21"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription21"].Value.ToString();
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
                                    re.DetailLad.GeneralDescription22 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription22"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc22,"))
                                {
                                    re.DetailLad.GeneralDescription22 = null;
                                    dgv.Rows[i].Cells["GeneralDescription22"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription22 = dgv.Rows[i].Cells["GeneralDescription22"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription22"].Value.ToString();
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
                                    re.DetailLad.GeneralDescription23 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription23"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc23,"))
                                {
                                    re.DetailLad.GeneralDescription23 = null;
                                    dgv.Rows[i].Cells["GeneralDescription23"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription23 = dgv.Rows[i].Cells["GeneralDescription23"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription23"].Value.ToString();
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
                                    re.DetailLad.GeneralDescription24 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription24"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc24,"))
                                {
                                    re.DetailLad.GeneralDescription24 = null;
                                    dgv.Rows[i].Cells["GeneralDescription24"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription24 = dgv.Rows[i].Cells["GeneralDescription24"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription24"].Value.ToString();
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
                                    re.DetailLad.GeneralDescription25 = invName;
                                    dgv.Rows[i].Cells["GeneralDescription25"].Value = invName;
                                }
                                else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains("GenDesc25,"))
                                {
                                    re.DetailLad.GeneralDescription25 = null;
                                    dgv.Rows[i].Cells["GeneralDescription25"].Value = null;
                                }
                                else
                                    re.DetailLad.GeneralDescription25 = dgv.Rows[i].Cells["GeneralDescription25"] == null ? "" : dgv.Rows[i].Cells["GeneralDescription25"].Value.ToString();
                            }

                            if ((dgv.Columns["TransactionDate"] != null) && (dgv.Columns["TransactionDate"].Visible == true))
                            {
                                re.DetailLad.TransactionDate = re.TransactionDate;//"ddMMyyyy"
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


                            if (double.Parse(re.TransactionAmount) > 0)
                                re.DebitCredit = "D";
                            else if (double.Parse(re.TransactionAmount) < 0)
                                re.DebitCredit = "C";
                            else if (double.Parse(re.TransactionAmount) == 0)
                            {
                                if (double.Parse(re.BaseAmount) > 0)
                                    re.DebitCredit = "D";
                                else if (double.Parse(re.BaseAmount) < 0)
                                    re.DebitCredit = "C";
                                else if (double.Parse(re.BaseAmount) == 0)
                                {
                                    if (double.Parse(re.Base2ReportingAmount) > 0)
                                        re.DebitCredit = "D";
                                    else if (double.Parse(re.Base2ReportingAmount) < 0)
                                        re.DebitCredit = "C";
                                }
                            }
                            if (!string.IsNullOrEmpty(re.Base2ReportingAmount))
                                re.Base2ReportingAmount = re.Base2ReportingAmount.Replace("-", "");
                            if (!string.IsNullOrEmpty(re.TransactionAmount))
                                re.TransactionAmount = re.TransactionAmount.Replace("-", "");
                            if (!string.IsNullOrEmpty(re.Value4Amount))
                                re.Value4Amount = re.Value4Amount.Replace("-", "");
                            if (!string.IsNullOrEmpty(re.BaseAmount))
                                re.BaseAmount = re.BaseAmount.Replace("-", "");

                            //if (re.TransactionAmount != "0.000")
                            newlist.Add(re);
                        }
                        string script = ft.GetXMLScript(newlist);
                        this.txtXML.Text = script;
                    }
                }
                catch (Exception ex)
                {
                    this.txtXML.Text = ex.Message;
                }
            }
            else if (this.tabControl1.SelectedIndex == 2)
            {
                try
                {
                    this.tabPage4.Controls.Clear();
                    if (!string.IsNullOrEmpty(SessionInfo.UserInfo.BalanceBy))
                    {
                        DataGridView dgv = new DataGridView();
                        dgv.Columns.Add("", "");
                        dgv.Columns.Add("BalanceBy", SessionInfo.UserInfo.BalanceBy);
                        dgv.Columns["BalanceBy"].DataPropertyName = "BalanceBy";
                        dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
                        dgv.AutoGenerateColumns = false;
                        dgv.ColumnHeadersHeight = 40;
                        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                        dgv.Dock = DockStyle.Top;
                        dgv.Visible = true;
                        DataTable tb1 = ft.ToDataTable(Ribbon2.outputPane.finallist);
                        string TransAmount = GetAmount(tb1, "TransactionAmount");
                        string BaseAmount = GetAmount(tb1, "BaseAmount");
                        string ndBaseAmount = GetAmount(tb1, "Base2ReportingAmount");
                        string thBaseAmount = GetAmount(tb1, "Value4Amount");
                        dgv.Rows.Add("Transaction Amount", TransAmount);
                        dgv.Rows.Add("Base Amount", BaseAmount);
                        dgv.Rows.Add("2nd Base Amount", ndBaseAmount);
                        dgv.Rows.Add("4th Amount", thBaseAmount);
                        this.tabPage4.Controls.Add(dgv);
                    }
                    else
                    {
                        Label tb = new Label();
                        tb.AutoSize = true;
                        tb.Text = "No BalanceBy Data!";
                        this.tabPage4.Controls.Add(tb);
                    }
                }
                catch
                { }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="tb1"></param>
        /// <param name="groupbyField"></param>
        /// <returns></returns>
        private string GetAmount(DataTable tb1, string groupbyField)
        {
            try
            {
                var query2 = from t in tb1.AsEnumerable()
                             group t by new { t1 = t.Field<string>(SessionInfo.UserInfo.BalanceBy) } into m
                             select new
                             {
                                 ReturnValue = m.Sum(k => Decimal.Parse(string.IsNullOrEmpty(k.Field<string>(groupbyField)) ? "0" : k.Field<string>(groupbyField))),
                             };
                DataTable newdt3 = ft.ToDataTable(query2.ToList());
                if (newdt3.Rows.Count > 0)
                {
                    return newdt3.Rows[0]["ReturnValue"].ToString();
                }
                else
                    return "0.000";
            }
            catch
            {
                return "0.000";
            }
        }
    }
}