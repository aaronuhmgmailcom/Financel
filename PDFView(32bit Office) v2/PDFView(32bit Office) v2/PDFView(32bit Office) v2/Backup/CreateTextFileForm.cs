/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<CreateTextFileForm>   
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
    public partial class CreateTextFileForm : Form
    {
        DataGridView dgv = null;
        string journalNumber = string.Empty;
        string sequenceNumbering = string.Empty;
        public static string removedColumns = string.Empty;
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
        public CreateTextFileForm()
        {
            InitializeComponent();
            this.Text = "Post XML Data (" + SessionInfo.UserInfo.CurrentRef + ") - RSystems FinanceTools v2";
        }
        /// <summary>
        /// 
        /// </summary>
        public void bddata()
        {
            try
            {
                dgv = new DataGridView();
                DataTable dt = ft.GetXMLorTextFileFieldsByComName(SessionInfo.UserInfo.ComName, SessionInfo.UserInfo.MethodName);
                dgv.Columns.Add("ReferenceNumber", "Ref");
                dgv.Columns["ReferenceNumber"].DataPropertyName = "ReferenceNumber";
                dgv.Columns["ReferenceNumber"].Visible = false;
                dgv.Columns.Add("LineIndicator", "Line Indicator");
                dgv.Columns["LineIndicator"].DataPropertyName = "LineIndicator";
                dgv.Columns["LineIndicator"].Visible = false;
                dgv.Columns.Add("StartinginCell", "StartinginCell");
                dgv.Columns["StartinginCell"].DataPropertyName = "StartinginCell";
                dgv.Columns["StartinginCell"].Visible = false;
                dgv.Columns.Add("IncludeHeaderRow", "IncludeHeaderRow");
                dgv.Columns["IncludeHeaderRow"].DataPropertyName = "IncludeHeaderRow";
                dgv.Columns["IncludeHeaderRow"].Visible = false;
                dgv.Columns.Add("SavePath", "SavePath");
                dgv.Columns["SavePath"].DataPropertyName = "SavePath";
                dgv.Columns["SavePath"].Visible = false;
                dgv.Columns.Add("SaveName", "SaveName");
                dgv.Columns["SaveName"].DataPropertyName = "SaveName";
                dgv.Columns["SaveName"].Visible = false;
                dgv.Columns.Add("SunComponent", "SunComponent");
                dgv.Columns["SunComponent"].DataPropertyName = "SunComponent";
                dgv.Columns["SunComponent"].Visible = false;
                dgv.Columns.Add("SunMethod", "SunMethod");
                dgv.Columns["SunMethod"].DataPropertyName = "SunMethod";
                dgv.Columns["SunMethod"].Visible = false;
                dgv.Columns.Add("ProcessName", "ProcessName");
                dgv.Columns["ProcessName"].DataPropertyName = "ProcessName";
                dgv.Columns["ProcessName"].Visible = false;
                dgv.Columns.Add("HeaderTextes", "HeaderTextes");
                dgv.Columns["HeaderTextes"].DataPropertyName = "HeaderTextes";
                dgv.Columns["HeaderTextes"].Visible = false;
                dgv.Columns.Add("HeaderValue", "HeaderValue");
                dgv.Columns["HeaderValue"].DataPropertyName = "HeaderValue";
                dgv.Columns["HeaderValue"].Visible = false;
                int count = 1;
                removedColumns = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (bool.Parse(dt.Rows[i]["Visible"].ToString()) == true)
                    {
                        if (dgv.Columns.Contains(dt.Rows[i]["Field"].ToString()))
                        {
                            dgv.Columns.Add(dt.Rows[i]["Field"].ToString() + count, dt.Rows[i]["FriendlyName"].ToString());
                            dgv.Columns[dt.Rows[i]["Field"].ToString() + count].DataPropertyName = "Column" + count;
                            dgv.Columns[dt.Rows[i]["Field"].ToString() + count].Tag = dt.Rows[i]["Field"].ToString();//dt.Rows[i]["DefaultValue"].ToString() + ",,," + dt.Rows[i]["Mandatory"].ToString() + ",,," + dt.Rows[i]["Separator"].ToString() + ",,," + dt.Rows[i]["TextLength"].ToString() + ",,," + dt.Rows[i]["Prefix"].ToString() + ",,," + dt.Rows[i]["Suffix"].ToString() + ",,," + dt.Rows[i]["RemoveCharacters"].ToString() + ",,," + dt.Rows[i]["Parent"].ToString() + ",,,";
                        }
                        else
                        {
                            dgv.Columns.Add(dt.Rows[i]["Field"].ToString(), dt.Rows[i]["FriendlyName"].ToString());
                            dgv.Columns[dt.Rows[i]["Field"].ToString()].DataPropertyName = "Column" + count;
                            dgv.Columns[dt.Rows[i]["Field"].ToString()].Tag = dt.Rows[i]["Field"].ToString();//dt.Rows[i]["DefaultValue"].ToString() + ",,," + dt.Rows[i]["Mandatory"].ToString() + ",,," + dt.Rows[i]["Separator"].ToString() + ",,," + dt.Rows[i]["TextLength"].ToString() + ",,," + dt.Rows[i]["Prefix"].ToString() + ",,," + dt.Rows[i]["Suffix"].ToString() + ",,," + dt.Rows[i]["RemoveCharacters"].ToString() + ",,," + dt.Rows[i]["Parent"].ToString() + ",,,";
                        }
                        count++;
                    }
                }
                dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
                dgv.AutoGenerateColumns = false;
                dgv.ColumnHeadersHeight = 40;
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dgv.Dock = DockStyle.Fill;
                dgv.Visible = true;
                dgv.ColumnHeaderMouseClick += new DataGridViewCellMouseEventHandler(DataGridView1_CellMouseDown);
                dgv.RowPostPaint += new DataGridViewRowPostPaintEventHandler(dgv_RowPostPaint);
                BindData();
                DataTable tb1 = ft.ToDataTable((List<RowCreateTextFile>)this.dgv.DataSource);
                for (int i = 11; i < tb1.Columns.Count; i++)
                {
                    string str = string.Empty;
                    for (int j = 0; j < tb1.Rows.Count; j++)
                    {
                        str += tb1.Rows[j][i].ToString();
                        if (tb1.Rows[j][i].ToString().ToUpper() == "[SEQUENCE]")
                            sequenceNumbering = tb1.Columns[i].ColumnName + "," + sequenceNumbering;
                    }
                    if (string.IsNullOrEmpty(str.Replace("0", "").Replace(".", "")) && dgv.Columns.Count > i)
                    {
                        dgv.Columns[i].Visible = false;//Remove the cloumns doesn't contain data
                        removedColumns += dgv.Columns[i].Tag + dgv.Columns[i].DataPropertyName + ",";
                    }
                }
                for (int i = 0; i < dgv.Columns.Count; i++)
                {
                    string type = ft.GetSectionFromDB(SessionInfo.UserInfo.ComName, SessionInfo.UserInfo.MethodName, dgv.Columns[i].Name, dgv.Columns[i].HeaderText);
                    if (type == "Header")
                        dgv.Columns[i].DefaultCellStyle.BackColor = Color.Ivory;
                    else if (type == "Footer")
                        dgv.Columns[i].DefaultCellStyle.BackColor = Color.Beige;
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
                this.dgv.DataSource = Ribbon2.outputPane.finallistCTF;
            }
            catch (Exception ex)
            {
                CreateTextFileForm.richTextBox1.Text += ex.Message;
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
            { }
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
                    sSSCVoucher = oSecMan.Voucher;//label1.Content = sSSCVoucher;RSI, RicoSimp2
                }
                else
                {
                    this.textBox1.Text = "SunSystems Server is not exist or Password for user is incorrect.";
                    if (Ribbon2.ctff.Visible == true)
                        MessageBox.Show(this.textBox1.Text, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else
                        SessionInfo.UserInfo.GlobalError += "Process:" + SessionInfo.UserInfo.ComName + " " + SessionInfo.UserInfo.MethodName + "(" + SessionInfo.UserInfo.CurrentRef + ") - Fail! " + this.textBox1.Text;

                    return false;
                }
            }
            catch (Exception ex)
            {
                this.textBox1.Text = "An error occurred in validation " + (oSecMan == null ? "" : oSecMan.ErrorMessage) + ex;
                if (Ribbon2.ctff.Visible == true)
                    MessageBox.Show(this.textBox1.Text, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    SessionInfo.UserInfo.GlobalError += "Process:" + SessionInfo.UserInfo.ComName + " " + SessionInfo.UserInfo.MethodName + "(" + SessionInfo.UserInfo.CurrentRef + ") - Fail! " + this.textBox1.Text;

                LogHelper.WriteLog(typeof(CreateTextFileForm), this.textBox1.Text);
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
                    Type t = asm.GetType("ComponentExecutor", true, true);//If in front of adding a namespace for the proxy class, here need to be added to the front of the namespace type.
                    object o = Activator.CreateInstance(t);
                    MethodInfo method = t.GetMethod("Execute");
                    string sInputPayload;
                    sInputPayload = this.txtXML.Text.Replace("\r\n", "").Replace("\n", ""); ;
                    string[] sArray1 = sInputPayload.Split(new char[3] { '*', '*', '*' });
                    object strResu = null;
                    if (sArray1.Length > 0)
                    {
                        foreach (string ss in sArray1)
                            if (!string.IsNullOrEmpty(ss))
                            {
                                strResu = method.Invoke(o, new object[] { sSSCVoucher, null, SessionInfo.UserInfo.ComName, SessionInfo.UserInfo.MethodName, null, ss });
                                this.textBox1.Text += GetErrorLines(strResu.ToString()) + "\r\n***\r\n";
                            }
                    }
                    else
                    {
                        strResu = method.Invoke(o, new object[] { sSSCVoucher, null, SessionInfo.UserInfo.ComName, SessionInfo.UserInfo.MethodName, null, sInputPayload });
                        this.textBox1.Text = GetErrorLines(strResu.ToString());
                    }
                    //if (Ribbon2.ctff.Visible == false)//(!Ribbon2.outputPane.chkShowForm.Checked && sender != null) || sender == null || 
                    //{
                    //    PostErrorFrm pef = new PostErrorFrm(this.textBox1.Text.Replace("This line has been rejected due to errors in other lines or posting options", ""));
                    //    pef.ShowDialog();
                    //}
                    if (!string.IsNullOrEmpty(this.textBox1.Text.Trim().Replace("***", "")))
                    {
                        SessionInfo.UserInfo.GlobalError += "Process:" + SessionInfo.UserInfo.ComName + " " + SessionInfo.UserInfo.MethodName + "(" + SessionInfo.UserInfo.CurrentRef + ") - Fail! " + this.textBox1.Text;
                        return false;
                    }
                    else
                    {
                        if (Ribbon2.ctff.Visible == true)
                            MessageBox.Show("Successful", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        else
                            SessionInfo.UserInfo.GlobalError += "Process:" + SessionInfo.UserInfo.ComName + " " + SessionInfo.UserInfo.MethodName + "(" + SessionInfo.UserInfo.CurrentRef + ") - Success! ";
                    }
                    if (strResu != null)
                        return true;
                    else
                        return false;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                if (Ribbon2.ctff.Visible == true)
                    MessageBox.Show(ex.ToString(), "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    SessionInfo.UserInfo.GlobalError += "Process:" + SessionInfo.UserInfo.ComName + " " + SessionInfo.UserInfo.MethodName + "(" + SessionInfo.UserInfo.CurrentRef + ") - Fail! " + ex.Message;

                LogHelper.WriteLog(typeof(CreateTextFileForm), ex.Message + "Do Post error");
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
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void contextMenuScript1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (currentColumnIndex >= 0)
                dgv.Columns.RemoveAt(currentColumnIndex);
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
                        List<RowCreateTextFile> newlist = new List<RowCreateTextFile>();
                        for (int i = 0; i < dgv.Rows.Count; i++)
                        {
                            RowCreateTextFile re = new RowCreateTextFile();
                            re.SunComponent = dgv.Rows[i].Cells[6].Value.ToString();
                            re.SunMethod = dgv.Rows[i].Cells[7].Value.ToString();
                            //re.DetailLad = new DetailLad();
                            if (dgv.Columns.Count > 11)
                                if ((dgv.Columns[11] != null) && (dgv.Columns[11].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[11].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column1 = invName;
                                        dgv.Rows[i].Cells[11].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[11].Name + ","))
                                    {
                                        re.Column1 = null;
                                        dgv.Rows[i].Cells[11].Value = null;
                                    }
                                    else
                                        re.Column1 = dgv.Rows[i].Cells[11] == null ? "" : dgv.Rows[i].Cells[11].Value.ToString();
                                }
                            if (dgv.Columns.Count > 12)
                                if ((dgv.Columns[12] != null) && dgv.Columns[12].Visible == true)
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[12].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column2 = invName;
                                        dgv.Rows[i].Cells[12].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[12].Name + ","))
                                    {
                                        re.Column2 = null;
                                        dgv.Rows[i].Cells[12].Value = null;
                                    }
                                    else
                                        re.Column2 = dgv.Rows[i].Cells[12] == null ? "" : dgv.Rows[i].Cells[12].Value.ToString();
                                }
                            if (dgv.Columns.Count > 13)
                                if ((dgv.Columns[13] != null) && (dgv.Columns[13].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[13].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column3 = invName;
                                        dgv.Rows[i].Cells[13].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[13].Name + ","))
                                    {
                                        re.Column3 = null;
                                        dgv.Rows[i].Cells[13].Value = null;
                                    }
                                    else
                                        re.Column3 = dgv.Rows[i].Cells[13] == null ? "" : dgv.Rows[i].Cells[13].Value.ToString();
                                }
                            if (dgv.Columns.Count > 14)
                                if ((dgv.Columns[14] != null) && (dgv.Columns[14].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[14].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column4 = invName;
                                        dgv.Rows[i].Cells[14].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[14].Name + ","))
                                    {
                                        re.Column4 = null;
                                        dgv.Rows[i].Cells[14].Value = null;
                                    }
                                    else
                                        re.Column4 = dgv.Rows[i].Cells[14] == null ? "" : dgv.Rows[i].Cells[14].Value.ToString();
                                }
                            if (dgv.Columns.Count > 15)
                                if ((dgv.Columns[15] != null) && (dgv.Columns[15].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[15].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column5 = invName;
                                        dgv.Rows[i].Cells[15].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[15].Name + ","))
                                    {
                                        re.Column5 = null;
                                        dgv.Rows[i].Cells[15].Value = null;
                                    }
                                    else
                                        re.Column5 = dgv.Rows[i].Cells[15] == null ? "" : dgv.Rows[i].Cells[15].Value.ToString();
                                }
                            if (dgv.Columns.Count > 16)
                                if ((dgv.Columns[16] != null) && (dgv.Columns[16].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[16].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column6 = invName;
                                        dgv.Rows[i].Cells[16].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[16].Name + ","))
                                    {
                                        re.Column6 = null;
                                        dgv.Rows[i].Cells[16].Value = null;
                                    }
                                    else
                                        re.Column6 = dgv.Rows[i].Cells[16] == null ? "" : dgv.Rows[i].Cells[16].Value.ToString();
                                }
                            if (dgv.Columns.Count > 17)
                                if ((dgv.Columns[17] != null) && (dgv.Columns[17].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[17].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column7 = invName;
                                        dgv.Rows[i].Cells[17].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[17].Name + ","))
                                    {
                                        re.Column7 = null;
                                        dgv.Rows[i].Cells[17].Value = null;
                                    }
                                    else
                                        re.Column7 = dgv.Rows[i].Cells[17] == null ? "" : dgv.Rows[i].Cells[17].Value.ToString();
                                }
                            if (dgv.Columns.Count > 18)
                                if ((dgv.Columns[18] != null) && (dgv.Columns[18].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[18].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column8 = invName;
                                        dgv.Rows[i].Cells[18].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[18].Name + ","))
                                    {
                                        re.Column8 = null;
                                        dgv.Rows[i].Cells[18].Value = null;
                                    }
                                    else
                                        re.Column8 = dgv.Rows[i].Cells[18] == null ? "" : dgv.Rows[i].Cells[18].Value.ToString();
                                }
                            if (dgv.Columns.Count > 19)
                                if ((dgv.Columns[19] != null) && (dgv.Columns[19].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[19].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column9 = invName;
                                        dgv.Rows[i].Cells[19].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[19].Name + ","))
                                    {
                                        re.Column9 = null;
                                        dgv.Rows[i].Cells[19].Value = null;
                                    }
                                    else
                                        re.Column9 = dgv.Rows[i].Cells[19] == null ? "" : dgv.Rows[i].Cells[19].Value.ToString();
                                }
                            if (dgv.Columns.Count > 20)
                                if ((dgv.Columns[20] != null) && (dgv.Columns[20].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[20].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column10 = invName;
                                        dgv.Rows[i].Cells[20].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[20].Name + ","))
                                    {
                                        re.Column10 = null;
                                        dgv.Rows[i].Cells[20].Value = null;
                                    }
                                    else
                                        re.Column10 = dgv.Rows[i].Cells[20] == null ? "" : dgv.Rows[i].Cells[20].Value.ToString();
                                }
                            if (dgv.Columns.Count > 21)
                                if ((dgv.Columns[21] != null) && (dgv.Columns[21].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[21].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column11 = invName;
                                        dgv.Rows[i].Cells[21].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[21].Name + ","))
                                    {
                                        re.Column11 = null;
                                        dgv.Rows[i].Cells[21].Value = null;
                                    }
                                    else
                                        re.Column11 = dgv.Rows[i].Cells[21] == null ? "" : dgv.Rows[i].Cells[21].Value.ToString();
                                }
                            if (dgv.Columns.Count > 22)
                                if ((dgv.Columns[22] != null) && (dgv.Columns[22].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[22].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column12 = invName;
                                        dgv.Rows[i].Cells[22].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[22].Name + ","))
                                    {
                                        re.Column12 = null;
                                        dgv.Rows[i].Cells[22].Value = null;
                                    }
                                    else
                                        re.Column12 = dgv.Rows[i].Cells[22] == null ? "" : dgv.Rows[i].Cells[22].Value.ToString();
                                }
                            if (dgv.Columns.Count > 23)
                                if ((dgv.Columns[23] != null) && (dgv.Columns[23].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[23].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column13 = invName;
                                        dgv.Rows[i].Cells[23].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[23].Name + ","))
                                    {
                                        re.Column13 = null;
                                        dgv.Rows[i].Cells[23].Value = null;
                                    }
                                    else
                                        re.Column13 = dgv.Rows[i].Cells[23] == null ? "" : dgv.Rows[i].Cells[23].Value.ToString();
                                }
                            if (dgv.Columns.Count > 24)
                                if ((dgv.Columns[24] != null) && (dgv.Columns[24].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[24].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column14 = invName;
                                        dgv.Rows[i].Cells[24].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[24].Name + ","))
                                    {
                                        re.Column14 = null;
                                        dgv.Rows[i].Cells[24].Value = null;
                                    }
                                    else
                                        re.Column14 = dgv.Rows[i].Cells[24] == null ? "" : dgv.Rows[i].Cells[24].Value.ToString();
                                }
                            if (dgv.Columns.Count > 25)
                                if ((dgv.Columns[25] != null) && (dgv.Columns[25].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[25].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column15 = invName;
                                        dgv.Rows[i].Cells[25].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[25].Name + ","))
                                    {
                                        re.Column15 = null;
                                        dgv.Rows[i].Cells[25].Value = null;
                                    }
                                    else
                                        re.Column15 = dgv.Rows[i].Cells[25] == null ? "" : dgv.Rows[i].Cells[25].Value.ToString();
                                }
                            if (dgv.Columns.Count > 26)
                                if ((dgv.Columns[26] != null) && (dgv.Columns[26].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[26].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column16 = invName;
                                        dgv.Rows[i].Cells[26].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[26].Name + ","))
                                    {
                                        re.Column16 = null;
                                        dgv.Rows[i].Cells[26].Value = null;
                                    }
                                    else
                                        re.Column16 = dgv.Rows[i].Cells[26] == null ? "" : dgv.Rows[i].Cells[26].Value.ToString();
                                }
                            if (dgv.Columns.Count > 27)
                                if ((dgv.Columns[27] != null) && (dgv.Columns[27].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[27].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column17 = invName;
                                        dgv.Rows[i].Cells[27].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[27].Name + ","))
                                    {
                                        re.Column17 = null;
                                        dgv.Rows[i].Cells[27].Value = null;
                                    }
                                    else
                                        re.Column17 = dgv.Rows[i].Cells[27] == null ? "" : dgv.Rows[i].Cells[27].Value.ToString();
                                }
                            if (dgv.Columns.Count > 28)
                                if ((dgv.Columns[28] != null) && (dgv.Columns[28].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[28].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column18 = invName;
                                        dgv.Rows[i].Cells[28].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[28].Name + ","))
                                    {
                                        re.Column18 = null;
                                        dgv.Rows[i].Cells[28].Value = null;
                                    }
                                    else
                                        re.Column18 = dgv.Rows[i].Cells[28] == null ? "" : dgv.Rows[i].Cells[28].Value.ToString();
                                }
                            if (dgv.Columns.Count > 29)
                                if ((dgv.Columns[29] != null) && (dgv.Columns[29].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[29].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column19 = invName;
                                        dgv.Rows[i].Cells[29].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[29].Name + ","))
                                    {
                                        re.Column19 = null;
                                        dgv.Rows[i].Cells[29].Value = null;
                                    }
                                    else
                                        re.Column19 = dgv.Rows[i].Cells[29] == null ? "" : dgv.Rows[i].Cells[29].Value.ToString();
                                }
                            if (dgv.Columns.Count > 30)
                                if ((dgv.Columns[30] != null) && (dgv.Columns[30].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[30].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column20 = invName;
                                        dgv.Rows[i].Cells[30].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[30].Name + ","))
                                    {
                                        re.Column20 = null;
                                        dgv.Rows[i].Cells[30].Value = null;
                                    }
                                    else
                                        re.Column20 = dgv.Rows[i].Cells[30] == null ? "" : dgv.Rows[i].Cells[30].Value.ToString();
                                }
                            if (dgv.Columns.Count > 31)
                                if ((dgv.Columns[31] != null) && (dgv.Columns[31].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[31].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column21 = invName;
                                        dgv.Rows[i].Cells[31].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[31].Name + ","))
                                    {
                                        re.Column21 = null;
                                        dgv.Rows[i].Cells[31].Value = null;
                                    }
                                    else
                                        re.Column21 = dgv.Rows[i].Cells[31] == null ? "" : dgv.Rows[i].Cells[31].Value.ToString();
                                }
                            if (dgv.Columns.Count > 32)
                                if ((dgv.Columns[32] != null) && (dgv.Columns[32].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[32].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column22 = invName;
                                        dgv.Rows[i].Cells[32].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[32].Name + ","))
                                    {
                                        re.Column22 = null;
                                        dgv.Rows[i].Cells[32].Value = null;
                                    }
                                    else
                                        re.Column22 = dgv.Rows[i].Cells[32] == null ? "" : dgv.Rows[i].Cells[32].Value.ToString();
                                }
                            if (dgv.Columns.Count > 33)
                                if ((dgv.Columns[33] != null) && (dgv.Columns[33].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[33].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column23 = invName;
                                        dgv.Rows[i].Cells[33].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[33].Name + ","))
                                    {
                                        re.Column23 = null;
                                        dgv.Rows[i].Cells[33].Value = null;
                                    }
                                    else
                                        re.Column23 = dgv.Rows[i].Cells[33] == null ? "" : dgv.Rows[i].Cells[33].Value.ToString();
                                }
                            if (dgv.Columns.Count > 34)
                                if ((dgv.Columns[34] != null) && (dgv.Columns[34].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[34].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column24 = invName;
                                        dgv.Rows[i].Cells[34].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[34].Name + ","))
                                    {
                                        re.Column24 = null;
                                        dgv.Rows[i].Cells[34].Value = null;
                                    }
                                    else
                                        re.Column24 = dgv.Rows[i].Cells[34] == null ? "" : dgv.Rows[i].Cells[34].Value.ToString();
                                }
                            if (dgv.Columns.Count > 35)
                                if ((dgv.Columns[35] != null) && (dgv.Columns[35].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[35].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column25 = invName;
                                        dgv.Rows[i].Cells[35].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[35].Name + ","))
                                    {
                                        re.Column25 = null;
                                        dgv.Rows[i].Cells[35].Value = null;
                                    }
                                    else
                                        re.Column25 = dgv.Rows[i].Cells[35] == null ? "" : dgv.Rows[i].Cells[35].Value.ToString();
                                }
                            if (dgv.Columns.Count > 36)
                                if ((dgv.Columns[36] != null) && (dgv.Columns[36].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[36].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column26 = invName;
                                        dgv.Rows[i].Cells[36].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[36].Name + ","))
                                    {
                                        re.Column26 = null;
                                        dgv.Rows[i].Cells[36].Value = null;
                                    }
                                    else
                                        re.Column26 = dgv.Rows[i].Cells[36] == null ? "" : dgv.Rows[i].Cells[36].Value.ToString();
                                }
                            if (dgv.Columns.Count > 37)
                                if ((dgv.Columns[37] != null) && (dgv.Columns[37].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[37].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column27 = invName;
                                        dgv.Rows[i].Cells[37].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[37].Name + ","))
                                    {
                                        re.Column27 = null;
                                        dgv.Rows[i].Cells[37].Value = null;
                                    }
                                    else
                                        re.Column27 = dgv.Rows[i].Cells[37] == null ? "" : dgv.Rows[i].Cells[37].Value.ToString();
                                }
                            if (dgv.Columns.Count > 38)
                                if ((dgv.Columns[38] != null) && (dgv.Columns[38].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[38].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column28 = invName;
                                        dgv.Rows[i].Cells[38].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[38].Name + ","))
                                    {
                                        re.Column28 = null;
                                        dgv.Rows[i].Cells[38].Value = null;
                                    }
                                    else
                                        re.Column28 = dgv.Rows[i].Cells[38] == null ? "" : dgv.Rows[i].Cells[38].Value.ToString();
                                }
                            if (dgv.Columns.Count > 39)
                                if ((dgv.Columns[39] != null) && (dgv.Columns[39].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[39].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column29 = invName;
                                        dgv.Rows[i].Cells[39].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[39].Name + ","))
                                    {
                                        re.Column29 = null;
                                        dgv.Rows[i].Cells[39].Value = null;
                                    }
                                    else
                                        re.Column29 = dgv.Rows[i].Cells[39] == null ? "" : dgv.Rows[i].Cells[39].Value.ToString();
                                }
                            if (dgv.Columns.Count > 40)
                                if ((dgv.Columns[40] != null) && (dgv.Columns[40].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[40].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column30 = invName;
                                        dgv.Rows[i].Cells[40].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[40].Name + ","))
                                    {
                                        re.Column30 = null;
                                        dgv.Rows[i].Cells[40].Value = null;
                                    }
                                    else
                                        re.Column30 = dgv.Rows[i].Cells[40] == null ? "" : dgv.Rows[i].Cells[40].Value.ToString();
                                }
                            if (dgv.Columns.Count > 41)
                                if ((dgv.Columns[41] != null) && (dgv.Columns[41].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[41].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column31 = invName;
                                        dgv.Rows[i].Cells[41].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[41].Name + ","))
                                    {
                                        re.Column31 = null;
                                        dgv.Rows[i].Cells[41].Value = null;
                                    }
                                    else
                                        re.Column31 = dgv.Rows[i].Cells[41] == null ? "" : dgv.Rows[i].Cells[41].Value.ToString();
                                }
                            if (dgv.Columns.Count > 42)
                                if ((dgv.Columns[42] != null) && (dgv.Columns[42].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[42].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column32 = invName;
                                        dgv.Rows[i].Cells[42].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[42].Name + ","))
                                    {
                                        re.Column32 = null;
                                        dgv.Rows[i].Cells[42].Value = null;
                                    }
                                    else
                                        re.Column32 = dgv.Rows[i].Cells[42] == null ? "" : dgv.Rows[i].Cells[42].Value.ToString();
                                }
                            if (dgv.Columns.Count > 43)
                                if ((dgv.Columns[43] != null) && (dgv.Columns[43].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[43].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column33 = invName;
                                        dgv.Rows[i].Cells[43].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[43].Name + ","))
                                    {
                                        re.Column33 = null;
                                        dgv.Rows[i].Cells[43].Value = null;
                                    }
                                    else
                                        re.Column33 = dgv.Rows[i].Cells[43] == null ? "" : dgv.Rows[i].Cells[43].Value.ToString();
                                }
                            if (dgv.Columns.Count > 44)
                                if ((dgv.Columns[44] != null) && (dgv.Columns[44].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[44].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column34 = invName;
                                        dgv.Rows[i].Cells[44].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[44].Name + ","))
                                    {
                                        re.Column34 = null;
                                        dgv.Rows[i].Cells[44].Value = null;
                                    }
                                    else
                                        re.Column34 = dgv.Rows[i].Cells[44] == null ? "" : dgv.Rows[i].Cells[44].Value.ToString();
                                }
                            if (dgv.Columns.Count > 45)
                                if ((dgv.Columns[45] != null) && (dgv.Columns[45].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[45].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column35 = invName;
                                        dgv.Rows[i].Cells[45].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[45].Name + ","))
                                    {
                                        re.Column35 = null;
                                        dgv.Rows[i].Cells[45].Value = null;
                                    }
                                    else
                                        re.Column35 = dgv.Rows[i].Cells[45] == null ? "" : dgv.Rows[i].Cells[45].Value.ToString();
                                }
                            if (dgv.Columns.Count > 46)
                                if ((dgv.Columns[46] != null) && (dgv.Columns[46].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[46].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column36 = invName;
                                        dgv.Rows[i].Cells[46].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[46].Name + ","))
                                    {
                                        re.Column36 = null;
                                        dgv.Rows[i].Cells[46].Value = null;
                                    }
                                    else
                                        re.Column36 = dgv.Rows[i].Cells[46] == null ? "" : dgv.Rows[i].Cells[46].Value.ToString();
                                }
                            if (dgv.Columns.Count > 47)
                                if ((dgv.Columns[47] != null) && (dgv.Columns[47].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[47].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column37 = invName;
                                        dgv.Rows[i].Cells[47].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[47].Name + ","))
                                    {
                                        re.Column37 = null;
                                        dgv.Rows[i].Cells[47].Value = null;
                                    }
                                    else
                                        re.Column37 = dgv.Rows[i].Cells[47] == null ? "" : dgv.Rows[i].Cells[47].Value.ToString();
                                }
                            if (dgv.Columns.Count > 48)
                                if ((dgv.Columns[48] != null) && (dgv.Columns[48].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[48].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column38 = invName;
                                        dgv.Rows[i].Cells[48].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[48].Name + ","))
                                    {
                                        re.Column38 = null;
                                        dgv.Rows[i].Cells[48].Value = null;
                                    }
                                    else
                                        re.Column38 = dgv.Rows[i].Cells[48] == null ? "" : dgv.Rows[i].Cells[48].Value.ToString();
                                }
                            if (dgv.Columns.Count > 49)
                                if ((dgv.Columns[49] != null) && (dgv.Columns[49].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[49].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column39 = invName;
                                        dgv.Rows[i].Cells[49].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[49].Name + ","))
                                    {
                                        re.Column39 = null;
                                        dgv.Rows[i].Cells[49].Value = null;
                                    }
                                    else
                                        re.Column39 = dgv.Rows[i].Cells[49] == null ? "" : dgv.Rows[i].Cells[49].Value.ToString();
                                }
                            if (dgv.Columns.Count > 50)
                                if ((dgv.Columns[50] != null) && (dgv.Columns[50].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[50].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column40 = invName;
                                        dgv.Rows[i].Cells[50].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[50].Name + ","))
                                    {
                                        re.Column40 = null;
                                        dgv.Rows[i].Cells[50].Value = null;
                                    }
                                    else
                                        re.Column40 = dgv.Rows[i].Cells[50] == null ? "" : dgv.Rows[i].Cells[50].Value.ToString();
                                }
                            if (dgv.Columns.Count > 51)
                                if ((dgv.Columns[51] != null) && (dgv.Columns[51].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[51].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column41 = invName;
                                        dgv.Rows[i].Cells[51].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[51].Name + ","))
                                    {
                                        re.Column41 = null;
                                        dgv.Rows[i].Cells[51].Value = null;
                                    }
                                    else
                                        re.Column41 = dgv.Rows[i].Cells[51] == null ? "" : dgv.Rows[i].Cells[51].Value.ToString();
                                }
                            if (dgv.Columns.Count > 52)
                                if ((dgv.Columns[52] != null) && (dgv.Columns[52].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[52].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column42 = invName;
                                        dgv.Rows[i].Cells[52].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[52].Name + ","))
                                    {
                                        re.Column42 = null;
                                        dgv.Rows[i].Cells[52].Value = null;
                                    }
                                    else
                                        re.Column42 = dgv.Rows[i].Cells[52] == null ? "" : dgv.Rows[i].Cells[52].Value.ToString();
                                }
                            if (dgv.Columns.Count > 53)
                                if ((dgv.Columns[53] != null) && (dgv.Columns[53].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[53].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column43 = invName;
                                        dgv.Rows[i].Cells[53].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[53].Name + ","))
                                    {
                                        re.Column43 = null;
                                        dgv.Rows[i].Cells[53].Value = null;
                                    }
                                    else
                                        re.Column43 = dgv.Rows[i].Cells[53] == null ? "" : dgv.Rows[i].Cells[53].Value.ToString();
                                }
                            if (dgv.Columns.Count > 54)
                                if ((dgv.Columns[54] != null) && (dgv.Columns[54].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[54].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column44 = invName;
                                        dgv.Rows[i].Cells[54].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[54].Name + ","))
                                    {
                                        re.Column44 = null;
                                        dgv.Rows[i].Cells[54].Value = null;
                                    }
                                    else
                                        re.Column44 = dgv.Rows[i].Cells[54] == null ? "" : dgv.Rows[i].Cells[54].Value.ToString();
                                }
                            if (dgv.Columns.Count > 55)
                                if ((dgv.Columns[55] != null) && (dgv.Columns[55].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[55].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column45 = invName;
                                        dgv.Rows[i].Cells[55].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[55].Name + ","))
                                    {
                                        re.Column45 = null;
                                        dgv.Rows[i].Cells[55].Value = null;
                                    }
                                    else
                                        re.Column45 = dgv.Rows[i].Cells[55] == null ? "" : dgv.Rows[i].Cells[55].Value.ToString();
                                }
                            if (dgv.Columns.Count > 56)
                                if ((dgv.Columns[56] != null) && (dgv.Columns[56].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[56].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column46 = invName;
                                        dgv.Rows[i].Cells[56].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[56].Name + ","))
                                    {
                                        re.Column46 = null;
                                        dgv.Rows[i].Cells[56].Value = null;
                                    }
                                    else
                                        re.Column46 = dgv.Rows[i].Cells[56] == null ? "" : dgv.Rows[i].Cells[56].Value.ToString();
                                }
                            if (dgv.Columns.Count > 57)
                                if ((dgv.Columns[57] != null) && (dgv.Columns[57].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[57].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column47 = invName;
                                        dgv.Rows[i].Cells[57].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[57].Name + ","))
                                    {
                                        re.Column47 = null;
                                        dgv.Rows[i].Cells[57].Value = null;
                                    }
                                    else
                                        re.Column47 = dgv.Rows[i].Cells[57] == null ? "" : dgv.Rows[i].Cells[57].Value.ToString();
                                }
                            if (dgv.Columns.Count > 58)
                                if ((dgv.Columns[58] != null) && (dgv.Columns[58].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[58].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column48 = invName;
                                        dgv.Rows[i].Cells[58].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[58].Name + ","))
                                    {
                                        re.Column48 = null;
                                        dgv.Rows[i].Cells[58].Value = null;
                                    }
                                    else
                                        re.Column48 = dgv.Rows[i].Cells[58] == null ? "" : dgv.Rows[i].Cells[58].Value.ToString();
                                }
                            if (dgv.Columns.Count > 59)
                                if ((dgv.Columns[59] != null) && (dgv.Columns[59].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[59].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column49 = invName;
                                        dgv.Rows[i].Cells[59].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[59].Name + ","))
                                    {
                                        re.Column49 = null;
                                        dgv.Rows[i].Cells[59].Value = null;
                                    }
                                    else
                                        re.Column49 = dgv.Rows[i].Cells[59] == null ? "" : dgv.Rows[i].Cells[59].Value.ToString();
                                }
                            if (dgv.Columns.Count > 60)
                                if ((dgv.Columns[60] != null) && (dgv.Columns[60].Visible == true))
                                {
                                    if (SessionInfo.UserInfo.UseSequenceNumbering == "1" && sequenceNumbering.Contains(dgv.Columns[60].Name + ","))
                                    {
                                        int? ii = 0;
                                        string invName = string.Empty;
                                        string prefix = string.Empty;
                                        try
                                        {
                                            ft.GetInvoiceInfo(ref prefix, ref invName, ref ii);
                                        }
                                        catch
                                        { }
                                        re.Column50 = invName;
                                        dgv.Rows[i].Cells[60].Value = invName;
                                    }
                                    else if (SessionInfo.UserInfo.UseSequenceNumbering == "0" && sequenceNumbering.Contains(dgv.Columns[60].Name + ","))
                                    {
                                        re.Column50 = null;
                                        dgv.Rows[i].Cells[60].Value = null;
                                    }
                                    else
                                        re.Column50 = dgv.Rows[i].Cells[60] == null ? "" : dgv.Rows[i].Cells[60].Value.ToString();
                                }
                            newlist.Add(re);
                        }
                        string script = ft.GetXMLProfileScript(newlist);
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
