using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;
using System.IO;
using ExcelAddIn3.Common;
using System.Text.RegularExpressions;

namespace ExcelAddIn3
{
    public partial class Finance_Tools
    {
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private bool StartService()
        {
            try
            {
                System.ServiceProcess.ServiceController myController =
                  new System.ServiceProcess.ServiceController("SunSystems Connect Server", SessionInfo.UserInfo.SunUserIP);//SunSystems Connect Server RSI RicoSimp1

                if (myController.Status.Equals(System.ServiceProcess.ServiceControllerStatus.Stopped) || myController.Status.Equals(System.ServiceProcess.ServiceControllerStatus.StopPending))
                {
                    myController.Start();
                    Thread.Sleep(10000);
                }
            }
            catch { MessageBox.Show("SunSystems Connect Server Error!"); return false; }
            return true;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private bool DoPost()
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
                    this.textBox1.Text = "Password for user is incorrect.";
                    return false;
                }
            }
            catch (Exception ex)
            {
                this.textBox1.Text = "An error occurred in validation " + oSecMan.ErrorMessage + ex;
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

                // 2. Create and format of WSDL document.
                ServiceDescription description = ServiceDescription.Read(stream);

                // 3. Create a client proxy proxy class.
                ServiceDescriptionImporter importer = new ServiceDescriptionImporter();

                importer.ProtocolName = "Soap"; // The specified access protocol.
                importer.Style = ServiceDescriptionImportStyle.Client; // To generate client proxy.
                importer.CodeGenerationOptions = CodeGenerationOptions.GenerateProperties | CodeGenerationOptions.GenerateNewAsync;

                importer.AddServiceDescription(description, null, null); // Add a WSDL document.

                // 4. Compile the client proxy classes using CodeDom.
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

                // 5. Using Reflection to call WebService.
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
                    if (!string.IsNullOrEmpty(this.textBox1.Text.Trim()))
                    {
                        return false;
                    }
                    else
                    {
                        try
                        {
                            //Save journal number
                            XmlDocument xdoc = new XmlDocument();
                            xdoc.LoadXml(strResu.ToString());
                            //XmlNode node = xdoc.DocumentElement;
                            //XmlNode node = xdoc.SelectSingleNode("Nodes");
                            journalNumber = xdoc.GetElementsByTagName("JournalNumber").Item(1).InnerText;

                            if (!string.IsNullOrEmpty(journalNumber))
                            {
                                SessionInfo.UserInfo.SunJournalNumber = journalNumber;
                                MessageBox.Show("journalNumber:" + journalNumber, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            //present journal number
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
                MessageBox.Show(ex.ToString(), "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //EventLog.WriteEntry("Finance Tool", ex.Message + "Do Post error", EventLogEntryType.Error);
                LogHelper.WriteLog(typeof(Ribbon1), ex.Message + "Do Post error");
                return false;
            }
        }
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
        public void SaveHistory()
        {
            if ((SessionInfo.UserInfo.UseSequenceNumbering == "1") || SessionInfo.UserInfo.UseCriteria)
            {
                int? i = 0;
                string invName = string.Empty;
                string prefix = string.Empty;
                try
                {
                    GetInvoiceInfo(ref prefix, ref invName, ref i);
                }
                catch
                {
                }
                if (!string.IsNullOrEmpty(invName))
                {
                    PopulateCell(invName);
                }
                Globals.ThisAddIn.Application.ActiveWorkbook.Save();
                SqlConnection conn = null;
                SqlDataReader rdr = null;
                try
                {
                    // create and open a connection object
                    conn = new
                        SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                    conn.Open();

                    string pathvalue = string.Empty;
                    if (SessionInfo.UserInfo.Dictionary.dict.Count != 0 && SessionInfo.UserInfo.Dictionary.dict.ContainsKey(SessionInfo.UserInfo.CachePath))
                    {
                        pathvalue = SessionInfo.UserInfo.Dictionary.dict[SessionInfo.UserInfo.CachePath];
                        string[] sArray = Regex.Split(pathvalue, ",");
                        pathvalue = sArray[0];
                    }
                    if (SessionInfo.UserInfo.CachePath == pathvalue || string.IsNullOrEmpty(pathvalue))
                    {
                        // 1. create a command object identifying
                        // the stored procedure
                        SqlCommand cmd = new SqlCommand("FT_Templates_Ins", conn);

                        // 2. set the command object so it knows
                        // to execute a stored procedure
                        cmd.CommandType = CommandType.StoredProcedure;
                        // 3. add parameter to command, which
                        // will be passed to the stored procedure
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
                        cmd.Parameters.Add(new SqlParameter("@Data", ft.GetData(SessionInfo.UserInfo.CachePath)));
                        cmd.Parameters.Add(new SqlParameter("@DataType", Path.GetExtension(SessionInfo.UserInfo.FilePath)));
                        cmd.Parameters.Add(new SqlParameter("@PDFData", ft.GetData(SessionInfo.UserInfo.Containerpath)));
                        cmd.Parameters.Add(new SqlParameter("@XMLData", this.txtXML.Text));
                        //cmd.Parameters.Add(new SqlParameter("@OwnUserID", SessionInfo.UserInfo.ID));
                        cmd.Parameters.Add(new SqlParameter("@TemplatePath", SessionInfo.UserInfo.FilePath));
                        cmd.Parameters.Add(new SqlParameter("@maxNum", i));
                        cmd.Parameters.Add(new SqlParameter("@TransactionName", invName));
                        cmd.Parameters.Add(new SqlParameter("@Prefix", prefix));
                        cmd.Parameters.Add(new SqlParameter("@SunJournalNumber", SessionInfo.UserInfo.SunJournalNumber));
                        // execute the command
                        rdr = cmd.ExecuteReader();
                    }
                    else
                    {
                        // 1. create a command object identifying
                        // the stored procedure
                        SqlCommand cmd = new SqlCommand("FT_Templates_Del", conn);

                        // 2. set the command object so it knows
                        // to execute a stored procedure
                        cmd.CommandType = CommandType.StoredProcedure;
                        // 3. add parameter to command, which
                        // will be passed to the stored procedure
                        cmd.Parameters.Add(new SqlParameter("@TemplatePath", SessionInfo.UserInfo.FilePath));
                        cmd.Parameters.Add(new SqlParameter("@TransactionName", invName));
                        // execute the command
                        rdr = cmd.ExecuteReader();

                        // 1. create a command object identifying
                        // the stored procedure
                        SqlCommand cmd2 = new SqlCommand("FT_Templates_Ins", conn);
                        // 2. set the command object so it knows
                        // to execute a stored procedure
                        cmd2.CommandType = CommandType.StoredProcedure;
                        // 3. add parameter to command, which
                        // will be passed to the stored procedure
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
                        cmd2.Parameters.Add(new SqlParameter("@Data", GetData(SessionInfo.UserInfo.CachePath)));
                        cmd2.Parameters.Add(new SqlParameter("@DataType", Path.GetExtension(SessionInfo.UserInfo.FilePath)));
                        cmd2.Parameters.Add(new SqlParameter("@PDFData", GetData(SessionInfo.UserInfo.Containerpath)));
                        cmd2.Parameters.Add(new SqlParameter("@XMLData", this.txtXML.Text));
                        //cmd.Parameters.Add(new SqlParameter("@OwnUserID", SessionInfo.UserInfo.ID));
                        cmd2.Parameters.Add(new SqlParameter("@TemplatePath", SessionInfo.UserInfo.FilePath));
                        cmd2.Parameters.Add(new SqlParameter("@maxNum", i));
                        cmd2.Parameters.Add(new SqlParameter("@TransactionName", invName));
                        cmd2.Parameters.Add(new SqlParameter("@Prefix", prefix));
                        cmd2.Parameters.Add(new SqlParameter("@SunJournalNumber", SessionInfo.UserInfo.SunJournalNumber));
                        // execute the command
                        rdr = cmd2.ExecuteReader();
                    }

                    if (!string.IsNullOrEmpty(invName))
                    {
                        MessageBox.Show("Your new Template Title is " + invName, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        //PopulateCell(invName);
                    }
                    string str = GetCriteriaStr();
                    if (!string.IsNullOrEmpty(str))
                    {
                        MessageBox.Show("Your new Template Criterias is: \r\n " + str, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        if (!string.IsNullOrEmpty(Ribbon1.outputPane.DropDownCriteria1.Text))
                            Ribbon1.outputPane.DropDownCriteria1.Enabled = false;
                        if (!string.IsNullOrEmpty(Ribbon1.outputPane.DropDownCriteria2.Text))
                            Ribbon1.outputPane.DropDownCriteria2.Enabled = false;
                        if (!string.IsNullOrEmpty(Ribbon1.outputPane.DropDownCriteria3.Text))
                            Ribbon1.outputPane.DropDownCriteria3.Enabled = false;
                        if (!string.IsNullOrEmpty(Ribbon1.outputPane.DropDownCriteria4.Text))
                            Ribbon1.outputPane.DropDownCriteria4.Enabled = false;
                        if (!string.IsNullOrEmpty(Ribbon1.outputPane.DropDownCriteria5.Text))
                            Ribbon1.outputPane.DropDownCriteria5.Enabled = false;
                    }
                }
                catch { }
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
                    //SessionInfo.UserInfo.IsOnlySaveHistory = false;
                }
            }
        }
    }
}
