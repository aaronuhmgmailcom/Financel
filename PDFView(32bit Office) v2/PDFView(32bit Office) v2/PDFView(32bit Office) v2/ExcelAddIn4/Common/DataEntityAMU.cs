
/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<SSC>   
 * Author：Peter.uhm  (yanb@shinetechchina.com)
 * Modify date：2015.10
 * Version : 2.0.0.2
 */
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelAddIn4.Common3
{

    public class SSC
    {
        public ErrorContext ErrorContext
        {
            get;
            set;
        }
        public User User
        {
            get;
            set;
        }
        public SunSystemsContext SunSystemsContext
        {
            get;
            set;
        }
        public List<AccountAllocations> Payload
        {
            get;
            set;
        }
        public string ErrorMessages
        {
            get;
            set;
        }
    }
    public class ErrorContext
    {
        public string CompatibilityMode
        {
            get;
            set;
        }
        public string ErrorOutput
        {
            get;
            set;
        }
        public string ErrorThreshold
        {
            get;
            set;
        }
    }
    public class User
    {
        public string Name
        {
            get;
            set;
        }
    }
    public class SunSystemsContext
    {
        public string BusinessUnit
        {
            get;
            set;
        }
        //public string BudgetCode
        //{
        //    get;
        //    set;
        //}
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        //public DataGridView IniGrdForAllocationMakerUpd()
        //{
        //    try
        //    {
        //        DataTable dt = GetUserDataFriendlyName();
        //        DataColumn[] keys = new DataColumn[1];
        //        keys[0] = dt.Columns["SunField"];
        //        dt.PrimaryKey = keys;
        //        DataGridView d = new DataGridView();
        //        if (dt.Rows.Count > 0)
        //        {
        //            d.Columns.Add("LineIndicator", "Line Indicator");
        //            d.Columns["LineIndicator"].DataPropertyName = "LineIndicator";
        //            d.Columns.Add("JournalNumber", "JournalNumber");
        //            d.Columns["JournalNumber"].DataPropertyName = "JournalNumber";
        //            d.Columns.Add("JournalLineNumber", "JournalLineNumber");
        //            d.Columns["JournalLineNumber"].DataPropertyName = "JournalLineNumber";
        //            d.Columns.Add("Ledger", dt.Rows.Find("Ledger")["UserFriendlyName"].ToString());
        //            d.Columns["Ledger"].DataPropertyName = "Ledger";
        //            d.Columns.Add("AccountCode", dt.Rows.Find("AccountCode")["UserFriendlyName"].ToString());
        //            d.Columns["AccountCode"].DataPropertyName = "ft_Account";
        //            d.Columns.Add("AccountingPeriod", dt.Rows.Find("AccountingPeriod")["UserFriendlyName"].ToString());
        //            d.Columns["AccountingPeriod"].DataPropertyName = "Period";
        //            d.Columns.Add("TransactionDate", dt.Rows.Find("TransactionDate")["UserFriendlyName"].ToString());
        //            d.Columns["TransactionDate"].DataPropertyName = "TransactionDate";
        //            d.Columns.Add("JournalType", dt.Rows.Find("JournalType")["UserFriendlyName"].ToString());
        //            d.Columns["JournalType"].DataPropertyName = "JrnlType";
        //            d.Columns.Add("TransactionReference", dt.Rows.Find("TransactionReference")["UserFriendlyName"].ToString());
        //            d.Columns["TransactionReference"].DataPropertyName = "TransRef";
        //            d.Columns.Add("AnalysisCode1", dt.Rows.Find("AnalysisCode1")["UserFriendlyName"].ToString());
        //            d.Columns["AnalysisCode1"].DataPropertyName = "LA1";
        //            d.Columns.Add("AnalysisCode2", dt.Rows.Find("AnalysisCode2")["UserFriendlyName"].ToString());
        //            d.Columns["AnalysisCode2"].DataPropertyName = "LA2";
        //            d.Columns.Add("AnalysisCode3", dt.Rows.Find("AnalysisCode3")["UserFriendlyName"].ToString());
        //            d.Columns["AnalysisCode3"].DataPropertyName = "LA3";
        //            d.Columns.Add("AnalysisCode4", dt.Rows.Find("AnalysisCode4")["UserFriendlyName"].ToString());
        //            d.Columns["AnalysisCode4"].DataPropertyName = "LA4";
        //            d.Columns.Add("AnalysisCode5", dt.Rows.Find("AnalysisCode5")["UserFriendlyName"].ToString());
        //            d.Columns["AnalysisCode5"].DataPropertyName = "LA5";
        //            d.Columns.Add("AnalysisCode6", dt.Rows.Find("AnalysisCode6")["UserFriendlyName"].ToString());
        //            d.Columns["AnalysisCode6"].DataPropertyName = "LA6";
        //            d.Columns.Add("AnalysisCode7", dt.Rows.Find("AnalysisCode7")["UserFriendlyName"].ToString());
        //            d.Columns["AnalysisCode7"].DataPropertyName = "LA7";
        //            d.Columns.Add("AnalysisCode8", dt.Rows.Find("AnalysisCode8")["UserFriendlyName"].ToString());
        //            d.Columns["AnalysisCode8"].DataPropertyName = "LA8";
        //            d.Columns.Add("AnalysisCode9", dt.Rows.Find("AnalysisCode9")["UserFriendlyName"].ToString());
        //            d.Columns["AnalysisCode9"].DataPropertyName = "LA9";
        //            d.Columns.Add("AnalysisCode10", dt.Rows.Find("AnalysisCode10")["UserFriendlyName"].ToString());
        //            d.Columns["AnalysisCode10"].DataPropertyName = "LA10";
        //            d.Columns.Add("AllocationMarker", dt.Rows.Find("AllocationMarker")["UserFriendlyName"].ToString());
        //            d.Columns["AllocationMarker"].DataPropertyName = "AlloctnMarker";
        //        }
        //        else
        //        {
        //            d.Columns.Add("LineIndicator", "Line Indicator");
        //            d.Columns["LineIndicator"].DataPropertyName = "LineIndicator";
        //            d.Columns.Add("JournalNumber", "JournalNumber");
        //            d.Columns["JournalNumber"].DataPropertyName = "JournalNumber";
        //            d.Columns.Add("JournalLineNumber", "JournalLineNumber");
        //            d.Columns["JournalLineNumber"].DataPropertyName = "JournalLineNumber";
        //            d.Columns.Add("Ledger", "Ledger");
        //            d.Columns["Ledger"].DataPropertyName = "Ledger";
        //            d.Columns.Add("AccountCode", "Account");
        //            d.Columns["AccountCode"].DataPropertyName = "ft_Account";
        //            d.Columns.Add("AccountingPeriod", "Period");
        //            d.Columns["AccountingPeriod"].DataPropertyName = "Period";
        //            d.Columns.Add("TransactionDate", "Trans Date");
        //            d.Columns["TransactionDate"].DataPropertyName = "TransactionDate";
        //            d.Columns.Add("JournalType", "Jrnl Type");
        //            d.Columns["JournalType"].DataPropertyName = "JrnlType";
        //            d.Columns.Add("TransactionReference", "Trans Ref");
        //            d.Columns["TransactionReference"].DataPropertyName = "TransRef";
        //            d.Columns.Add("AnalysisCode1", "LA1");
        //            d.Columns["AnalysisCode1"].DataPropertyName = "LA1";
        //            d.Columns.Add("AnalysisCode2", "LA2");
        //            d.Columns["AnalysisCode2"].DataPropertyName = "LA2";
        //            d.Columns.Add("AnalysisCode3", "LA3");
        //            d.Columns["AnalysisCode3"].DataPropertyName = "LA3";
        //            d.Columns.Add("AnalysisCode4", "LA4");
        //            d.Columns["AnalysisCode4"].DataPropertyName = "LA4";
        //            d.Columns.Add("AnalysisCode5", "LA5");
        //            d.Columns["AnalysisCode5"].DataPropertyName = "LA5";
        //            d.Columns.Add("AnalysisCode6", "LA6");
        //            d.Columns["AnalysisCode6"].DataPropertyName = "LA6";
        //            d.Columns.Add("AnalysisCode7", "LA7");
        //            d.Columns["AnalysisCode7"].DataPropertyName = "LA7";
        //            d.Columns.Add("AnalysisCode8", "LA8");
        //            d.Columns["AnalysisCode8"].DataPropertyName = "LA8";
        //            d.Columns.Add("AnalysisCode9", "LA9");
        //            d.Columns["AnalysisCode9"].DataPropertyName = "LA9";
        //            d.Columns.Add("AnalysisCode10", "LA10");
        //            d.Columns["AnalysisCode10"].DataPropertyName = "LA10";
        //            d.Columns.Add("AllocationMarker", "Alloctn Marker");
        //            d.Columns["AllocationMarker"].DataPropertyName = "AlloctnMarker";
        //        }
        //        d.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
        //        d.AutoGenerateColumns = false;
        //        d.ColumnHeadersHeight = 40;
        //        d.Dock = DockStyle.Fill;
        //        d.Visible = true;
        //        for (int i = 0; i < d.Columns.Count; i++)
        //        {
        //            d.Columns[i].Width = 55;
        //        }
        //        return d;
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception(ex.Message + ex.StackTrace);
        //    }
        //}
        ///// <summary>
        ///// 
        ///// </summary>
        ///// <returns></returns>
        //public static string GetJournalQuery(string xml)
        //{
        //    //tabControl1_SelectedIndexChanged(null, null);
        //    string sSSCVoucher = "";
        //    SunSystems.Connect.Client.SecurityManager oSecMan = new SunSystems.Connect.Client.SecurityManager(SessionInfo.UserInfo.SunUserIP);
        //    ////http://95.138.187.185:81/SecurityWebServer/Login.aspx?redirect=http://95.138.187.185:8080/ssc/login.jsp
        //    try
        //    {
        //        oSecMan.Login(SessionInfo.UserInfo.SunUserID, SessionInfo.UserInfo.SunUserPass);
        //        if (oSecMan.Authorised)
        //        {
        //            sSSCVoucher = oSecMan.Voucher;
        //            //label1.Content = sSSCVoucher;RSI, RicoSimp2
        //        }
        //        else
        //        {
        //            throw new Exception("SunSystems Server is not exist or Password for user is incorrect.");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception("An error occurred in validation " + oSecMan.ErrorMessage + ex);
        //    }
        //    finally
        //    {
        //        oSecMan = null;
        //    }

        //    try
        //    {
        //        WebClient web = new WebClient();
        //        Stream stream = web.OpenRead("http://" + SessionInfo.UserInfo.SunUserIP + ":8080/connect/wsdl/ComponentExecutor?wsdl");

        //        // 2. Create and format of WSDL document.
        //        ServiceDescription description = ServiceDescription.Read(stream);

        //        // 3. Create a client proxy proxy class.
        //        ServiceDescriptionImporter importer = new ServiceDescriptionImporter();

        //        importer.ProtocolName = "Soap"; // The specified access protocol.
        //        importer.Style = ServiceDescriptionImportStyle.Client; // To generate client proxy.
        //        importer.CodeGenerationOptions = CodeGenerationOptions.GenerateProperties | CodeGenerationOptions.GenerateNewAsync;

        //        importer.AddServiceDescription(description, null, null); // Add a WSDL document.

        //        // 4. Compile the client proxy classes using CodeDom.
        //        CodeNamespace nmspace = new CodeNamespace(); // Add a namespace for the proxy class, the default is the global space.
        //        CodeCompileUnit unit = new CodeCompileUnit();
        //        unit.Namespaces.Add(nmspace);

        //        ServiceDescriptionImportWarnings warning = importer.Import(nmspace, unit);
        //        CodeDomProvider provider = CodeDomProvider.CreateProvider("CSharp");

        //        CompilerParameters parameter = new CompilerParameters();
        //        parameter.GenerateExecutable = false;
        //        parameter.GenerateInMemory = true;
        //        parameter.ReferencedAssemblies.Add("System.dll");
        //        parameter.ReferencedAssemblies.Add("System.XML.dll");
        //        parameter.ReferencedAssemblies.Add("System.Web.Services.dll");
        //        parameter.ReferencedAssemblies.Add("System.Data.dll");

        //        CompilerResults result = provider.CompileAssemblyFromDom(parameter, unit);

        //        // 5. Using Reflection to call WebService.
        //        if (!result.Errors.HasErrors)
        //        {
        //            Assembly asm = result.CompiledAssembly;
        //            Type t = asm.GetType("ComponentExecutor", true, true); // If in front of adding a namespace for the proxy class, here need to be added to the front of the namespace type.

        //            object o = Activator.CreateInstance(t);
        //            MethodInfo method = t.GetMethod("Execute");
        //            string sInputPayload = xml;
        //            object strResu = method.Invoke(o, new object[] { sSSCVoucher, null, "Journal", "Query", null, sInputPayload });

        //            if (strResu != null)
        //            {
        //                return strResu.ToString();
        //            }
        //            else
        //            {
        //                return "";
        //            }
        //        }
        //        else
        //        {
        //            return "";
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.ToString(), "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //        //System.Diagnostics.EventLog.WriteEntry("Finance Tool", ex.Message + "CSL error", System.Diagnostics.EventLogEntryType.Error);
        //        ExcelAddIn4.Common.LogHelper.WriteLog(typeof(Finance_Tools), ex.Message + "Drill Down error");
        //        return "";
        //    }
        //}


        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        //public DataGridView IniXMLFormGrdForAllocationMakerUpd()
        //{
        //    DataTable dt = GetUserDataFriendlyName();
        //    DataColumn[] keys = new DataColumn[1];
        //    keys[0] = dt.Columns["SunField"];
        //    dt.PrimaryKey = keys;

        //    DataGridView d = new DataGridView();
        //    if (dt.Rows.Count > 0)
        //    {
        //        #region LineIndicator
        //        d.Columns.Add("LineIndicator", "Line Indicator");
        //        d.Columns["LineIndicator"].DataPropertyName = "LineIndicator";
        //        #endregion

        //        #region JournalNumber
        //        d.Columns.Add("JournalNumber", "JournalNumber");
        //        d.Columns["JournalNumber"].DataPropertyName = "JournalNumber";
        //        #endregion

        //        #region JournalLineNumber
        //        d.Columns.Add("JournalLineNumber", "JournalLineNumber");
        //        d.Columns["JournalLineNumber"].DataPropertyName = "JournalLineNumber";
        //        #endregion

        //        #region Ledger
        //        d.Columns.Add("Ledger", dt.Rows.Find("Ledger")["UserFriendlyName"].ToString());
        //        d.Columns["Ledger"].DataPropertyName = "Ledger";
        //        #endregion

        //        #region AccountCode
        //        d.Columns.Add("AccountCode", dt.Rows.Find("AccountCode")["UserFriendlyName"].ToString());
        //        d.Columns["AccountCode"].DataPropertyName = "AccountCode";
        //        #endregion

        //        #region AccountingPeriod
        //        d.Columns.Add("AccountingPeriod", dt.Rows.Find("AccountingPeriod")["UserFriendlyName"].ToString());
        //        d.Columns["AccountingPeriod"].DataPropertyName = "AccountingPeriod";
        //        #endregion

        //        #region TransactionDate
        //        d.Columns.Add("TransactionDate", dt.Rows.Find("TransactionDate")["UserFriendlyName"].ToString());
        //        d.Columns["TransactionDate"].DataPropertyName = "TransactionDate";
        //        #endregion

        //        #region JournalType
        //        d.Columns.Add("JournalType", dt.Rows.Find("JournalType")["UserFriendlyName"].ToString());
        //        d.Columns["JournalType"].DataPropertyName = "JournalType";
        //        #endregion

        //        #region TransactionReference
        //        d.Columns.Add("TransactionReference", dt.Rows.Find("TransactionReference")["UserFriendlyName"].ToString());
        //        d.Columns["TransactionReference"].DataPropertyName = "TransactionReference";
        //        #endregion

        //        #region AnalysisCode1
        //        d.Columns.Add("AnalysisCode1", dt.Rows.Find("AnalysisCode1")["UserFriendlyName"].ToString());
        //        d.Columns["AnalysisCode1"].DataPropertyName = "AnalysisCode1";
        //        #endregion

        //        #region AnalysisCode2
        //        d.Columns.Add("AnalysisCode2", dt.Rows.Find("AnalysisCode2")["UserFriendlyName"].ToString());
        //        d.Columns["AnalysisCode2"].DataPropertyName = "AnalysisCode2";
        //        #endregion

        //        #region AnalysisCode3
        //        d.Columns.Add("AnalysisCode3", dt.Rows.Find("AnalysisCode3")["UserFriendlyName"].ToString());
        //        d.Columns["AnalysisCode3"].DataPropertyName = "AnalysisCode3";
        //        #endregion

        //        #region AnalysisCode4
        //        d.Columns.Add("AnalysisCode4", dt.Rows.Find("AnalysisCode4")["UserFriendlyName"].ToString());
        //        d.Columns["AnalysisCode4"].DataPropertyName = "AnalysisCode4";
        //        #endregion

        //        #region AnalysisCode5
        //        d.Columns.Add("AnalysisCode5", dt.Rows.Find("AnalysisCode5")["UserFriendlyName"].ToString());
        //        d.Columns["AnalysisCode5"].DataPropertyName = "AnalysisCode5";
        //        #endregion

        //        #region AnalysisCode6
        //        d.Columns.Add("AnalysisCode6", dt.Rows.Find("AnalysisCode6")["UserFriendlyName"].ToString());
        //        d.Columns["AnalysisCode6"].DataPropertyName = "AnalysisCode6";
        //        #endregion

        //        #region AnalysisCode7
        //        d.Columns.Add("AnalysisCode7", dt.Rows.Find("AnalysisCode7")["UserFriendlyName"].ToString());
        //        d.Columns["AnalysisCode7"].DataPropertyName = "AnalysisCode7";
        //        #endregion

        //        #region AnalysisCode8
        //        d.Columns.Add("AnalysisCode8", dt.Rows.Find("AnalysisCode8")["UserFriendlyName"].ToString());
        //        d.Columns["AnalysisCode8"].DataPropertyName = "AnalysisCode8";
        //        #endregion

        //        #region AnalysisCode9
        //        d.Columns.Add("AnalysisCode9", dt.Rows.Find("AnalysisCode9")["UserFriendlyName"].ToString());
        //        d.Columns["AnalysisCode9"].DataPropertyName = "AnalysisCode9";
        //        #endregion

        //        #region AnalysisCode10
        //        d.Columns.Add("AnalysisCode10", dt.Rows.Find("AnalysisCode10")["UserFriendlyName"].ToString());
        //        d.Columns["AnalysisCode10"].DataPropertyName = "AnalysisCode10";
        //        #endregion

        //        #region AllocationMarker
        //        d.Columns.Add("AllocationMarker", dt.Rows.Find("AllocationMarker")["UserFriendlyName"].ToString());
        //        d.Columns["AllocationMarker"].DataPropertyName = "AllocationMarker";
        //        #endregion
        //    }
        //    else
        //    {
        //        d.Columns.Add("LineIndicator", "Line Indicator");
        //        d.Columns["LineIndicator"].DataPropertyName = "LineIndicator";
        //        d.Columns.Add("JournalNumber", "JournalNumber");
        //        d.Columns["JournalNumber"].DataPropertyName = "JournalNumber";
        //        d.Columns.Add("JournalLineNumber", "JournalLineNumber");
        //        d.Columns["JournalLineNumber"].DataPropertyName = "JournalLineNumber";
        //        d.Columns.Add("Ledger", "Ledger");
        //        d.Columns["Ledger"].DataPropertyName = "Ledger";
        //        d.Columns.Add("AccountCode", "Account");
        //        d.Columns["AccountCode"].DataPropertyName = "AccountCode";
        //        d.Columns.Add("AccountingPeriod", "Period");
        //        d.Columns["AccountingPeriod"].DataPropertyName = "AccountingPeriod";
        //        d.Columns.Add("TransactionDate", "TransactionDate");
        //        d.Columns["TransactionDate"].DataPropertyName = "TransactionDate";
        //        d.Columns.Add("JournalType", "Jrnl Type");
        //        d.Columns["JournalType"].DataPropertyName = "JournalType";
        //        d.Columns.Add("TransactionReference", "Trans Ref");
        //        d.Columns["TransactionReference"].DataPropertyName = "TransactionReference";
        //        d.Columns.Add("AnalysisCode1", "LA1");
        //        d.Columns["AnalysisCode1"].DataPropertyName = "AnalysisCode1";
        //        d.Columns.Add("AnalysisCode2", "LA2");
        //        d.Columns["AnalysisCode2"].DataPropertyName = "AnalysisCode2";
        //        d.Columns.Add("AnalysisCode3", "LA3");
        //        d.Columns["AnalysisCode3"].DataPropertyName = "AnalysisCode3";
        //        d.Columns.Add("AnalysisCode4", "LA4");
        //        d.Columns["AnalysisCode4"].DataPropertyName = "AnalysisCode4";
        //        d.Columns.Add("AnalysisCode5", "LA5");
        //        d.Columns["AnalysisCode5"].DataPropertyName = "AnalysisCode5";
        //        d.Columns.Add("AnalysisCode6", "LA6");
        //        d.Columns["AnalysisCode6"].DataPropertyName = "AnalysisCode6";
        //        d.Columns.Add("AnalysisCode7", "LA7");
        //        d.Columns["AnalysisCode7"].DataPropertyName = "AnalysisCode7";
        //        d.Columns.Add("AnalysisCode8", "LA8");
        //        d.Columns["AnalysisCode8"].DataPropertyName = "AnalysisCode8";
        //        d.Columns.Add("AnalysisCode9", "LA9");
        //        d.Columns["AnalysisCode9"].DataPropertyName = "AnalysisCode9";
        //        d.Columns.Add("AnalysisCode10", "LA10");
        //        d.Columns["AnalysisCode10"].DataPropertyName = "AnalysisCode10";
        //        d.Columns.Add("AllocationMarker", "Alloctn Marker");
        //        d.Columns["AllocationMarker"].DataPropertyName = "AllocationMarker";
        //    }
        //    return d;
        //}
    }
}
