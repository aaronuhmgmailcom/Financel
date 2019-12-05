/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<Finance_Tools>   
 * Author：Peter.uhm  (yanb@shinetechchina.com)
 * Modify date：2016.04
 * Modify date：2016.09
 * Version : 2.0.0.2
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Text;
using System.Configuration;
using System.IO;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml.Serialization;
using System.Windows.Forms;
using System.Xml;
using System.Runtime.InteropServices;
using System.Net;
using System.Web.Services.Description;
using System.CodeDom;
using System.CodeDom.Compiler;
using ciloci.FormulaEngine;
using System.ComponentModel.Design;
using ExcelAddIn4.Common;
namespace ExcelAddIn4
{
    public partial class Finance_Tools
    {
        /// <summary>
        /// 
        /// </summary>
        internal static string strRSDataCache
        {
            get { return AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\"; }
        }
        /// <summary>
        /// 
        /// </summary>
        internal static RSFinanceToolsEntities db
        {
            get { return new RSFinanceToolsEntities(); }
        }
        /// <summary>
        /// 
        /// </summary>
        internal static string FileIds
        {
            get { return GetFileIds(); }
            set { SaveFileIds(value.ToString()); }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static string GetFileIds()
        {
            Guid guid = new Guid(SessionInfo.UserInfo.ID);
            var FileIds = (from FT_sett in db.rsUsers
                           where FT_sett.ft_id == guid
                           select FT_sett.FileIds).First();
            if (string.IsNullOrEmpty(FileIds)) return "";
            return FileIds;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ids"></param>
        public static void SaveFileIds(string ids)
        {
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsUsers_UserInfo_FileIdsUpd", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@id", SessionInfo.UserInfo.ID));
                cmd.Parameters.Add(new SqlParameter("@FileIds", ids));
                cmd.ExecuteNonQuery();
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        internal static int totalRunTimes
        {
            get { return int.Parse(System.Configuration.ConfigurationManager.AppSettings["TotalRunTimes"]); }
            set { Finance_Tools.AppSettingSave("TotalRunTimes", value.ToString()); }
        }
        /// <summary>
        /// 
        /// </summary>
        internal static string TemplateCount
        {
            get { return System.Configuration.ConfigurationManager.AppSettings["TemplateCount"]; }
            set { Finance_Tools.AppSettingSave("TemplateCount", value.ToString()); }
        }
        /// <summary>
        /// 
        /// </summary>
        internal static int MaxColumnCount
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        internal static string simpleID
        {
            get { return System.Security.Principal.WindowsIdentity.GetCurrent().Name.Replace("\\", "").Replace("//", "").Replace("-", ""); }
        }
        /// <summary>
        /// 
        /// </summary>
        internal static string FolderFileXMLPath
        {
            get { return Finance_Tools.GetAppConfig("FolderFileXmlPath"); }
        }
        /// <summary>
        /// 
        /// </summary>
        internal static FolderFileXMLHelper ffxh
        {
            get { return new FolderFileXMLHelper(RootPath); }
        }
        /// <summary>
        /// 
        /// </summary>
        internal static string RootPath
        {
            get { return GetRootPath(); }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="foldername"></param>
        /// <returns></returns>
        public DirectoryInfo CreateFolder(string foldername)
        {
            DirectoryInfo mypath = new DirectoryInfo(foldername);
            if (mypath.Exists)
            { }
            else
            {
                mypath.Create();
                DirectorySecurity dSecurity = mypath.GetAccessControl();
                dSecurity.AddAccessRule(new FileSystemAccessRule(@"Everyone", System.Security.AccessControl.FileSystemRights.FullControl, AccessControlType.Allow));
                mypath.SetAccessControl(dSecurity);
            }
            return mypath;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        private void DeleteDirectory(string path)
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(path);
                if (dir.Exists)
                {
                    DirectoryInfo[] childs = dir.GetDirectories();
                    foreach (DirectoryInfo child in childs)
                    {
                        deleteCache(child.FullName);
                        child.Delete(true);
                    }
                    dir.Delete(true);
                }
            }
            catch (Exception ex)
            { MessageBox.Show(ex.Message); }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="strDir"></param>
        public void deleteCache(string strDir)
        {
            try
            {
                if (Directory.Exists(strDir))
                {
                    string[] strFiles = Directory.GetFiles(strDir);
                    foreach (string strFile in strFiles)
                        File.Delete(strFile);
                }
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="userID"></param>
        /// <returns></returns>
        public bool IsTemplateRenew(string userID)
        {
            try
            {
                Guid guid = new Guid(userID);
                var TemplateUpdateFlag = (from FT_sett in db.rsUsers
                                          where FT_sett.ft_id == guid
                                          select FT_sett.TemplateUpdateFlag).First();
                return (bool)TemplateUpdateFlag;
            }
            catch { return true; }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="UserID"></param>
        public void Update(string UserID)
        {
            if (IsTemplateRenew(UserID))
            {
                DeleteDirectory(Finance_Tools.RootPath);
                CreateFolder(Finance_Tools.RootPath);
                Finance_Tools.FileIds = "";
                InitializevwJournal();
                InitializeTemplatesFiles(UserID);
                uPDATETemplateUpdateFlag(UserID, false);
                Finance_Tools.TemplateCount = "-1";
            }
        }
        /// <summary>
        /// 
        /// </summary>
        private void InitializevwJournal()
        {
            try
            {
                int Templateid = GetTemplateIDByName("vwJournal");
                var data = TemplateData(Templateid);
                DirectoryInfo mypath = new DirectoryInfo(Finance_Tools.RootPath);
                string i = "";
                DetainTemplate(data, mypath, "vwJournal", ".xlsm", Templateid.ToString(), ref i);
            }
            catch { };
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="UserID"></param>
        /// <param name="flag"></param>
        public void uPDATETemplateUpdateFlag(string UserID, bool flag)
        {
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsUsers_TemplateUpdateFlag_Upd", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@TemplateUpdateFlag", flag));
                cmd.Parameters.Add(new SqlParameter("@id", UserID));
                cmd.ExecuteNonQuery();
            }
            finally
            {
                if (conn != null)
                    conn.Close();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="flag"></param>
        public void uPDATETemplateUpdateFlag(bool flag)
        {
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsUsers_TemplateUpdateFlag_UpdAll", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@TemplateUpdateFlag", flag));
                cmd.ExecuteNonQuery();
            }
            finally
            {
                if (conn != null)
                    conn.Close();
            }
        }
        /// <summary>
        ///  
        /// </summary>
        /// <param name="UserID"></param>
        public void InitializeTemplatesFiles(string UserID)
        {
            List<string> list = GetUserGroups(UserID);
            string ids = string.Empty;
            string templates = ",";
            for (int i = 0; i < list.Count; i++)
            {
                try
                {
                    string groupid = list[i];
                    bool? groupDisable = GetGroupDisableByID(int.Parse(groupid));
                    if (!(bool)groupDisable)
                    {
                        List<string> list2 = GetPermissionsByGroupID(groupid);
                        for (int j = 0; j < list2.Count; j++)
                        {
                            try
                            {
                                string permissionid = list2[j];
                                string Templateid = GetTemplateIDByPermissionID(int.Parse(permissionid));
                                string templateVisible = GetTemplateVisibleByID(UserID, Templateid);
                                string templateName = GetTemplateNameByID(int.Parse(Templateid));
                                string templateType = GetTemplateTypeByID(int.Parse(Templateid));
                                DirectoryInfo mypath = null;
                                string Folder = GetFolderByID(int.Parse(permissionid));
                                if (!string.IsNullOrEmpty(Folder))
                                    mypath = CreateFolder(RootPath + "\\" + Folder);

                                var data = TemplateData(int.Parse(Templateid));
                                if (templateVisible.Trim() == "True" && (!templates.Contains("," + templateName + ",")))
                                {
                                    DetainTemplate(data, mypath, templateName, templateType, Templateid, ref ids);
                                    templates += templateName + ",";//to calculate whether the templatename has exist in the string
                                }
                                else if (templateVisible.Trim() == "False" && (!templates.Contains("," + templateName + ",")))
                                {
                                    ids += mypath.FullName + "\\" + templateName + templateType + "-" + Templateid + ",,,";
                                    templates += templateName + ",";
                                }
                            }
                            catch { continue; }
                        }
                    }
                }
                catch { continue; }
            }
            Finance_Tools.FileIds = ids;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="userid"></param>
        /// <param name="templateid"></param>
        /// <returns></returns>
        private string GetTemplateVisibleByID(string userid, string templateid)
        {
            string dis = (from vi in db.rsUsersTemplatesVisibles
                          where vi.UserID == userid && vi.TemplateID == templateid
                          select vi.Visible).First();
            return dis;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="data"></param>
        /// <param name="mypath"></param>
        /// <param name="name"></param>
        /// <param name="type"></param>
        /// <param name="id"></param>
        /// <param name="ids"></param>
        private void DetainTemplate(Byte[] data, DirectoryInfo mypath, string name, string type, string id, ref string ids)
        {
            if (!File.Exists(mypath.FullName + "\\" + name + type))
            {
                var file = new FileStream(mypath.FullName + "\\" + name + type, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                var bw = new BinaryWriter(file);
                bw.Write(data);
                bw.Close();
                file.Close();
                ids += mypath.FullName + "\\" + name + type + "-" + id + ",,,";
            }
        }
        /// <summary>
        /// true = same ; false =different
        /// </summary>
        /// <returns></returns>
        internal static bool CompareFileCountWithConfigCount()
        {
            if (TemplateCount != ffxh.GetFilesCount() && RootPath != null)
                return false;
            else
                return true;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ids"></param>
        /// <returns></returns>
        public string getFilePath(string ids)
        {
            string[] sArray = Finance_Tools.FileIds.Split(new char[3] { ',', ',', ',' });
            try
            {
                for (int x = 0; x < sArray.Length; x++)
                    if (!string.IsNullOrEmpty(sArray[x]) && (sArray[x].Substring(sArray[x].LastIndexOf("-") + 1) == ids))
                        return sArray[x].Substring(0, sArray[x].LastIndexOf("-"));
            }
            catch
            {
                return "";
            }
            return "";
        }
        /// <summary>
        /// 
        /// </summary>
        internal static void UpdateFolderFileXMLStructure()
        {
            ffxh.InitializeFileXMLStructure();
        }
        /// <summary>
        /// Config file structure file to current user.
        /// </summary>
        public static void ConfigFileStructureFile()
        {
            if (!File.Exists(FolderFileXMLPath + simpleID + ".xml"))
            {
                FileStream aFile4 = new FileStream(FolderFileXMLPath + simpleID + ".xml", FileMode.OpenOrCreate, FileAccess.Write, FileShare.ReadWrite);
                StreamWriter sw4 = new StreamWriter(aFile4);
                sw4.Write("{0}", XMLCollection.xmlGenFolderStructure);
                sw4.Close();
                Finance_Tools.AddDirectorySecurity(FolderFileXMLPath + simpleID + ".xml", @"Everyone", System.Security.AccessControl.FileSystemRights.FullControl, System.Security.AccessControl.AccessControlType.Allow);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="DateTime1"></param>
        /// <param name="DateTime2"></param>
        /// <returns></returns>
        public static string DateDiff(DateTime DateTime1, DateTime DateTime2)
        {
            string dateDiff = null;
            try
            {
                TimeSpan ts = DateTime2 - DateTime1;
                if (ts.Days >= 1)
                    dateDiff = DateTime1.Month.ToString() + " month" + DateTime1.Day.ToString() + " day";
                else
                {
                    if (ts.Hours >= 1)
                        dateDiff = ts.Hours.ToString() + " hours " + ts.Minutes.ToString() + " minutes " + ts.Seconds.ToString() + " seconds ";
                    else
                    {
                        if (ts.Minutes >= 1)
                            dateDiff = ts.Minutes.ToString() + " minutes " + ts.Seconds.ToString() + " seconds ";
                        else
                            dateDiff = ts.Seconds.ToString() + " seconds ";
                    }
                }
            }
            catch
            { }
            return dateDiff;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="vd_type"></param>
        /// <param name="vd_prefix"></param>
        /// <param name="vd_folder"></param>
        /// <param name="vd_use_ref_as_name"></param>
        /// <param name="vd_file"></param>
        /// <param name="vd_filetype"></param>
        /// <param name="vd_macro01"></param>
        public void sprocVIEW_DOC_INS(string vd_type, string vd_prefix,
                                 string vd_folder, bool vd_use_ref_as_name,
                                 string vd_file, string vd_filetype, string vd_macro01)
        {
            SqlConnection conn = null;
            SqlDataReader rdr = null;

            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsGlobalDocumentViews_Ins", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@vd_type", vd_type));
                cmd.Parameters.Add(new SqlParameter("@vd_prefix", vd_prefix));
                cmd.Parameters.Add(new SqlParameter("@vd_folder", vd_folder));
                cmd.Parameters.Add(new SqlParameter("@vd_use_ref_as_name", vd_use_ref_as_name));
                cmd.Parameters.Add(new SqlParameter("@vd_file", vd_file));
                cmd.Parameters.Add(new SqlParameter("@vd_filetype", vd_filetype));
                cmd.Parameters.Add(new SqlParameter("@vd_macro01", vd_macro01));
                rdr = cmd.ExecuteReader();
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
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static string GetRootPath()
        {
            return "C:\\ProgramData\\RSDataV2\\RSDataTemplates" + Finance_Tools.simpleID;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string ProcessPrefix()
        {
            var prefix = (from FT_sett in db.rsTemplateTransactions
                          where FT_sett.TemplateID == SessionInfo.UserInfo.File_ftid
                          select FT_sett.Prefix).First();
            return prefix;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public int? ProcessMaxNumber()
        {
            var maxnum = (from FT_sett in db.rsTemplateTransactions
                          where FT_sett.TemplateID == SessionInfo.UserInfo.File_ftid
                          select FT_sett.maxNum).Max();
            return maxnum;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string ProcessJournalNumber()
        {
            try
            {
                var SunJournalNumber = (from FT_sett in db.rsTemplateTransactions
                                        where FT_sett.TemplateID == SessionInfo.UserInfo.File_ftid
                                        & FT_sett.maxNum == SessionInfo.UserInfo.InvNumber
                                        select FT_sett.SunJournalNumber).First();
                return SunJournalNumber;
            }
            catch
            {
                return "";
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellvalue"></param>
        /// <returns></returns>
        public string ProcessFileType(string cellvalue)
        {
            var fileType = (from FT_sett in db.rsTemplateTransactions
                            where FT_sett.TransactionName == cellvalue
                            select FT_sett.DataType).First();
            return fileType;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellvalue"></param>
        /// <returns></returns>
        public string ProcessFilePath(string cellvalue)
        {
            try
            {
                var TemplateID = (from FT_sett in db.rsTemplateTransactions
                                  where FT_sett.TransactionName == cellvalue
                                  select FT_sett.TemplateID).First();
                return TemplateID;
            }
            catch
            {
                return "";
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellvalue"></param>
        /// <returns></returns>
        public int? ProcessInvNumber(string cellvalue)
        {
            try
            {
                var maxNum = (from FT_sett in db.rsTemplateTransactions
                              where FT_sett.TransactionName == cellvalue
                              select FT_sett.maxNum).First();
                return maxNum;
            }
            catch
            {
                return 0;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="c1"></param>
        /// <param name="v1"></param>
        /// <param name="c2"></param>
        /// <param name="v2"></param>
        /// <param name="c3"></param>
        /// <param name="v3"></param>
        /// <param name="c4"></param>
        /// <param name="v4"></param>
        /// <param name="c5"></param>
        /// <param name="v5"></param>
        /// <param name="cellvalue"></param>
        /// <returns></returns>
        public bool CheckFileExist(string c1, string v1, string c2, string v2, string c3, string v3, string c4, string v4, string c5, string v5, string cellvalue)
        {
            try
            {
                var fileType = (from FT_sett in db.rsTemplateTransactions
                                where FT_sett.TemplateID == cellvalue & FT_sett.Criteria1 == c1 & FT_sett.Criteria2 == c2 & FT_sett.Criteria3 == c3 & FT_sett.Criteria4 == c4 & FT_sett.Criteria5 == c5
                            & FT_sett.Value1 == v1 & FT_sett.Value2 == v2 & FT_sett.Value3 == v3 & FT_sett.Value4 == v4 & FT_sett.Value5 == v5
                                select FT_sett.DataType).First();
                return true;
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellvalue"></param>
        /// <param name="cris"></param>
        /// <param name="vals"></param>
        /// <param name="transName"></param>
        /// <returns></returns>
        public string ReportFileType(string cellvalue, List<string> cris, List<string> vals, string transName)
        {
            string c1 = string.IsNullOrEmpty(cris[0]) ? " " : cris[0];
            string c2 = string.IsNullOrEmpty(cris[1]) ? " " : cris[1];
            string c3 = string.IsNullOrEmpty(cris[2]) ? " " : cris[2];
            string c4 = string.IsNullOrEmpty(cris[3]) ? " " : cris[3];
            string c5 = string.IsNullOrEmpty(cris[4]) ? " " : cris[4];
            string v1 = string.IsNullOrEmpty(vals[0]) ? " " : vals[0];
            string v2 = string.IsNullOrEmpty(vals[1]) ? " " : vals[1];
            string v3 = string.IsNullOrEmpty(vals[2]) ? " " : vals[2];
            string v4 = string.IsNullOrEmpty(vals[3]) ? " " : vals[3];
            string v5 = string.IsNullOrEmpty(vals[4]) ? " " : vals[4];
            var fileType = (from FT_sett in db.rsTemplateTransactions
                            where FT_sett.TemplateID == cellvalue
                         & ((FT_sett.Criteria1 == c1 & FT_sett.Value1 == v1)
                          || (FT_sett.Criteria2 == c1 & FT_sett.Value2 == v1)
                           || (FT_sett.Criteria3 == c1 & FT_sett.Value3 == v1)
                        || (FT_sett.Criteria4 == c1 & FT_sett.Value4 == v1)
                        || (FT_sett.Criteria5 == c1 & FT_sett.Value5 == v1))
                        & ((FT_sett.Criteria1 == c2 & FT_sett.Value1 == v2)
                          || (FT_sett.Criteria2 == c2 & FT_sett.Value2 == v2)
                           || (FT_sett.Criteria3 == c2 & FT_sett.Value3 == v2)
                        || (FT_sett.Criteria4 == c2 & FT_sett.Value4 == v2)
                        || (FT_sett.Criteria5 == c2 & FT_sett.Value5 == v2))
                        & ((FT_sett.Criteria1 == c3 & FT_sett.Value1 == v3)
                          || (FT_sett.Criteria2 == c3 & FT_sett.Value2 == v3)
                           || (FT_sett.Criteria3 == c3 & FT_sett.Value3 == v3)
                        || (FT_sett.Criteria4 == c3 & FT_sett.Value4 == v3)
                        || (FT_sett.Criteria5 == c3 & FT_sett.Value5 == v3))
                        & ((FT_sett.Criteria1 == c4 & FT_sett.Value1 == v4)
                          || (FT_sett.Criteria2 == c4 & FT_sett.Value2 == v4)
                           || (FT_sett.Criteria3 == c4 & FT_sett.Value3 == v4)
                        || (FT_sett.Criteria4 == c4 & FT_sett.Value4 == v4)
                        || (FT_sett.Criteria5 == c4 & FT_sett.Value5 == v4))
                        & ((FT_sett.Criteria1 == c5 & FT_sett.Value1 == v5)
                          || (FT_sett.Criteria2 == c5 & FT_sett.Value2 == v5)
                           || (FT_sett.Criteria3 == c5 & FT_sett.Value3 == v5)
                        || (FT_sett.Criteria4 == c5 & FT_sett.Value4 == v5)
                        || (FT_sett.Criteria5 == c5 & FT_sett.Value5 == v5))
                        && FT_sett.TransactionName == transName
                            select FT_sett.DataType).First();
            return fileType;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellvalue"></param>
        /// <param name="cris"></param>
        /// <param name="vals"></param>
        /// <param name="transName"></param>
        /// <returns></returns>
        public string ReportFilePath(string cellvalue, List<string> cris, List<string> vals, string transName)
        {
            string c1 = string.IsNullOrEmpty(cris[0]) ? " " : cris[0];
            string c2 = string.IsNullOrEmpty(cris[1]) ? " " : cris[1];
            string c3 = string.IsNullOrEmpty(cris[2]) ? " " : cris[2];
            string c4 = string.IsNullOrEmpty(cris[3]) ? " " : cris[3];
            string c5 = string.IsNullOrEmpty(cris[4]) ? " " : cris[4];
            string v1 = string.IsNullOrEmpty(vals[0]) ? " " : vals[0];
            string v2 = string.IsNullOrEmpty(vals[1]) ? " " : vals[1];
            string v3 = string.IsNullOrEmpty(vals[2]) ? " " : vals[2];
            string v4 = string.IsNullOrEmpty(vals[3]) ? " " : vals[3];
            string v5 = string.IsNullOrEmpty(vals[4]) ? " " : vals[4];
            var TemplatePath = (from FT_sett in db.rsTemplateTransactions
                                where FT_sett.TemplateID == cellvalue
                             & ((FT_sett.Criteria1 == c1 & FT_sett.Value1 == v1)
                              || (FT_sett.Criteria2 == c1 & FT_sett.Value2 == v1)
                               || (FT_sett.Criteria3 == c1 & FT_sett.Value3 == v1)
                            || (FT_sett.Criteria4 == c1 & FT_sett.Value4 == v1)
                            || (FT_sett.Criteria5 == c1 & FT_sett.Value5 == v1))
                            & ((FT_sett.Criteria1 == c2 & FT_sett.Value1 == v2)
                              || (FT_sett.Criteria2 == c2 & FT_sett.Value2 == v2)
                               || (FT_sett.Criteria3 == c2 & FT_sett.Value3 == v2)
                            || (FT_sett.Criteria4 == c2 & FT_sett.Value4 == v2)
                            || (FT_sett.Criteria5 == c2 & FT_sett.Value5 == v2))
                            & ((FT_sett.Criteria1 == c3 & FT_sett.Value1 == v3)
                              || (FT_sett.Criteria2 == c3 & FT_sett.Value2 == v3)
                               || (FT_sett.Criteria3 == c3 & FT_sett.Value3 == v3)
                            || (FT_sett.Criteria4 == c3 & FT_sett.Value4 == v3)
                            || (FT_sett.Criteria5 == c3 & FT_sett.Value5 == v3))
                            & ((FT_sett.Criteria1 == c4 & FT_sett.Value1 == v4)
                              || (FT_sett.Criteria2 == c4 & FT_sett.Value2 == v4)
                               || (FT_sett.Criteria3 == c4 & FT_sett.Value3 == v4)
                            || (FT_sett.Criteria4 == c4 & FT_sett.Value4 == v4)
                            || (FT_sett.Criteria5 == c4 & FT_sett.Value5 == v4))
                            & ((FT_sett.Criteria1 == c5 & FT_sett.Value1 == v5)
                              || (FT_sett.Criteria2 == c5 & FT_sett.Value2 == v5)
                               || (FT_sett.Criteria3 == c5 & FT_sett.Value3 == v5)
                            || (FT_sett.Criteria4 == c5 & FT_sett.Value4 == v5)
                            || (FT_sett.Criteria5 == c5 & FT_sett.Value5 == v5))
                            && FT_sett.TransactionName == transName
                                select FT_sett.TemplateID).First();
            return TemplatePath;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellvalue"></param>
        /// <param name="cris"></param>
        /// <param name="vals"></param>
        /// <param name="transName"></param>
        /// <returns></returns>
        public int? ReportInvNumber(string cellvalue, List<string> cris, List<string> vals, string transName)
        {
            string c1 = string.IsNullOrEmpty(cris[0]) ? " " : cris[0];
            string c2 = string.IsNullOrEmpty(cris[1]) ? " " : cris[1];
            string c3 = string.IsNullOrEmpty(cris[2]) ? " " : cris[2];
            string c4 = string.IsNullOrEmpty(cris[3]) ? " " : cris[3];
            string c5 = string.IsNullOrEmpty(cris[4]) ? " " : cris[4];
            string v1 = string.IsNullOrEmpty(vals[0]) ? " " : vals[0];
            string v2 = string.IsNullOrEmpty(vals[1]) ? " " : vals[1];
            string v3 = string.IsNullOrEmpty(vals[2]) ? " " : vals[2];
            string v4 = string.IsNullOrEmpty(vals[3]) ? " " : vals[3];
            string v5 = string.IsNullOrEmpty(vals[4]) ? " " : vals[4];
            var maxNum = (from FT_sett in db.rsTemplateTransactions
                          where FT_sett.TemplateID == cellvalue
                       & ((FT_sett.Criteria1 == c1 & FT_sett.Value1 == v1)
                        || (FT_sett.Criteria2 == c1 & FT_sett.Value2 == v1)
                         || (FT_sett.Criteria3 == c1 & FT_sett.Value3 == v1)
                      || (FT_sett.Criteria4 == c1 & FT_sett.Value4 == v1)
                      || (FT_sett.Criteria5 == c1 & FT_sett.Value5 == v1))
                      & ((FT_sett.Criteria1 == c2 & FT_sett.Value1 == v2)
                        || (FT_sett.Criteria2 == c2 & FT_sett.Value2 == v2)
                         || (FT_sett.Criteria3 == c2 & FT_sett.Value3 == v2)
                      || (FT_sett.Criteria4 == c2 & FT_sett.Value4 == v2)
                      || (FT_sett.Criteria5 == c2 & FT_sett.Value5 == v2))
                      & ((FT_sett.Criteria1 == c3 & FT_sett.Value1 == v3)
                        || (FT_sett.Criteria2 == c3 & FT_sett.Value2 == v3)
                         || (FT_sett.Criteria3 == c3 & FT_sett.Value3 == v3)
                      || (FT_sett.Criteria4 == c3 & FT_sett.Value4 == v3)
                      || (FT_sett.Criteria5 == c3 & FT_sett.Value5 == v3))
                      & ((FT_sett.Criteria1 == c4 & FT_sett.Value1 == v4)
                        || (FT_sett.Criteria2 == c4 & FT_sett.Value2 == v4)
                         || (FT_sett.Criteria3 == c4 & FT_sett.Value3 == v4)
                      || (FT_sett.Criteria4 == c4 & FT_sett.Value4 == v4)
                      || (FT_sett.Criteria5 == c4 & FT_sett.Value5 == v4))
                      & ((FT_sett.Criteria1 == c5 & FT_sett.Value1 == v5)
                        || (FT_sett.Criteria2 == c5 & FT_sett.Value2 == v5)
                         || (FT_sett.Criteria3 == c5 & FT_sett.Value3 == v5)
                      || (FT_sett.Criteria4 == c5 & FT_sett.Value4 == v5)
                      || (FT_sett.Criteria5 == c5 & FT_sett.Value5 == v5))
                      & FT_sett.TransactionName == transName
                          select FT_sett.maxNum).First();
            return maxNum;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellvalue"></param>
        /// <returns></returns>
        public Byte[] ProcessData(string cellvalue)
        {
            var data = (from FT_sett in db.rsTemplateTransactions
                        where FT_sett.TransactionName == cellvalue
                        select FT_sett.Data).First();
            return data;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateID"></param>
        /// <param name="maxNum"></param>
        /// <returns></returns>
        public string ProcessXML(string templateID, int? maxNum)
        {
            var data = (from FT_sett in db.rsTemplateTransactions
                        where FT_sett.TemplateID == templateID & FT_sett.maxNum == maxNum
                        select FT_sett.XMLData).First();
            return data;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public Byte[] TemplateData(int id)
        {
            var data = (from FT_sett in db.rsTemplates
                        where FT_sett.ID == id
                        select FT_sett.TemplateData).First();
            return data;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellvalue"></param>
        /// <param name="ID"></param>
        /// <param name="invnumber"></param>
        /// <returns></returns>
        public Byte[] ProcessPDFData(string cellvalue = "", string ID = "", int? invnumber = 0)
        {
            var data = (from FT_sett in db.rsTemplateTransactions
                        where cellvalue != "" ? FT_sett.TransactionName == cellvalue : (FT_sett.TemplateID == ID & FT_sett.maxNum == invnumber)
                        select FT_sett.PDFData).First();
            return data;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ID"></param>
        /// <returns></returns>
        public Byte[] HelpHtmlData(string ID)
        {
            var data = (from FT_sett in db.rsTemplateHelps
                        where FT_sett.TemplateID == ID
                        select FT_sett.helpFileData).First();
            return data;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ID"></param>
        /// <returns></returns>
        public string HelpFileType(string ID)
        {
            var type = (from FT_sett in db.rsTemplateHelps
                        where FT_sett.TemplateID == ID
                        select FT_sett.helpFileType).First();
            return type;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellvalue"></param>
        /// <returns></returns>
        public string ProcessTemplate(string cellvalue)
        {
            var data = (from FT_sett in db.rsTemplateTransactions
                        where FT_sett.TransactionName == cellvalue
                        select FT_sett.TemplateID).First();
            return data;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellvalue"></param>
        /// <param name="cris"></param>
        /// <param name="vals"></param>
        /// <returns></returns>
        public List<rsTemplateTransaction> ReportDataList(string cellvalue, List<string> cris, List<string> vals)
        {
            string c1 = string.IsNullOrEmpty(cris[0]) ? " " : cris[0];
            string c2 = string.IsNullOrEmpty(cris[1]) ? " " : cris[1];
            string c3 = string.IsNullOrEmpty(cris[2]) ? " " : cris[2];
            string c4 = string.IsNullOrEmpty(cris[3]) ? " " : cris[3];
            string c5 = string.IsNullOrEmpty(cris[4]) ? " " : cris[4];
            string v1 = string.IsNullOrEmpty(vals[0]) ? " " : vals[0];
            string v2 = string.IsNullOrEmpty(vals[1]) ? " " : vals[1];
            string v3 = string.IsNullOrEmpty(vals[2]) ? " " : vals[2];
            string v4 = string.IsNullOrEmpty(vals[3]) ? " " : vals[3];
            string v5 = string.IsNullOrEmpty(vals[4]) ? " " : vals[4];
            var data = (from FT_sett in db.rsTemplateTransactions
                        where FT_sett.TemplateID == cellvalue
                         & ((FT_sett.Criteria1 == c1 & FT_sett.Value1 == v1)
                          || (FT_sett.Criteria2 == c1 & FT_sett.Value2 == v1)
                           || (FT_sett.Criteria3 == c1 & FT_sett.Value3 == v1)
                        || (FT_sett.Criteria4 == c1 & FT_sett.Value4 == v1)
                        || (FT_sett.Criteria5 == c1 & FT_sett.Value5 == v1))
                        & ((FT_sett.Criteria1 == c2 & FT_sett.Value1 == v2)
                          || (FT_sett.Criteria2 == c2 & FT_sett.Value2 == v2)
                           || (FT_sett.Criteria3 == c2 & FT_sett.Value3 == v2)
                        || (FT_sett.Criteria4 == c2 & FT_sett.Value4 == v2)
                        || (FT_sett.Criteria5 == c2 & FT_sett.Value5 == v2))
                        & ((FT_sett.Criteria1 == c3 & FT_sett.Value1 == v3)
                          || (FT_sett.Criteria2 == c3 & FT_sett.Value2 == v3)
                           || (FT_sett.Criteria3 == c3 & FT_sett.Value3 == v3)
                        || (FT_sett.Criteria4 == c3 & FT_sett.Value4 == v3)
                        || (FT_sett.Criteria5 == c3 & FT_sett.Value5 == v3))
                        & ((FT_sett.Criteria1 == c4 & FT_sett.Value1 == v4)
                          || (FT_sett.Criteria2 == c4 & FT_sett.Value2 == v4)
                           || (FT_sett.Criteria3 == c4 & FT_sett.Value3 == v4)
                        || (FT_sett.Criteria4 == c4 & FT_sett.Value4 == v4)
                        || (FT_sett.Criteria5 == c4 & FT_sett.Value5 == v4))
                        & ((FT_sett.Criteria1 == c5 & FT_sett.Value1 == v5)
                          || (FT_sett.Criteria2 == c5 & FT_sett.Value2 == v5)
                           || (FT_sett.Criteria3 == c5 & FT_sett.Value3 == v5)
                        || (FT_sett.Criteria4 == c5 & FT_sett.Value4 == v5)
                        || (FT_sett.Criteria5 == c5 & FT_sett.Value5 == v5))
                        select FT_sett).ToList();
            return data;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellvalue"></param>
        /// <param name="cris"></param>
        /// <param name="vals"></param>
        /// <returns></returns>
        public int ReportDataCount(string cellvalue, List<string> cris, List<string> vals)
        {
            string c1 = string.IsNullOrEmpty(cris[0]) ? " " : cris[0];
            string c2 = string.IsNullOrEmpty(cris[1]) ? " " : cris[1];
            string c3 = string.IsNullOrEmpty(cris[2]) ? " " : cris[2];
            string c4 = string.IsNullOrEmpty(cris[3]) ? " " : cris[3];
            string c5 = string.IsNullOrEmpty(cris[4]) ? " " : cris[4];
            string v1 = string.IsNullOrEmpty(vals[0]) ? " " : vals[0];
            string v2 = string.IsNullOrEmpty(vals[1]) ? " " : vals[1];
            string v3 = string.IsNullOrEmpty(vals[2]) ? " " : vals[2];
            string v4 = string.IsNullOrEmpty(vals[3]) ? " " : vals[3];
            string v5 = string.IsNullOrEmpty(vals[4]) ? " " : vals[4];
            var data = (from FT_sett in db.rsTemplateTransactions
                        where FT_sett.TemplateID == cellvalue
                         & ((FT_sett.Criteria1 == c1 & FT_sett.Value1 == v1)
                          || (FT_sett.Criteria2 == c1 & FT_sett.Value2 == v1)
                           || (FT_sett.Criteria3 == c1 & FT_sett.Value3 == v1)
                        || (FT_sett.Criteria4 == c1 & FT_sett.Value4 == v1)
                        || (FT_sett.Criteria5 == c1 & FT_sett.Value5 == v1))
                        & ((FT_sett.Criteria1 == c2 & FT_sett.Value1 == v2)
                          || (FT_sett.Criteria2 == c2 & FT_sett.Value2 == v2)
                           || (FT_sett.Criteria3 == c2 & FT_sett.Value3 == v2)
                        || (FT_sett.Criteria4 == c2 & FT_sett.Value4 == v2)
                        || (FT_sett.Criteria5 == c2 & FT_sett.Value5 == v2))
                        & ((FT_sett.Criteria1 == c3 & FT_sett.Value1 == v3)
                          || (FT_sett.Criteria2 == c3 & FT_sett.Value2 == v3)
                           || (FT_sett.Criteria3 == c3 & FT_sett.Value3 == v3)
                        || (FT_sett.Criteria4 == c3 & FT_sett.Value4 == v3)
                        || (FT_sett.Criteria5 == c3 & FT_sett.Value5 == v3))
                        & ((FT_sett.Criteria1 == c4 & FT_sett.Value1 == v4)
                          || (FT_sett.Criteria2 == c4 & FT_sett.Value2 == v4)
                           || (FT_sett.Criteria3 == c4 & FT_sett.Value3 == v4)
                        || (FT_sett.Criteria4 == c4 & FT_sett.Value4 == v4)
                        || (FT_sett.Criteria5 == c4 & FT_sett.Value5 == v4))
                        & ((FT_sett.Criteria1 == c5 & FT_sett.Value1 == v5)
                          || (FT_sett.Criteria2 == c5 & FT_sett.Value2 == v5)
                           || (FT_sett.Criteria3 == c5 & FT_sett.Value3 == v5)
                        || (FT_sett.Criteria4 == c5 & FT_sett.Value4 == v5)
                        || (FT_sett.Criteria5 == c5 & FT_sett.Value5 == v5))
                        select FT_sett).ToList();
            return data.Count;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellvalue"></param>
        /// <param name="cris"></param>
        /// <param name="vals"></param>
        /// <param name="transName"></param>
        /// <returns></returns>
        public Byte[] ReportData(string cellvalue, List<string> cris, List<string> vals, string transName)
        {
            string c1 = string.IsNullOrEmpty(cris[0]) ? " " : cris[0];
            string c2 = string.IsNullOrEmpty(cris[1]) ? " " : cris[1];
            string c3 = string.IsNullOrEmpty(cris[2]) ? " " : cris[2];
            string c4 = string.IsNullOrEmpty(cris[3]) ? " " : cris[3];
            string c5 = string.IsNullOrEmpty(cris[4]) ? " " : cris[4];
            string v1 = string.IsNullOrEmpty(vals[0]) ? " " : vals[0];
            string v2 = string.IsNullOrEmpty(vals[1]) ? " " : vals[1];
            string v3 = string.IsNullOrEmpty(vals[2]) ? " " : vals[2];
            string v4 = string.IsNullOrEmpty(vals[3]) ? " " : vals[3];
            string v5 = string.IsNullOrEmpty(vals[4]) ? " " : vals[4];
            var data = (from FT_sett in db.rsTemplateTransactions
                        where FT_sett.TemplateID == cellvalue
                         & ((FT_sett.Criteria1 == c1 & FT_sett.Value1 == v1)
                          || (FT_sett.Criteria2 == c1 & FT_sett.Value2 == v1)
                           || (FT_sett.Criteria3 == c1 & FT_sett.Value3 == v1)
                        || (FT_sett.Criteria4 == c1 & FT_sett.Value4 == v1)
                        || (FT_sett.Criteria5 == c1 & FT_sett.Value5 == v1))
                        & ((FT_sett.Criteria1 == c2 & FT_sett.Value1 == v2)
                          || (FT_sett.Criteria2 == c2 & FT_sett.Value2 == v2)
                           || (FT_sett.Criteria3 == c2 & FT_sett.Value3 == v2)
                        || (FT_sett.Criteria4 == c2 & FT_sett.Value4 == v2)
                        || (FT_sett.Criteria5 == c2 & FT_sett.Value5 == v2))
                        & ((FT_sett.Criteria1 == c3 & FT_sett.Value1 == v3)
                          || (FT_sett.Criteria2 == c3 & FT_sett.Value2 == v3)
                           || (FT_sett.Criteria3 == c3 & FT_sett.Value3 == v3)
                        || (FT_sett.Criteria4 == c3 & FT_sett.Value4 == v3)
                        || (FT_sett.Criteria5 == c3 & FT_sett.Value5 == v3))
                        & ((FT_sett.Criteria1 == c4 & FT_sett.Value1 == v4)
                          || (FT_sett.Criteria2 == c4 & FT_sett.Value2 == v4)
                           || (FT_sett.Criteria3 == c4 & FT_sett.Value3 == v4)
                        || (FT_sett.Criteria4 == c4 & FT_sett.Value4 == v4)
                        || (FT_sett.Criteria5 == c4 & FT_sett.Value5 == v4))
                        & ((FT_sett.Criteria1 == c5 & FT_sett.Value1 == v5)
                          || (FT_sett.Criteria2 == c5 & FT_sett.Value2 == v5)
                           || (FT_sett.Criteria3 == c5 & FT_sett.Value3 == v5)
                        || (FT_sett.Criteria4 == c5 & FT_sett.Value4 == v5)
                        || (FT_sett.Criteria5 == c5 & FT_sett.Value5 == v5))
                        & FT_sett.TransactionName == transName
                        select FT_sett.Data).First();
            return data;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public Byte[] ReportData(string id)
        {
            Guid g = new Guid(id);
            var data = (from FT_sett in db.rsTemplateTransactions
                        where FT_sett.GUID == g
                        select FT_sett.Data).First();
            return data;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public string ReportFileType(string id)
        {
            Guid g = new Guid(id);
            var fileType = (from FT_sett in db.rsTemplateTransactions
                            where FT_sett.GUID == g
                            select FT_sett.DataType).First();
            return fileType;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public Array vdPrefix()
        {
            var prefixes = from VIEW_DOC in db.rsGlobalDocumentViews
                           select VIEW_DOC.prefix;
            return prefixes.ToArray();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="prefix"></param>
        /// <returns></returns>
        public string vdFilepath(string prefix)
        {
            var folder = (from VIEW_DOC in db.rsGlobalDocumentViews
                          where VIEW_DOC.prefix == prefix
                          select VIEW_DOC.folder).First();
            return folder.ToString();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="prefix"></param>
        /// <returns></returns>
        public bool vdUseFile(string prefix)
        {
            var usefile = (from VIEW_DOC in db.rsGlobalDocumentViews
                           where VIEW_DOC.prefix == prefix
                           select VIEW_DOC.useRefAsName).First();
            return (bool)usefile;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="prefix"></param>
        /// <returns></returns>
        public string vdFilename(string prefix)
        {
            var file = (from VIEW_DOC in db.rsGlobalDocumentViews
                        where VIEW_DOC.prefix == prefix
                        select VIEW_DOC.file).First();
            return file.ToString();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="prefix"></param>
        /// <returns></returns>
        public string vdFiletype(string prefix)
        {
            var ftype = (from VIEW_DOC in db.rsGlobalDocumentViews
                         where VIEW_DOC.prefix == prefix
                         select VIEW_DOC.type).First();
            return ftype.ToString();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="prefix"></param>
        /// <returns></returns>
        public string vdMacro01(string prefix)
        {
            var mac01 = (from VIEW_DOC in db.rsGlobalDocumentViews
                         where VIEW_DOC.prefix == prefix
                         select VIEW_DOC.macro01).First();
            return mac01.ToString();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="strKey"></param>
        /// <returns></returns>
        public static string GetAppConfig(string strKey)
        {
            foreach (string key in ConfigurationManager.AppSettings)
                if (key == strKey)
                    return ConfigurationManager.AppSettings[strKey];

            return null;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateID"></param>
        /// <returns></returns>
        public bool IsOutPutPaneVisiable(string templateID)
        {
            try
            {
                string v = (from FT_sett in db.rsTemplateVisibles
                            where FT_sett.TemplateID == templateID
                   & FT_sett.UserID == SessionInfo.UserInfo.ID
                            select FT_sett.OutputPaneVisiable).First();

                if (v == "1") return true;
                else return false;
            }
            catch { return false; }
        }
        /// <summary>
        /// Folder Security Principal
        /// </summary>
        /// <param name="d"></param>
        /// <returns></returns>
        public static bool IsAddedToGallery(DirectoryInfo d)
        {
            try
            {
                bool returnValue = false;
                DirectorySecurity dSecurity = d.GetAccessControl();
                WindowsPrincipal winPrincipal = new WindowsPrincipal(WindowsIdentity.GetCurrent());
                AuthorizationRuleCollection rules = dSecurity.GetAccessRules(true, true, typeof(NTAccount));

                //WindowsIdentity ii = WindowsIdentity.GetCurrent();
                //SecurityIdentifier sid = ii.User;
                //NTAccount ntacc = (NTAccount)sid.Translate(typeof(NTAccount));
                //string userrights = @"RW";//Permission string, their own definitions

                foreach (FileSystemAccessRule ar in rules)
                {
                    if (winPrincipal.IsInRole(ar.IdentityReference.Value))
                    {
                        if ((ar.FileSystemRights & FileSystemRights.Read) != 0)
                            returnValue = true;
                    }
                }
                return returnValue;
            }
            catch
            {
                return false;
            }
        }

        // Adds an ACL entry on the specified directory for the specified account.
        public static void AddDirectorySecurity(string FileName, string Account, FileSystemRights Rights, AccessControlType ControlType)
        {
            // Create a new DirectoryInfo object.
            DirectoryInfo dInfo = new DirectoryInfo(FileName);

            // Get a DirectorySecurity object that represents the 
            // current security settings.
            DirectorySecurity dSecurity = dInfo.GetAccessControl();

            // Add the FileSystemAccessRule to the security settings. 
            dSecurity.AddAccessRule(new FileSystemAccessRule(Account,
                                                            Rights,
                                                            ControlType));

            // Set the new access settings.
            dInfo.SetAccessControl(dSecurity);
        }

        // Removes an ACL entry on the specified directory for the specified account.
        public static void RemoveDirectorySecurity(string FileName, string Account, FileSystemRights Rights, AccessControlType ControlType)
        {
            DirectoryInfo dInfo = new DirectoryInfo(FileName);
            DirectorySecurity dSecurity = dInfo.GetAccessControl();
            dSecurity.RemoveAccessRule(new FileSystemAccessRule(Account,
                                                            Rights,
                                                            ControlType));
            dInfo.SetAccessControl(dSecurity);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public static string RemoveNumber(string key)
        {
            return Regex.Replace(key, @"\d", "");
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public static string RemoveNotNumber(string key)
        {
            return Regex.Replace(key, @"[^\d]*", "");
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static bool IsLetterOrDigit(string str)
        {
            if (str == null || str.Length == 0)
                return false;
            foreach (char c in str)
                if (!Char.IsLetterOrDigit(c))
                    return false;

            return true;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public byte[] GetData(string fileName)
        {
            try
            {
                FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                byte[] bytes = new byte[fs.Length];
                fs.Read(bytes, 0, (int)fs.Length);
                fs.Close();
                return bytes;
            }
            catch
            {
                return new byte[0];
            }
        }
        /// <summary>
        /// Get User FriendlyName
        /// </summary>
        /// <returns></returns>
        public DataTable GetUserDataFriendlyName()
        {
            var all = (from Reports in db.rsGlobalFields
                       orderby Reports.version ascending
                       select Reports).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// Get Create text file FriendlyName
        /// </summary>
        /// <returns></returns>
        public DataTable GetCreateTextFileHeader()
        {
            var all = (from createtextfile in db.rsTemplateXMLTextFileDGVs
                       where createtextfile.TemplateID == SessionInfo.UserInfo.File_ftid
                       orderby createtextfile.ft_id ascending
                       select createtextfile).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// Get User GenFriendlyName
        /// </summary>
        /// <returns></returns>
        public DataTable GetUserDataGenFriendlyName()
        {
            var all = (from Reports in db.rsTemplateGenDescFields
                       where Reports.TemplateID == SessionInfo.UserInfo.File_ftid
                       orderby Reports.version ascending
                       select Reports).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// Get User Reports
        /// </summary>
        /// <returns></returns>
        public DataTable GetUserReports()
        {
            var all = (from Reports in db.rsTemplateTransactions
                       where
                          ((Reports.Criteria1 != null & Reports.Criteria1 != "")
                          || (Reports.Criteria2 != null & Reports.Criteria2 != "")
                           || (Reports.Criteria3 != null & Reports.Criteria3 != "")
                        || (Reports.Criteria4 != null & Reports.Criteria4 != "")
                        || (Reports.Criteria5 != null & Reports.Criteria5 != ""))
                       select Reports).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// Get User Reports
        /// </summary>
        /// <returns></returns>
        public DataTable GetReportsViaTemplatePath()
        {
            var all = (from Reports in db.rsTemplateTransactions
                       where
                          ((Reports.Criteria1 != null & Reports.Criteria1 != "")
                          || (Reports.Criteria2 != null & Reports.Criteria2 != "")
                           || (Reports.Criteria3 != null & Reports.Criteria3 != "")
                        || (Reports.Criteria4 != null & Reports.Criteria4 != "")
                        || (Reports.Criteria5 != null & Reports.Criteria5 != "")) & Reports.TemplateID == SessionInfo.UserInfo.File_ftid
                       select Reports).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="type"></param>
        /// <param name="reference"></param>
        /// <returns></returns>
        public DataTable GetLineDetailDataFromDB(string type, string reference)
        {
            var all = (from OutPutProfile in db.rsTemplateJournals
                       where OutPutProfile.TemplateID == SessionInfo.UserInfo.File_ftid
                       & OutPutProfile.Type == type & OutPutProfile.Reference.Trim() == reference
                       orderby OutPutProfile.Reference ascending
                       select OutPutProfile).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public DataTable GetAllJournalRefOfTemplate()
        {
            var all = (from OutPutProfile in db.rsTemplateJournals
                       where OutPutProfile.TemplateID == SessionInfo.UserInfo.File_ftid
                       group OutPutProfile by new { t2 = OutPutProfile.Reference } into g
                       orderby g.Key.t2 ascending
                       select new
                       {
                           references = g.Key.t2
                       }).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public DataTable GetGroupsFromDB()
        {
            var all = (from Groups in db.rsGroups
                       select Groups).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ID"></param>
        /// <returns></returns>
        public bool? GetGroupDisableByID(int ID)
        {
            try
            {
                bool? dis = (from Groups in db.rsGroups
                             where Groups.ID == ID
                             select Groups.GroupDisable).First();
                return dis;
            }
            catch { return true; }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public DataTable GetUsersFromDB()
        {
            var all = (from Users in db.rsUsers
                       select Users).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="groupid"></param>
        /// <returns></returns>
        public DataTable GetGroupPermissionsView(string groupid)
        {
            var all = (from Permissions in db.View_GroupPermissions
                       where Permissions.GroupID == groupid
                       orderby Permissions.ID ascending
                       select Permissions).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public DataTable GetPermissionsFromDB()
        {
            var all = (from Permissions in db.View_TemplatesPermissions
                       orderby Permissions.ID ascending
                       select Permissions).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="groupid"></param>
        /// <returns></returns>
        public List<string> GetUsersByGroupID(string groupid)
        {
            var all = (from GroupUsers in db.rsUserGroups
                       where GroupUsers.GroupID == groupid
                       select GroupUsers.UserID).ToList();
            return all;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="userid"></param>
        /// <param name="templateid"></param>
        /// <returns></returns>
        public string GetVisibleByUserIDTemplateID(string userid, string templateid)
        {
            try
            {
                var all = (from ut in db.rsUsersTemplatesVisibles
                           where ut.UserID == userid && ut.TemplateID == templateid
                           select ut.Visible).First();
                return all;
            }
            catch { return ""; }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="groupid"></param>
        /// <returns></returns>
        public List<string> GetPermissionsByGroupID(string groupid)
        {
            var all = (from GroupPermissions in db.rsGroupPermissions
                       where GroupPermissions.GroupID == groupid
                       select GroupPermissions.PermissionID).ToList();
            return all;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="actionid"></param>
        /// <param name="templateid"></param>
        /// <returns></returns>
        public List<string> GetGroupsByActionID(string actionid, string templateid)
        {
            var all = (from GroupActions in db.View_GroupActions
                       where GroupActions.ActionID == actionid && GroupActions.TemplateID == templateid
                       select GroupActions.GroupID).ToList();
            return all;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="processName"></param>
        /// <returns></returns>
        public DataTable GetCreateTextFileDataFromDB(string processName)
        {
            var all = (from OutPutCTF in db.rsTemplateXMLTextFileDGVs
                       where OutPutCTF.TemplateID == SessionInfo.UserInfo.File_ftid
                       && OutPutCTF.ProcessName == processName
                       orderby OutPutCTF.ReferenceNumber ascending
                       select OutPutCTF).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="comName"></param>
        /// <param name="methodName"></param>
        /// <returns></returns>
        public DataTable GetCreateTextFileDataFromDB(string comName, string methodName)
        {
            var all = (from OutPutCTF in db.rsTemplateXMLTextFileDGVs
                       where OutPutCTF.TemplateID == SessionInfo.UserInfo.File_ftid
                       && OutPutCTF.SunComponent == comName && OutPutCTF.SunMethod == methodName
                       orderby OutPutCTF.ReferenceNumber ascending
                       select OutPutCTF).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="comName"></param>
        /// <param name="methodName"></param>
        /// <param name="field"></param>
        /// <param name="friendName"></param>
        /// <returns></returns>
        public string GetSectionFromDB(string comName, string methodName, string field, string friendName)
        {
            try
            {
                var all = (from OutPutCTF in db.rsTemplateCreateXMLTextProfiles
                           where OutPutCTF.SunComponentName == comName && OutPutCTF.SunMethod == methodName
                          && OutPutCTF.Field == field && OutPutCTF.FriendlyName == friendName
                           select OutPutCTF.Section).First();
                return all;
            }
            catch
            {
                return "";
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public DataTable GetXMLorTextFileNames(int type)
        {
            var all = (from XMLorTextFile in db.rsTemplateXMLTEXTFiles
                       where XMLorTextFile.FileType == type
                       select XMLorTextFile).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="comName"></param>
        /// <param name="methodName"></param>
        /// <returns></returns>
        public DataTable GetXMLorTextFileFieldsByComName(string comName, string methodName)
        {
            var all = (from XMLorTextFile in db.rsTemplateCreateXMLTextProfiles
                       where XMLorTextFile.SunComponentName == comName && XMLorTextFile.SunMethod == methodName
                       select XMLorTextFile).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="comName"></param>
        /// <param name="methodName"></param>
        /// <returns></returns>
        public DataTable GetXMLorTextFileVFieldsByComName(string comName, string methodName)
        {
            var all = (from XMLorTextFile in db.rsTemplateCreateXMLTextProfiles
                       where XMLorTextFile.SunComponentName == comName && XMLorTextFile.SunMethod == methodName
                       && XMLorTextFile.Visible == true
                       select XMLorTextFile).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="comName"></param>
        /// <param name="methodName"></param>
        /// <returns></returns>
        public DataTable GetXMLorTextFileHeaderByComName(string comName, string methodName)
        {
            var all = (from XMLorTextFile in db.rsTemplateCreateXMLTextProfiles
                       where XMLorTextFile.SunComponentName == comName && XMLorTextFile.SunMethod == methodName
                       && XMLorTextFile.Section == "Header" && XMLorTextFile.Visible == true
                       select XMLorTextFile).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="comName"></param>
        /// <param name="methodName"></param>
        /// <returns></returns>
        public DataTable GetXMLorTextFileLineByComName(string comName, string methodName)
        {
            var all = (from XMLorTextFile in db.rsTemplateCreateXMLTextProfiles
                       where XMLorTextFile.SunComponentName == comName && XMLorTextFile.SunMethod == methodName
                       && XMLorTextFile.Section == "Line" && XMLorTextFile.Visible == true
                       select XMLorTextFile).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="comName"></param>
        /// <param name="methodName"></param>
        /// <returns></returns>
        public string GetFirstLineByComName(string comName, string methodName)
        {
            string all = (from XMLorTextFile in db.rsTemplateCreateXMLTextProfiles
                          where XMLorTextFile.SunComponentName == comName && XMLorTextFile.SunMethod == methodName
                          && XMLorTextFile.Section == "Line"
                          orderby XMLorTextFile.ID ascending
                          select XMLorTextFile.Field).First();
            return all;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="comName"></param>
        /// <param name="methodName"></param>
        /// <param name="field"></param>
        /// <returns></returns>
        public string GetParentNode(string comName, string methodName, string field)
        {
            var str = (from XMLorTextFile in db.rsTemplateCreateXMLTextProfiles
                       where XMLorTextFile.SunComponentName == comName && XMLorTextFile.SunMethod == methodName
                       && XMLorTextFile.Field == field
                       select XMLorTextFile.Parent).First();
            return str;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="comName"></param>
        /// <param name="methodName"></param>
        /// <param name="field"></param>
        /// <returns></returns>
        public DataTable GetSonNode(string comName, string methodName, string field)
        {
            var all = (from XMLorTextFile in db.rsTemplateCreateXMLTextProfiles
                       where XMLorTextFile.SunComponentName == comName && XMLorTextFile.SunMethod == methodName
                       && XMLorTextFile.Parent == field
                       select XMLorTextFile).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public DataTable GetXMLorTextFileFieldsByFileName(string fileName)
        {
            var all = (from XMLorTextFile in db.rsTemplateCreateXMLTextProfiles
                       where XMLorTextFile.TextFileName == fileName
                       select XMLorTextFile).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateID"></param>
        /// <param name="name"></param>
        /// <param name="refn"></param>
        /// <returns></returns>
        public string GetProcessIDFromDB(string templateID, string name, string refn)
        {
            string all = string.Empty;
            if (name == "Journal Post")
            {
                all = (from TextFile in db.rsTemplateJournals
                       where (TextFile.TemplateID == templateID) && (TextFile.Type == "1") && (TextFile.Reference == refn)
                       select TextFile.ft_id).First().ToString();
            }
            else if (name == "Journal Update")
            {
                all = (from TextFile in db.rsTemplateJournals
                       where (TextFile.TemplateID == templateID) && (TextFile.Type == "2") && (TextFile.Reference == refn)
                       select TextFile.ft_id).First().ToString();
            }
            else
            {
                all = (from TextFile in db.rsTemplateXMLTextFileDGVs
                       where (TextFile.TemplateID == templateID) && (TextFile.ReferenceNumber == refn) && ((TextFile.ProcessName + TextFile.SunComponent + "." + TextFile.SunMethod) == name)
                       select TextFile.ft_id).First().ToString();
            }
            return all;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateID"></param>
        /// <param name="buttonName"></param>
        /// <returns></returns>
        public bool ButtonNameExist(string templateID, string buttonName)
        {
            try
            {
                int all = (from TextFile in db.rsTemplateActions
                           where TextFile.TemplateID == templateID && TextFile.ButtonName == buttonName
                           select TextFile.ID).First();
                if (all > 0)
                    return true;
                else
                    return false;
            }
            catch { return false; }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateID"></param>
        /// <returns></returns>
        public DataTable GetProcessesFromDB(string templateID)
        {
            var all = (from TextFile in db.rsTemplateXMLTextFileDGVs
                       where TextFile.TemplateID == templateID
                       group TextFile by new { t1 = TextFile.ProcessName + TextFile.SunComponent + "." + TextFile.SunMethod, t2 = TextFile.TemplateID } into g
                       select new
                       {
                           id = g.Key.t2,
                           name = g.Key.t1
                       }).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateID"></param>
        /// <returns></returns>
        public DataTable GetSaveRefFromDB(string templateID)
        {
            var all = (from TextFile in db.rsTemplateSettings
                       where TextFile.TemplateID == templateID
                       select TextFile).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateID"></param>
        /// <param name="processName"></param>
        /// <returns></returns>
        public DataTable GetProcessesRefFromDB(string templateID, string processName)
        {
            if (processName == "Journal Post")
            {
                var all = (from TextFile in db.rsTemplateJournals
                           where TextFile.TemplateID == templateID && (TextFile.Type == "1" || TextFile.Type == "0")
                           group TextFile by new { t2 = TextFile.TemplateID, t3 = TextFile.Reference } into g
                           select new
                           {
                               reference = g.Key.t3
                           }).ToList();
                return ToDataTable(all);
            }
            else if (processName == "Journal Update")
            {
                var all = (from TextFile in db.rsTemplateJournals
                           where TextFile.TemplateID == templateID && TextFile.Type == "2"
                           group TextFile by new { t2 = TextFile.TemplateID, t3 = TextFile.Reference } into g
                           select new
                           {
                               reference = g.Key.t3
                           }).ToList();
                return ToDataTable(all);
            }
            else
            {
                var all = (from TextFile in db.rsTemplateXMLTextFileDGVs
                           where TextFile.TemplateID == templateID && (TextFile.ProcessName + TextFile.SunComponent + "." + TextFile.SunMethod) == processName
                           group TextFile by new { t1 = TextFile.ProcessName + TextFile.SunComponent + "." + TextFile.SunMethod, t2 = TextFile.TemplateID, t3 = TextFile.ReferenceNumber } into g
                           select new
                           {
                               reference = g.Key.t3
                           }).ToList();
                return ToDataTable(all);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ID"></param>
        /// <returns></returns>
        public DataTable GetTemplateButtons(string ID)
        {
            var all = (from buttons in db.rsTemplateActions
                       where buttons.TemplateID == ID
                       group buttons by new { t1 = buttons.ButtonName, t2 = buttons.ButtonGroup, t3 = buttons.GroupOrder } into g
                       orderby g.Key.t3 ascending
                       select new
                       {
                           name = g.Key.t1,
                           buttonGroup = g.Key.t2,
                           GroupOrder = g.Key.t3
                       }
                       ).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateid"></param>
        /// <param name="buttonName"></param>
        /// <returns></returns>
        public DataTable GetProcessMacroFromDB(string templateid, string buttonName)
        {
            var all = (from ProcessesMacros in db.rsTemplateActions
                       where ProcessesMacros.TemplateID == templateid && ProcessesMacros.ButtonName == buttonName
                       orderby ProcessesMacros.ProcessMacroOrder ascending
                       select ProcessesMacros).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateid"></param>
        /// <param name="reference"></param>
        /// <returns></returns>
        public DataTable GetTemplateActionByRef(string templateid, string reference)
        {
            var all = (from ProcessesMacros in db.rsTemplateActions
                       where ProcessesMacros.TemplateID == templateid && ProcessesMacros.Reference.Trim() == reference && ProcessesMacros.Type == 1
                       select ProcessesMacros).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateid"></param>
        /// <param name="groupname"></param>
        /// <returns></returns>
        public DataTable GetGroupViaGroupName(string templateid, string groupname)
        {
            var all = (from ProcessesMacros in db.rsTemplateActions
                       where ProcessesMacros.TemplateID == templateid && ProcessesMacros.ButtonGroup == groupname
                       select ProcessesMacros).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public List<string> GetTextFiles()
        {
            var all = (from files in db.rsTemplateXMLTEXTFiles
                       where files.FileType == 0
                       select files.RelatedName).ToList();
            return all;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public bool FileExist(string filename)
        {
            try
            {
                var fileid = (from files in db.rsTemplateXMLTEXTFiles
                              where files.RelatedName == filename
                              select files.ID).First();
                if (fileid > 0) return true;
                else return false;
            }
            catch { return false; }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public List<string> GetXMLFiles()
        {
            var all = (from files in db.rsTemplateXMLTEXTFiles
                       where files.FileType == 1
                       select files.RelatedName).ToList();
            return all;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="com"></param>
        /// <param name="method"></param>
        /// <returns></returns>
        public DataTable GetXMLData(string com, string method)
        {
            var all = (from files in db.rsTemplateCreateXMLTextProfiles
                       where files.SunComponentName == com && files.SunMethod == method
                       select files).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public DataTable GetFileData(string filename)
        {
            var all = (from files in db.rsTemplateCreateXMLTextProfiles
                       where files.TextFileName == filename
                       select files).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public string GetXMLFileContent(string filename)
        {
            var all = (from files in db.rsTemplateXMLTEXTFiles
                       where files.RelatedName == filename
                       select files.FileContent).First();
            return all;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public string GetXMLFileTemplate(string filename)
        {
            var all = (from files in db.rsTemplateXMLTEXTFiles
                       where files.RelatedName == filename
                       select files.XMLTemplate).First();
            return all;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public string GetFileSeparator(string filename)
        {
            var all = (from files in db.rsTemplateCreateXMLTextProfiles
                       where files.TextFileName == filename
                       select files.Separator).First();
            return all;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ID"></param>
        /// <returns></returns>
        public List<string> GetUserGroups(string ID)
        {
            var all = (from groups in db.rsUserGroups
                       where (groups.UserID == ID) && (groups.GroupID != "")
                       select groups.GroupID).ToList();
            return all;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public bool GetTemplateByPath(string path)
        {
            try
            {
                int templateID = (from template in db.rsTemplates
                                  where template.OriginTemplatePath == path
                                  select template.ID).First();
                if (templateID > 0) return true;
                else return false;
            }
            catch { return false; }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        /// <param name="fileType"></param>
        /// <returns></returns>
        public bool GetTemplateByNameAndType(string name, string fileType)
        {
            try
            {
                int templateID = (from template in db.rsTemplates
                                  where template.TemplateName == name && template.FileType == fileType
                                  select template.ID).First();
                if (templateID > 0) return true;
                else return false;
            }
            catch { return false; }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public bool GetPermissionByName(string name)
        {
            try
            {
                int id = (from p in db.rsPermissions
                          where p.PermissionName == (name + " - Global")
                          select p.ID).First();
                if (id > 0) return true;
                else return false;
            }
            catch { return false; }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ID"></param>
        /// <returns></returns>
        public string GetTemplateNameByID(int ID)
        {
            string name = (from p in db.rsTemplates
                           where p.ID == ID
                           select p.TemplateName).First();
            return name;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ID"></param>
        /// <returns></returns>
        public string GetTemplatePathByID(int ID)
        {
            string path = (from p in db.rsTemplates
                           where p.ID == ID
                           select p.OriginTemplatePath).First();
            return path;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ID"></param>
        /// <returns></returns>
        public string GetTemplateTypeByID(int ID)
        {
            string type = (from p in db.rsTemplates
                           where p.ID == ID
                           select p.FileType).First();
            return type;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ID"></param>
        /// <returns></returns>
        public string GetTemplateIDByPermissionID(int ID)
        {
            string templateID = (from p in db.rsPermissions
                                 where p.ID == ID
                                 select p.TemplateID).First();
            return templateID;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public bool IsGUID(string str)
        {
            Match m = Regex.Match(str, @"^[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}$", RegexOptions.IgnoreCase);
            if (m.Success)
                return true;
            else
                return false;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Name"></param>
        /// <returns></returns>
        public int GetTemplateIDByName(string Name)
        {
            try
            {
                int templateID = (from p in db.rsTemplates
                                  where p.TemplateName == Name
                                  select p.ID).First();
                return templateID;
            }
            catch { return -1; }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ID"></param>
        /// <returns></returns>
        public string GetPermissionRemarkByID(int ID)
        {
            string remark = (from p in db.rsPermissions
                             where p.ID == ID
                             select p.remark).First();
            return remark;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ID"></param>
        /// <returns></returns>
        public string GetFolderByID(int ID)
        {
            string folder = (from p in db.rsPermissions
                             where p.ID == ID
                             select p.Folder).First();
            return folder;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="groupid"></param>
        /// <returns></returns>
        public DataTable GetGroupButtonsView(string groupid)
        {
            var all = (from buttons in db.View_GroupActions
                       where buttons.GroupID == groupid
                       && buttons.ButtonGroup != ""
                       orderby buttons.GroupOrder ascending, buttons.ButtonOrder ascending
                       select buttons).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ID"></param>
        /// <param name="templateID"></param>
        /// <returns></returns>
        public DataTable GetGroupButtonsView(string ID, string templateID)
        {
            var all = (from buttons in db.View_GroupActions
                       where buttons.TemplateID == templateID && buttons.GroupID == ID
                       && buttons.ButtonGroup != ""
                       orderby buttons.GroupOrder ascending, buttons.ButtonOrder
                       select buttons).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public string GetButtonGroup(int id)
        {
            string Group = (from buttons in db.rsTemplateActions
                            where buttons.ID == id
                            select buttons.ButtonGroup).First();
            return Group;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="groupname"></param>
        /// <returns></returns>
        public DataTable GetGroupButtons(string groupname)
        {
            var buttonlist = (from buttons in db.rsTemplateActions
                              where buttons.ButtonGroup == groupname
                              group buttons by new { t1 = buttons.ButtonOrder, t2 = buttons.ButtonGroup, t3 = buttons.ButtonName } into g
                              orderby g.Key.t1 ascending
                              select new
                              {
                                  order = g.Key.t1,
                                  buttonGroup = g.Key.t2,
                                  buttonName = g.Key.t3
                              }).ToList();
            return ToDataTable(buttonlist);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public DataTable GetPDFRelation()
        {
            var all = (from FT_user in db.rsTemplateContainers
                       select FT_user).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateID"></param>
        /// <returns></returns>
        public DataTable GetReportCriteria(string templateID)
        {
            var all = (from FT_user in db.rsTemplateSettings
                       where FT_user.TemplateID == templateID
                       select FT_user).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateID"></param>
        /// <param name="reference"></param>
        /// <returns></returns>
        public DataTable GetReportCriteriaByRef(string templateID, string reference)
        {
            var all = (from FT_user in db.rsTemplateSettings
                       where FT_user.TemplateID == templateID && FT_user.Reference == reference
                       select FT_user).ToList();
            return ToDataTable(all);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateID"></param>
        /// <returns></returns>
        public bool? GetReportOpenTransUponSaveChk(string templateID)
        {
            bool? all = (from FT_user in db.rsTemplateSettings
                         where FT_user.TemplateID == templateID
                         select FT_user.OpenTransUponSave).First();
            return all;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileID"></param>
        /// <param name="pdfpath"></param>
        /// <returns></returns>
        public string GetPDFRelationExist(string fileID, string pdfpath)
        {
            var str = (from FT_user in db.rsTemplateContainers
                       where FT_user.TemplateID == fileID
                         & FT_user.ft_relatefilepath == pdfpath
                       select FT_user).First();
            return str.ToString();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileID"></param>
        /// <returns></returns>
        public string GetPDFViaTemplatePath(string fileID)
        {
            var str = (from FT_user in db.rsTemplateContainers
                       where FT_user.TemplateID == fileID
                       select FT_user.ft_relatefilepath).First();
            return str.ToString();
        }
        /// <summary>
        /// 
        /// </summary>
        public void ReadPDFOfCurrentTransaction()
        {
            string IDvalue = string.Empty;
            int InvNumber = 0;
            if (SessionInfo.UserInfo.Dictionary.dict.Count != 0 && SessionInfo.UserInfo.Dictionary.dict.ContainsKey(SessionInfo.UserInfo.CachePath))
            {
                IDvalue = SessionInfo.UserInfo.Dictionary.dict[SessionInfo.UserInfo.CachePath];
                string[] sArray = Regex.Split(IDvalue, ",");
                IDvalue = sArray[0];
                int.TryParse(sArray[1], out InvNumber);
            }
            if (!(!IsGUID(Path.GetFileNameWithoutExtension(SessionInfo.UserInfo.CachePath)) || string.IsNullOrEmpty(IDvalue)))
            {
                ReadPDFinDB(ID: IDvalue, invnumber: InvNumber);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string OpenHelpFile()
        {
            try
            {
                var data = HelpHtmlData(SessionInfo.UserInfo.File_ftid);
                var fileType = HelpFileType(SessionInfo.UserInfo.File_ftid);
                string tmp = Guid.NewGuid().ToString();
                var file = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType, FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
                var bw = new BinaryWriter(file);
                bw.Write(data);
                bw.Close();
                file.Close();
                return AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType;
            }
            catch { return ""; }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellvalue"></param>
        /// <param name="ID"></param>
        /// <param name="invnumber"></param>
        /// <returns></returns>
        public bool ReadPDFinDB(string cellvalue = "", string ID = "", int? invnumber = 0)
        {
            try
            {
                var data = ProcessPDFData(cellvalue: cellvalue, ID: ID, invnumber: invnumber);
                if (data.Length == 0) return false;
                var fileType = ".pdf";
                string tmp = Guid.NewGuid().ToString();
                var file = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType, FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
                var bw = new BinaryWriter(file);
                bw.Write(data);
                bw.Close();
                file.Close();
                System.Diagnostics.Process.Start(AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType);
                return true;
            }
            catch { return false; }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public static byte[] Serialize(object data)
        {
            System.Runtime.Serialization.Formatters.Binary.BinaryFormatter formatter = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();
            MemoryStream rems = new MemoryStream();
            formatter.Serialize(rems, data);
            return rems.ToArray();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellvalue"></param>
        /// <returns></returns>
        public bool ReadPDFinViewDocumentSetting(string cellvalue)
        {
            Array arr = vdPrefix();
            foreach (string prefix in arr)
                if (prefix.Trim().ToLower() == cellvalue.Substring(0, prefix.Trim().Length).Trim().ToLower())
                {
                    string filename = vdFilepath(prefix).Trim();
                    bool file = vdUseFile(prefix);
                    if (file)
                        filename = filename + "\\" + cellvalue;
                    else
                        filename = filename + "\\" + vdFilename(prefix).Trim();
                    string ftype = vdFiletype(prefix).Trim();
                    if (ftype == "pdf")
                    {
                        try
                        {
                            System.Diagnostics.Process.Start(filename.Trim() + "." + ftype);
                            return true;
                        }
                        catch
                        {
                            return false;
                        }
                    }
                }
            return false;
        }
        /// <summary>
        /// 
        /// </summary>
        public DataTable InitialReportTemplates()
        {
            var path = GetRootPath();
            DataTable dt = new DataTable();
            dt.Columns.Add("value");
            dt.Columns.Add("name");
            if (path != null)
            {
                DirectoryInfo di = new DirectoryInfo(path);
                DirectoryInfo[] fldrs = di.GetDirectories("*.*");
                foreach (DirectoryInfo d in fldrs)
                    if (Finance_Tools.IsAddedToGallery(d) && d.Name == "Reporting")
                    {
                        try
                        {
                            FileInfo[] myfile = d.GetFiles();
                            foreach (FileInfo f in myfile)
                            {
                                DirectoryInfo dir = new DirectoryInfo(d.FullName);
                                if (f.Extension == ".xls" || f.Extension == ".xlsx" || f.Extension == ".xlsm" || f.Extension == ".pdf")
                                {
                                    DataRow dr = dt.NewRow();
                                    dr["value"] = f.FullName;
                                    dr["name"] = f.Name;
                                    dt.Rows.Add(dr);
                                }
                            }
                        }
                        catch { }
                    }
            }
            return dt;
        }
        /// <summary>
        /// 
        /// </summary>
        public DataTable GetLocalFolders()
        {
            var path = GetRootPath();
            DataTable dt = new DataTable();
            dt.Columns.Add("value");
            dt.Columns.Add("name");
            if (path != null)
            {
                DirectoryInfo di = new DirectoryInfo(path);
                DirectoryInfo[] fldrs = di.GetDirectories("*.*");
                foreach (DirectoryInfo d in fldrs)
                    if (Finance_Tools.IsAddedToGallery(d))
                    {
                        try
                        {
                            DataRow dr = dt.NewRow();
                            dr["value"] = d.FullName;
                            dr["name"] = d.Name;
                            dt.Rows.Add(dr);
                        }
                        catch { }
                    }
            }
            return dt;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="folderPath"></param>
        /// <returns></returns>
        public DataTable GetLocalFolderFiles(string folderPath)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("value");
            dt.Columns.Add("name");
            if (folderPath != null)
            {
                DirectoryInfo di = new DirectoryInfo(folderPath);
                if (Finance_Tools.IsAddedToGallery(di))
                {
                    try
                    {
                        FileInfo[] myfile = di.GetFiles();
                        foreach (FileInfo f in myfile)
                        {
                            DirectoryInfo dir = new DirectoryInfo(di.FullName);
                            if (f.Extension == ".xls" || f.Extension == ".xlsx" || f.Extension == ".xlsm" || f.Extension == ".pdf")
                            {
                                DataRow dr = dt.NewRow();
                                dr["value"] = f.FullName;
                                dr["name"] = f.Name;
                                dt.Rows.Add(dr);
                            }
                        }
                    }
                    catch { }
                }
            }
            return dt;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="varlist"></param>
        /// <returns></returns>
        public DataTable ToDataTable<T>(IEnumerable<T> varlist)
        {
            DataTable dtReturn = new DataTable();
            PropertyInfo[] oProps = null;
            if (varlist == null)
                return dtReturn;
            foreach (T rec in varlist)
            {
                if (oProps == null)
                {
                    oProps = ((Type)rec.GetType()).GetProperties();
                    foreach (PropertyInfo pi in oProps)
                    {
                        Type colType = pi.PropertyType;
                        if ((colType.IsGenericType) && (colType.GetGenericTypeDefinition()
                             == typeof(Nullable<>)))
                        {
                            colType = colType.GetGenericArguments()[0];
                        }
                        dtReturn.Columns.Add(new DataColumn(pi.Name, colType));
                    }
                }
                DataRow dr = dtReturn.NewRow();
                foreach (PropertyInfo pi in oProps)
                {
                    dr[pi.Name] = pi.GetValue(rec, null) == null ? DBNull.Value : pi.GetValue
                    (rec, null);
                }
                dtReturn.Rows.Add(dr);
            }
            return dtReturn;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="inList"></param>
        /// <returns></returns>
        public List<Line> IniSpecialList(List<Specialist> inList)
        {
            List<Line> list = new List<Line>();
            for (int i = 0; i < inList.Count; i++)
            {
                Line a = new Line();
                a.DetailLad = new DetailLad();
                a.Ledger = inList[i].Ledger;
                a.AccountCode = inList[i].AccountCode;
                a.AccountingPeriod = inList[i].AccountingPeriod;
                a.AllocationMarker = inList[i].AllocationMarker;
                a.AnalysisCode1 = inList[i].AnalysisCode1;
                a.AnalysisCode10 = inList[i].AnalysisCode10;
                a.AnalysisCode2 = inList[i].AnalysisCode2;
                a.AnalysisCode3 = inList[i].AnalysisCode3;
                a.AnalysisCode4 = inList[i].AnalysisCode4;
                a.AnalysisCode5 = inList[i].AnalysisCode5;
                a.AnalysisCode6 = inList[i].AnalysisCode6;
                a.AnalysisCode7 = inList[i].AnalysisCode7;
                a.AnalysisCode8 = inList[i].AnalysisCode8;
                a.AnalysisCode9 = inList[i].AnalysisCode9;
                a.Base2ReportingAmount = inList[i].Base2ReportingAmount;
                a.BaseAmount = inList[i].BaseAmount;
                a.CurrencyCode = inList[i].CurrencyCode;
                a.DebitCredit = inList[i].DebitCredit;
                a.Description = inList[i].Description;
                a.DetailLad.AccountCode = inList[i].AccountCode;
                a.DetailLad.AccountingPeriod = inList[i].AccountingPeriod;
                a.DetailLad.GeneralDescription1 = inList[i].GenDesc1;
                a.DetailLad.GeneralDescription2 = inList[i].GenDesc2;
                a.DetailLad.GeneralDescription3 = inList[i].GenDesc3;
                a.DetailLad.GeneralDescription4 = inList[i].GenDesc4;
                a.DetailLad.GeneralDescription5 = inList[i].GenDesc5;
                a.DetailLad.GeneralDescription6 = inList[i].GenDesc6;
                a.DetailLad.GeneralDescription7 = inList[i].GenDesc7;
                a.DetailLad.GeneralDescription8 = inList[i].GenDesc8;
                a.DetailLad.GeneralDescription9 = inList[i].GenDesc9;
                a.DetailLad.GeneralDescription10 = inList[i].GenDesc10;
                a.DetailLad.GeneralDescription11 = inList[i].GenDesc11;
                a.DetailLad.GeneralDescription12 = inList[i].GenDesc12;
                a.DetailLad.GeneralDescription13 = inList[i].GenDesc13;
                a.DetailLad.GeneralDescription14 = inList[i].GenDesc14;
                a.DetailLad.GeneralDescription15 = inList[i].GenDesc15;
                a.DetailLad.GeneralDescription16 = inList[i].GenDesc16;
                a.DetailLad.GeneralDescription17 = inList[i].GenDesc17;
                a.DetailLad.GeneralDescription18 = inList[i].GenDesc18;
                a.DetailLad.GeneralDescription19 = inList[i].GenDesc19;
                a.DetailLad.GeneralDescription20 = inList[i].GenDesc20;
                a.DetailLad.GeneralDescription21 = inList[i].GenDesc21;
                a.DetailLad.GeneralDescription22 = inList[i].GenDesc22;
                a.DetailLad.GeneralDescription23 = inList[i].GenDesc23;
                a.DetailLad.GeneralDescription24 = inList[i].GenDesc24;
                a.DetailLad.GeneralDescription25 = inList[i].GenDesc25;
                a.JournalSource = inList[i].JournalSource;
                a.JournalType = inList[i].JournalType;
                a.TransactionAmount = inList[i].TransactionAmount;
                a.TransactionDate = inList[i].TransactionDate;
                a.DueDate = inList[i].DueDate;
                a.TransactionReference = inList[i].TransactionReference;
                a.Value4Amount = inList[i].Value4Amount;
                list.Add(a);
            }
            return list;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dgv"></param>
        /// <returns></returns>
        public DataTable IniBalanceBy(DataGridView dgv)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("name");
            dt.Columns.Add("value");
            dt.Rows.Add("", "");
            foreach (DataGridViewColumn dc in dgv.Columns)
            {
                if ("AnalysisCode1" == dc.Name) dt.Rows.Add(dc.HeaderText, "AnalysisCode1");
                if ("AnalysisCode2" == dc.Name) dt.Rows.Add(dc.HeaderText, "AnalysisCode2");
                if ("AnalysisCode3" == dc.Name) dt.Rows.Add(dc.HeaderText, "AnalysisCode3");
                if ("AnalysisCode4" == dc.Name) dt.Rows.Add(dc.HeaderText, "AnalysisCode4");
                if ("AnalysisCode5" == dc.Name) dt.Rows.Add(dc.HeaderText, "AnalysisCode5");
                if ("AnalysisCode6" == dc.Name) dt.Rows.Add(dc.HeaderText, "AnalysisCode6");
                if ("AnalysisCode7" == dc.Name) dt.Rows.Add(dc.HeaderText, "AnalysisCode7");
                if ("AnalysisCode8" == dc.Name) dt.Rows.Add(dc.HeaderText, "AnalysisCode8");
                if ("AnalysisCode9" == dc.Name) dt.Rows.Add(dc.HeaderText, "AnalysisCode9");
                if ("AnalysisCode10" == dc.Name) dt.Rows.Add(dc.HeaderText, "AnalysisCode10");
            }
            return dt;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public DataGridView IniGrdForTransUpd()
        {
            try
            {
                DataTable dt = GetUserDataFriendlyName();
                DataColumn[] keys = new DataColumn[1];
                keys[0] = dt.Columns["SunField"];
                dt.PrimaryKey = keys;
                DataTable dt2 = GetUserDataGenFriendlyName();
                DataColumn[] keys2 = new DataColumn[1];
                keys2[0] = dt2.Columns["SunField"];
                dt2.PrimaryKey = keys2;
                DataGridView d = new DataGridView();
                if (dt.Rows.Count > 0)
                {
                    d.Columns.Add("LineIndicator", "Line Indicator");
                    d.Columns["LineIndicator"].DataPropertyName = "LineIndicator";
                    d.Columns.Add("JournalNumber", "JournalNumber");
                    d.Columns["JournalNumber"].DataPropertyName = "JournalNumber";
                    d.Columns.Add("JournalLineNumber", "JournalLineNumber");
                    d.Columns["JournalLineNumber"].DataPropertyName = "JournalLineNumber";
                    d.Columns.Add("Ledger", dt.Rows.Find("Ledger")["UserFriendlyName"].ToString());
                    d.Columns["Ledger"].DataPropertyName = "Ledger";
                    d.Columns.Add("AccountCode", dt.Rows.Find("AccountCode")["UserFriendlyName"].ToString());
                    d.Columns["AccountCode"].DataPropertyName = "ft_Account";
                    d.Columns.Add("AccountingPeriod", dt.Rows.Find("AccountingPeriod")["UserFriendlyName"].ToString());
                    d.Columns["AccountingPeriod"].DataPropertyName = "Period";
                    d.Columns.Add("TransactionDate", dt.Rows.Find("TransactionDate")["UserFriendlyName"].ToString());
                    d.Columns["TransactionDate"].DataPropertyName = "TransDate";
                    d.Columns.Add("DueDate", dt.Rows.Find("DueDate")["UserFriendlyName"].ToString());
                    d.Columns["DueDate"].DataPropertyName = "DueDate";
                    d.Columns.Add("JournalType", dt.Rows.Find("JournalType")["UserFriendlyName"].ToString());
                    d.Columns["JournalType"].DataPropertyName = "JrnlType";
                    d.Columns.Add("JournalSource", dt.Rows.Find("JournalSource")["UserFriendlyName"].ToString());
                    d.Columns["JournalSource"].DataPropertyName = "JrnlSource";
                    d.Columns.Add("TransactionReference", dt.Rows.Find("TransactionReference")["UserFriendlyName"].ToString());
                    d.Columns["TransactionReference"].DataPropertyName = "TransRef";
                    d.Columns.Add("Description", dt.Rows.Find("Description")["UserFriendlyName"].ToString());
                    d.Columns["Description"].DataPropertyName = "Description";
                    d.Columns.Add("AllocationMarker", dt.Rows.Find("AllocationMarker")["UserFriendlyName"].ToString());
                    d.Columns["AllocationMarker"].DataPropertyName = "AlloctnMarker";
                    d.Columns.Add("AnalysisCode1", dt.Rows.Find("AnalysisCode1")["UserFriendlyName"].ToString());
                    d.Columns["AnalysisCode1"].DataPropertyName = "LA1";
                    d.Columns.Add("AnalysisCode2", dt.Rows.Find("AnalysisCode2")["UserFriendlyName"].ToString());
                    d.Columns["AnalysisCode2"].DataPropertyName = "LA2";
                    d.Columns.Add("AnalysisCode3", dt.Rows.Find("AnalysisCode3")["UserFriendlyName"].ToString());
                    d.Columns["AnalysisCode3"].DataPropertyName = "LA3";
                    d.Columns.Add("AnalysisCode4", dt.Rows.Find("AnalysisCode4")["UserFriendlyName"].ToString());
                    d.Columns["AnalysisCode4"].DataPropertyName = "LA4";
                    d.Columns.Add("AnalysisCode5", dt.Rows.Find("AnalysisCode5")["UserFriendlyName"].ToString());
                    d.Columns["AnalysisCode5"].DataPropertyName = "LA5";
                    d.Columns.Add("AnalysisCode6", dt.Rows.Find("AnalysisCode6")["UserFriendlyName"].ToString());
                    d.Columns["AnalysisCode6"].DataPropertyName = "LA6";
                    d.Columns.Add("AnalysisCode7", dt.Rows.Find("AnalysisCode7")["UserFriendlyName"].ToString());
                    d.Columns["AnalysisCode7"].DataPropertyName = "LA7";
                    d.Columns.Add("AnalysisCode8", dt.Rows.Find("AnalysisCode8")["UserFriendlyName"].ToString());
                    d.Columns["AnalysisCode8"].DataPropertyName = "LA8";
                    d.Columns.Add("AnalysisCode9", dt.Rows.Find("AnalysisCode9")["UserFriendlyName"].ToString());
                    d.Columns["AnalysisCode9"].DataPropertyName = "LA9";
                    d.Columns.Add("AnalysisCode10", dt.Rows.Find("AnalysisCode10")["UserFriendlyName"].ToString());
                    d.Columns["AnalysisCode10"].DataPropertyName = "LA10";
                    if (dt2.Rows.Count > 0)
                    {
                        d.Columns.Add("GeneralDescription1", dt2.Rows.Find("GeneralDescription1")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription1"].DataPropertyName = "GenDesc1";
                        d.Columns.Add("GeneralDescription2", dt2.Rows.Find("GeneralDescription2")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription2"].DataPropertyName = "GenDesc2";
                        d.Columns.Add("GeneralDescription3", dt2.Rows.Find("GeneralDescription3")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription3"].DataPropertyName = "GenDesc3";
                        d.Columns.Add("GeneralDescription4", dt2.Rows.Find("GeneralDescription4")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription4"].DataPropertyName = "GenDesc4";
                        d.Columns.Add("GeneralDescription5", dt2.Rows.Find("GeneralDescription5")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription5"].DataPropertyName = "GenDesc5";
                        d.Columns.Add("GeneralDescription6", dt2.Rows.Find("GeneralDescription6")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription6"].DataPropertyName = "GenDesc6";
                        d.Columns.Add("GeneralDescription7", dt2.Rows.Find("GeneralDescription7")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription7"].DataPropertyName = "GenDesc7";
                        d.Columns.Add("GeneralDescription8", dt2.Rows.Find("GeneralDescription8")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription8"].DataPropertyName = "GenDesc8";
                        d.Columns.Add("GeneralDescription9", dt2.Rows.Find("GeneralDescription9")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription9"].DataPropertyName = "GenDesc9";
                        d.Columns.Add("GeneralDescription10", dt2.Rows.Find("GeneralDescription10")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription10"].DataPropertyName = "GenDesc10";
                        d.Columns.Add("GeneralDescription11", dt2.Rows.Find("GeneralDescription11")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription11"].DataPropertyName = "GenDesc11";
                        d.Columns.Add("GeneralDescription12", dt2.Rows.Find("GeneralDescription12")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription12"].DataPropertyName = "GenDesc12";
                        d.Columns.Add("GeneralDescription13", dt2.Rows.Find("GeneralDescription13")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription13"].DataPropertyName = "GenDesc13";
                        d.Columns.Add("GeneralDescription14", dt2.Rows.Find("GeneralDescription14")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription14"].DataPropertyName = "GenDesc14";
                        d.Columns.Add("GeneralDescription15", dt2.Rows.Find("GeneralDescription15")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription15"].DataPropertyName = "GenDesc15";
                        d.Columns.Add("GeneralDescription16", dt2.Rows.Find("GeneralDescription16")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription16"].DataPropertyName = "GenDesc16";
                        d.Columns.Add("GeneralDescription17", dt2.Rows.Find("GeneralDescription17")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription17"].DataPropertyName = "GenDesc17";
                        d.Columns.Add("GeneralDescription18", dt2.Rows.Find("GeneralDescription18")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription18"].DataPropertyName = "GenDesc18";
                        d.Columns.Add("GeneralDescription19", dt2.Rows.Find("GeneralDescription19")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription19"].DataPropertyName = "GenDesc19";
                        d.Columns.Add("GeneralDescription20", dt2.Rows.Find("GeneralDescription20")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription20"].DataPropertyName = "GenDesc20";
                        d.Columns.Add("GeneralDescription21", dt2.Rows.Find("GeneralDescription21")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription21"].DataPropertyName = "GenDesc21";
                        d.Columns.Add("GeneralDescription22", dt2.Rows.Find("GeneralDescription22")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription22"].DataPropertyName = "GenDesc22";
                        d.Columns.Add("GeneralDescription23", dt2.Rows.Find("GeneralDescription23")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription23"].DataPropertyName = "GenDesc23";
                        d.Columns.Add("GeneralDescription24", dt2.Rows.Find("GeneralDescription24")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription24"].DataPropertyName = "GenDesc24";
                        d.Columns.Add("GeneralDescription25", dt2.Rows.Find("GeneralDescription25")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription25"].DataPropertyName = "GenDesc25";
                    }
                    else
                    {
                        d.Columns.Add("GeneralDescription1", "Gen Desc1");
                        d.Columns["GeneralDescription1"].DataPropertyName = "GenDesc1";
                        d.Columns.Add("GeneralDescription2", "Gen Desc2");
                        d.Columns["GeneralDescription2"].DataPropertyName = "GenDesc2";
                        d.Columns.Add("GeneralDescription3", "Gen Desc3");
                        d.Columns["GeneralDescription3"].DataPropertyName = "GenDesc3";
                        d.Columns.Add("GeneralDescription4", "Gen Desc4");
                        d.Columns["GeneralDescription4"].DataPropertyName = "GenDesc4";
                        d.Columns.Add("GeneralDescription5", "Gen Desc5");
                        d.Columns["GeneralDescription5"].DataPropertyName = "GenDesc5";
                        d.Columns.Add("GeneralDescription6", "Gen Desc6");
                        d.Columns["GeneralDescription6"].DataPropertyName = "GenDesc6";
                        d.Columns.Add("GeneralDescription7", "Gen Desc7");
                        d.Columns["GeneralDescription7"].DataPropertyName = "GenDesc7";
                        d.Columns.Add("GeneralDescription8", "Gen Desc8");
                        d.Columns["GeneralDescription8"].DataPropertyName = "GenDesc8";
                        d.Columns.Add("GeneralDescription9", "Gen Desc9");
                        d.Columns["GeneralDescription9"].DataPropertyName = "GenDesc9";
                        d.Columns.Add("GeneralDescription10", "Gen Desc10");
                        d.Columns["GeneralDescription10"].DataPropertyName = "GenDesc10";
                        d.Columns.Add("GeneralDescription11", "Gen Desc11");
                        d.Columns["GeneralDescription11"].DataPropertyName = "GenDesc11";
                        d.Columns.Add("GeneralDescription12", "Gen Desc12");
                        d.Columns["GeneralDescription12"].DataPropertyName = "GenDesc12";
                        d.Columns.Add("GeneralDescription13", "Gen Desc13");
                        d.Columns["GeneralDescription13"].DataPropertyName = "GenDesc13";
                        d.Columns.Add("GeneralDescription14", "Gen Desc14");
                        d.Columns["GeneralDescription14"].DataPropertyName = "GenDesc14";
                        d.Columns.Add("GeneralDescription15", "Gen Desc15");
                        d.Columns["GeneralDescription15"].DataPropertyName = "GenDesc15";
                        d.Columns.Add("GeneralDescription16", "Gen Desc16");
                        d.Columns["GeneralDescription16"].DataPropertyName = "GenDesc16";
                        d.Columns.Add("GeneralDescription17", "Gen Desc17");
                        d.Columns["GeneralDescription17"].DataPropertyName = "GenDesc17";
                        d.Columns.Add("GeneralDescription18", "Gen Desc18");
                        d.Columns["GeneralDescription18"].DataPropertyName = "GenDesc18";
                        d.Columns.Add("GeneralDescription19", "Gen Desc19");
                        d.Columns["GeneralDescription19"].DataPropertyName = "GenDesc19";
                        d.Columns.Add("GeneralDescription20", "Gen Desc20");
                        d.Columns["GeneralDescription20"].DataPropertyName = "GenDesc20";
                        d.Columns.Add("GeneralDescription21", "Gen Desc21");
                        d.Columns["GeneralDescription21"].DataPropertyName = "GenDesc21";
                        d.Columns.Add("GeneralDescription22", "Gen Desc22");
                        d.Columns["GeneralDescription22"].DataPropertyName = "GenDesc22";
                        d.Columns.Add("GeneralDescription23", "Gen Desc23");
                        d.Columns["GeneralDescription23"].DataPropertyName = "GenDesc23";
                        d.Columns.Add("GeneralDescription24", "Gen Desc24");
                        d.Columns["GeneralDescription24"].DataPropertyName = "GenDesc24";
                        d.Columns.Add("GeneralDescription25", "Gen Desc25");
                        d.Columns["GeneralDescription25"].DataPropertyName = "GenDesc25";
                    }
                    d.Columns.Add("TransactionAmount", dt.Rows.Find("TransactionAmount")["UserFriendlyName"].ToString());
                    d.Columns["TransactionAmount"].DataPropertyName = "TransAmount";
                    d.Columns.Add("CurrencyCode", dt.Rows.Find("CurrencyCode")["UserFriendlyName"].ToString());
                    d.Columns["CurrencyCode"].DataPropertyName = "Currency";
                    d.Columns.Add("BaseAmount", dt.Rows.Find("BaseAmount")["UserFriendlyName"].ToString());
                    d.Columns["BaseAmount"].DataPropertyName = "BaseAmount";
                    d.Columns.Add("Base2ReportingAmount", dt.Rows.Find("Base2ReportingAmount")["UserFriendlyName"].ToString());
                    d.Columns["Base2ReportingAmount"].DataPropertyName = "C2ndBase";
                    d.Columns.Add("Value4Amount", dt.Rows.Find("Value4Amount")["UserFriendlyName"].ToString());
                    d.Columns["Value4Amount"].DataPropertyName = "C4thAmount";
                }
                else
                {
                    d.Columns.Add("LineIndicator", "Line Indicator");
                    d.Columns["LineIndicator"].DataPropertyName = "LineIndicator";
                    d.Columns.Add("JournalNumber", "JournalNumber");
                    d.Columns["JournalNumber"].DataPropertyName = "JournalNumber";
                    d.Columns.Add("JournalLineNumber", "JournalLineNumber");
                    d.Columns["JournalLineNumber"].DataPropertyName = "JournalLineNumber";
                    d.Columns.Add("Ledger", "Ledger");
                    d.Columns["Ledger"].DataPropertyName = "Ledger";
                    d.Columns.Add("AccountCode", "Account");
                    d.Columns["AccountCode"].DataPropertyName = "ft_Account";
                    d.Columns.Add("AccountingPeriod", "Period");
                    d.Columns["AccountingPeriod"].DataPropertyName = "Period";
                    d.Columns.Add("TransactionDate", "Trans Date");
                    d.Columns["TransactionDate"].DataPropertyName = "TransDate";
                    d.Columns.Add("DueDate", "Due Date");
                    d.Columns["DueDate"].DataPropertyName = "DueDate";
                    d.Columns.Add("JournalType", "Jrnl Type");
                    d.Columns["JournalType"].DataPropertyName = "JrnlType";
                    d.Columns.Add("JournalSource", "Jrnl Source");
                    d.Columns["JournalSource"].DataPropertyName = "JrnlSource";
                    d.Columns.Add("TransactionReference", "Trans Ref");
                    d.Columns["TransactionReference"].DataPropertyName = "TransRef";
                    d.Columns.Add("Description", "Description");
                    d.Columns["Description"].DataPropertyName = "Description";
                    d.Columns.Add("AllocationMarker", "Alloctn Marker");
                    d.Columns["AllocationMarker"].DataPropertyName = "AlloctnMarker";
                    d.Columns.Add("AnalysisCode1", "LA1");
                    d.Columns["AnalysisCode1"].DataPropertyName = "LA1";
                    d.Columns.Add("AnalysisCode2", "LA2");
                    d.Columns["AnalysisCode2"].DataPropertyName = "LA2";
                    d.Columns.Add("AnalysisCode3", "LA3");
                    d.Columns["AnalysisCode3"].DataPropertyName = "LA3";
                    d.Columns.Add("AnalysisCode4", "LA4");
                    d.Columns["AnalysisCode4"].DataPropertyName = "LA4";
                    d.Columns.Add("AnalysisCode5", "LA5");
                    d.Columns["AnalysisCode5"].DataPropertyName = "LA5";
                    d.Columns.Add("AnalysisCode6", "LA6");
                    d.Columns["AnalysisCode6"].DataPropertyName = "LA6";
                    d.Columns.Add("AnalysisCode7", "LA7");
                    d.Columns["AnalysisCode7"].DataPropertyName = "LA7";
                    d.Columns.Add("AnalysisCode8", "LA8");
                    d.Columns["AnalysisCode8"].DataPropertyName = "LA8";
                    d.Columns.Add("AnalysisCode9", "LA9");
                    d.Columns["AnalysisCode9"].DataPropertyName = "LA9";
                    d.Columns.Add("AnalysisCode10", "LA10");
                    d.Columns["AnalysisCode10"].DataPropertyName = "LA10";
                    if (dt2.Rows.Count > 0)
                    {
                        d.Columns.Add("GeneralDescription1", dt2.Rows.Find("GeneralDescription1")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription1"].DataPropertyName = "GenDesc1";
                        d.Columns.Add("GeneralDescription2", dt2.Rows.Find("GeneralDescription2")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription2"].DataPropertyName = "GenDesc2";
                        d.Columns.Add("GeneralDescription3", dt2.Rows.Find("GeneralDescription3")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription3"].DataPropertyName = "GenDesc3";
                        d.Columns.Add("GeneralDescription4", dt2.Rows.Find("GeneralDescription4")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription4"].DataPropertyName = "GenDesc4";
                        d.Columns.Add("GeneralDescription5", dt2.Rows.Find("GeneralDescription5")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription5"].DataPropertyName = "GenDesc5";
                        d.Columns.Add("GeneralDescription6", dt2.Rows.Find("GeneralDescription6")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription6"].DataPropertyName = "GenDesc6";
                        d.Columns.Add("GeneralDescription7", dt2.Rows.Find("GeneralDescription7")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription7"].DataPropertyName = "GenDesc7";
                        d.Columns.Add("GeneralDescription8", dt2.Rows.Find("GeneralDescription8")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription8"].DataPropertyName = "GenDesc8";
                        d.Columns.Add("GeneralDescription9", dt2.Rows.Find("GeneralDescription9")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription9"].DataPropertyName = "GenDesc9";
                        d.Columns.Add("GeneralDescription10", dt2.Rows.Find("GeneralDescription10")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription10"].DataPropertyName = "GenDesc10";
                        d.Columns.Add("GeneralDescription11", dt2.Rows.Find("GeneralDescription11")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription11"].DataPropertyName = "GenDesc11";
                        d.Columns.Add("GeneralDescription12", dt2.Rows.Find("GeneralDescription12")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription12"].DataPropertyName = "GenDesc12";
                        d.Columns.Add("GeneralDescription13", dt2.Rows.Find("GeneralDescription13")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription13"].DataPropertyName = "GenDesc13";
                        d.Columns.Add("GeneralDescription14", dt2.Rows.Find("GeneralDescription14")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription14"].DataPropertyName = "GenDesc14";
                        d.Columns.Add("GeneralDescription15", dt2.Rows.Find("GeneralDescription15")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription15"].DataPropertyName = "GenDesc15";
                        d.Columns.Add("GeneralDescription16", dt2.Rows.Find("GeneralDescription16")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription16"].DataPropertyName = "GenDesc16";
                        d.Columns.Add("GeneralDescription17", dt2.Rows.Find("GeneralDescription17")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription17"].DataPropertyName = "GenDesc17";
                        d.Columns.Add("GeneralDescription18", dt2.Rows.Find("GeneralDescription18")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription18"].DataPropertyName = "GenDesc18";
                        d.Columns.Add("GeneralDescription19", dt2.Rows.Find("GeneralDescription19")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription19"].DataPropertyName = "GenDesc19";
                        d.Columns.Add("GeneralDescription20", dt2.Rows.Find("GeneralDescription20")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription20"].DataPropertyName = "GenDesc20";
                        d.Columns.Add("GeneralDescription21", dt2.Rows.Find("GeneralDescription21")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription21"].DataPropertyName = "GenDesc21";
                        d.Columns.Add("GeneralDescription22", dt2.Rows.Find("GeneralDescription22")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription22"].DataPropertyName = "GenDesc22";
                        d.Columns.Add("GeneralDescription23", dt2.Rows.Find("GeneralDescription23")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription23"].DataPropertyName = "GenDesc23";
                        d.Columns.Add("GeneralDescription24", dt2.Rows.Find("GeneralDescription24")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription24"].DataPropertyName = "GenDesc24";
                        d.Columns.Add("GeneralDescription25", dt2.Rows.Find("GeneralDescription25")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription25"].DataPropertyName = "GenDesc25";
                    }
                    else
                    {
                        d.Columns.Add("GeneralDescription1", "Gen Desc1");
                        d.Columns["GeneralDescription1"].DataPropertyName = "GenDesc1";
                        d.Columns.Add("GeneralDescription2", "Gen Desc2");
                        d.Columns["GeneralDescription2"].DataPropertyName = "GenDesc2";
                        d.Columns.Add("GeneralDescription3", "Gen Desc3");
                        d.Columns["GeneralDescription3"].DataPropertyName = "GenDesc3";
                        d.Columns.Add("GeneralDescription4", "Gen Desc4");
                        d.Columns["GeneralDescription4"].DataPropertyName = "GenDesc4";
                        d.Columns.Add("GeneralDescription5", "Gen Desc5");
                        d.Columns["GeneralDescription5"].DataPropertyName = "GenDesc5";
                        d.Columns.Add("GeneralDescription6", "Gen Desc6");
                        d.Columns["GeneralDescription6"].DataPropertyName = "GenDesc6";
                        d.Columns.Add("GeneralDescription7", "Gen Desc7");
                        d.Columns["GeneralDescription7"].DataPropertyName = "GenDesc7";
                        d.Columns.Add("GeneralDescription8", "Gen Desc8");
                        d.Columns["GeneralDescription8"].DataPropertyName = "GenDesc8";
                        d.Columns.Add("GeneralDescription9", "Gen Desc9");
                        d.Columns["GeneralDescription9"].DataPropertyName = "GenDesc9";
                        d.Columns.Add("GeneralDescription10", "Gen Desc10");
                        d.Columns["GeneralDescription10"].DataPropertyName = "GenDesc10";
                        d.Columns.Add("GeneralDescription11", "Gen Desc11");
                        d.Columns["GeneralDescription11"].DataPropertyName = "GenDesc11";
                        d.Columns.Add("GeneralDescription12", "Gen Desc12");
                        d.Columns["GeneralDescription12"].DataPropertyName = "GenDesc12";
                        d.Columns.Add("GeneralDescription13", "Gen Desc13");
                        d.Columns["GeneralDescription13"].DataPropertyName = "GenDesc13";
                        d.Columns.Add("GeneralDescription14", "Gen Desc14");
                        d.Columns["GeneralDescription14"].DataPropertyName = "GenDesc14";
                        d.Columns.Add("GeneralDescription15", "Gen Desc15");
                        d.Columns["GeneralDescription15"].DataPropertyName = "GenDesc15";
                        d.Columns.Add("GeneralDescription16", "Gen Desc16");
                        d.Columns["GeneralDescription16"].DataPropertyName = "GenDesc16";
                        d.Columns.Add("GeneralDescription17", "Gen Desc17");
                        d.Columns["GeneralDescription17"].DataPropertyName = "GenDesc17";
                        d.Columns.Add("GeneralDescription18", "Gen Desc18");
                        d.Columns["GeneralDescription18"].DataPropertyName = "GenDesc18";
                        d.Columns.Add("GeneralDescription19", "Gen Desc19");
                        d.Columns["GeneralDescription19"].DataPropertyName = "GenDesc19";
                        d.Columns.Add("GeneralDescription20", "Gen Desc20");
                        d.Columns["GeneralDescription20"].DataPropertyName = "GenDesc20";
                        d.Columns.Add("GeneralDescription21", "Gen Desc21");
                        d.Columns["GeneralDescription21"].DataPropertyName = "GenDesc21";
                        d.Columns.Add("GeneralDescription22", "Gen Desc22");
                        d.Columns["GeneralDescription22"].DataPropertyName = "GenDesc22";
                        d.Columns.Add("GeneralDescription23", "Gen Desc23");
                        d.Columns["GeneralDescription23"].DataPropertyName = "GenDesc23";
                        d.Columns.Add("GeneralDescription24", "Gen Desc24");
                        d.Columns["GeneralDescription24"].DataPropertyName = "GenDesc24";
                        d.Columns.Add("GeneralDescription25", "Gen Desc25");
                        d.Columns["GeneralDescription25"].DataPropertyName = "GenDesc25";
                    }
                    d.Columns.Add("TransactionAmount", "Trans Amount");
                    d.Columns["TransactionAmount"].DataPropertyName = "TransAmount";
                    d.Columns.Add("CurrencyCode", "Currency");
                    d.Columns["CurrencyCode"].DataPropertyName = "Currency";
                    d.Columns.Add("BaseAmount", "Base Amount");
                    d.Columns["BaseAmount"].DataPropertyName = "BaseAmount";
                    d.Columns.Add("Base2ReportingAmount", "2nd Base");
                    d.Columns["Base2ReportingAmount"].DataPropertyName = "C2ndBase";
                    d.Columns.Add("Value4Amount", "4th Amount");
                    d.Columns["Value4Amount"].DataPropertyName = "C4thAmount";
                }
                d.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
                d.AutoGenerateColumns = false;
                d.ColumnHeadersHeight = 40;
                d.Dock = DockStyle.Fill;
                d.Visible = true;
                d.DataBindingComplete += new DataGridViewBindingCompleteEventHandler(dataGridView1_DataBindingComplete);
                for (int i = 0; i < d.Columns.Count; i++)
                    d.Columns[i].Width = 55;
                return d;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + ex.StackTrace);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public DataGridView IniXMLTextGrd()
        {
            DataGridView d = new DataGridView();
            try
            {
                d.Columns.Add("Reference", "Ref");
                d.Columns["Reference"].DataPropertyName = "Reference";
                d.Columns.Add("CriteriaName1", "Criteria1");
                d.Columns["CriteriaName1"].DataPropertyName = "CriteriaName1";
                d.Columns.Add("CellReference1", "Cell Ref1");
                d.Columns["CellReference1"].DataPropertyName = "CellReference1";
                d.Columns.Add("CriteriaName2", "Criteria2");
                d.Columns["CriteriaName2"].DataPropertyName = "CriteriaName2";
                d.Columns.Add("CellReference2", "Cell Ref2");
                d.Columns["CellReference2"].DataPropertyName = "CellReference2";
                d.Columns.Add("CriteriaName3", "Criteria3");
                d.Columns["CriteriaName3"].DataPropertyName = "CriteriaName3";
                d.Columns.Add("CellReference3", "Cell Ref3");
                d.Columns["CellReference3"].DataPropertyName = "CellReference3";
                d.Columns.Add("CriteriaName4", "Criteria4");
                d.Columns["CriteriaName4"].DataPropertyName = "CriteriaName4";
                d.Columns.Add("CellReference4", "Cell Ref4");
                d.Columns["CellReference4"].DataPropertyName = "CellReference4";
                d.Columns.Add("CriteriaName5", "Criteria5");
                d.Columns["CriteriaName5"].DataPropertyName = "CriteriaName5";
                d.Columns.Add("CellReference5", "Cell Ref5");
                d.Columns["CellReference5"].DataPropertyName = "CellReference5";
                d.Columns.Add("OpenTransUponSave", "Open Transaction Upon Save");
                d.Columns["OpenTransUponSave"].DataPropertyName = "OpenTransUponSave";
                d.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
                d.AutoGenerateColumns = false;
                d.ColumnHeadersHeight = 40;
                d.Dock = DockStyle.Fill;
                d.Visible = true;
                for (int i = 0; i < d.Columns.Count; i++)
                    d.Columns[i].Width = 75;
                d.Columns[0].Width = 35;
                d.Columns[11].Width = 115;
                return d;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + ex.StackTrace);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public DataGridView IniSaveOptionsGrd()
        {
            DataGridView d = new DataGridView();
            try
            {
                d.Columns.Add("Reference", "Ref");
                d.Columns["Reference"].DataPropertyName = "Reference";
                DataGridViewComboBoxColumn comboxc1 = new DataGridViewComboBoxColumn();
                comboxc1.HeaderText = "Criteria1";
                comboxc1.Name = "CriteriaName1";
                comboxc1.DataPropertyName = "CriteriaName1";
                comboxc1.SortMode = DataGridViewColumnSortMode.NotSortable;
                BindDropdowns(comboxc1, IniGrd());
                d.Columns.Add(comboxc1);
                d.Columns.Add("CellReference1", "Cell Ref1");
                d.Columns["CellReference1"].DataPropertyName = "CellReference1";
                DataGridViewComboBoxColumn comboxc2 = new DataGridViewComboBoxColumn();
                comboxc2.HeaderText = "Criteria2";
                comboxc2.Name = "CriteriaName2";
                comboxc2.DataPropertyName = "CriteriaName2";
                comboxc2.SortMode = DataGridViewColumnSortMode.NotSortable;
                BindDropdowns(comboxc2, IniGrd());
                d.Columns.Add(comboxc2);
                d.Columns.Add("CellReference2", "Cell Ref2");
                d.Columns["CellReference2"].DataPropertyName = "CellReference2";
                DataGridViewComboBoxColumn comboxc3 = new DataGridViewComboBoxColumn();
                comboxc3.HeaderText = "Criteria3";
                comboxc3.Name = "CriteriaName3";
                comboxc3.DataPropertyName = "CriteriaName3";
                comboxc3.SortMode = DataGridViewColumnSortMode.NotSortable;
                BindDropdowns(comboxc3, IniGrd());
                d.Columns.Add(comboxc3);
                d.Columns.Add("CellReference3", "Cell Ref3");
                d.Columns["CellReference3"].DataPropertyName = "CellReference3";
                DataGridViewComboBoxColumn comboxc4 = new DataGridViewComboBoxColumn();
                comboxc4.HeaderText = "Criteria4";
                comboxc4.Name = "CriteriaName4";
                comboxc4.DataPropertyName = "CriteriaName4";
                comboxc4.SortMode = DataGridViewColumnSortMode.NotSortable;
                BindDropdowns(comboxc4, IniGrd());
                d.Columns.Add(comboxc4);
                d.Columns.Add("CellReference4", "Cell Ref4");
                d.Columns["CellReference4"].DataPropertyName = "CellReference4";
                DataGridViewComboBoxColumn comboxc5 = new DataGridViewComboBoxColumn();
                comboxc5.HeaderText = "Criteria5";
                comboxc5.Name = "CriteriaName5";
                comboxc5.DataPropertyName = "CriteriaName5";
                comboxc5.SortMode = DataGridViewColumnSortMode.NotSortable;
                BindDropdowns(comboxc5, IniGrd());
                d.Columns.Add(comboxc5);
                d.Columns.Add("CellReference5", "Cell Ref5");
                d.Columns["CellReference5"].DataPropertyName = "CellReference5";
                DataGridViewComboBoxColumn combox = new DataGridViewComboBoxColumn();
                combox.HeaderText = "Open Transaction Upon Save";
                combox.Name = "OpenTransUponSave";
                combox.Items.Add("True");
                combox.Items.Add("False");
                combox.DataPropertyName = "OpenTransUponSave";
                combox.SortMode = DataGridViewColumnSortMode.NotSortable;
                d.Columns.Add(combox);
                d.Columns.Add("SequencePrefix", "Sequence Prefix");
                d.Columns["SequencePrefix"].DataPropertyName = "SequencePrefix";
                d.Columns.Add("PopulateCell", "PopulateCell With SN");
                d.Columns["PopulateCell"].DataPropertyName = "PopulateCell";
                d.Columns["PopulateCell"].ToolTipText = "PopulateCell With SequenceNumber";
                d.Columns.Add("PDFFolder", "PDF Folder");
                d.Columns["PDFFolder"].DataPropertyName = "PDFFolder";
                d.Columns["PDFFolder"].ToolTipText = "PDF Folder";
                d.Columns.Add("PDFName", "PDF Name");
                d.Columns["PDFName"].DataPropertyName = "PDFName";
                d.Columns["PDFName"].ToolTipText = "PDF Name";
                d.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
                d.AutoGenerateColumns = false;
                d.ColumnHeadersHeight = 40;
                d.Dock = DockStyle.Fill;
                d.Visible = true;
                for (int i = 0; i < d.Columns.Count; i++)
                    d.Columns[i].Width = 75;
                d.Columns[0].Width = 35;
                d.Columns[11].Width = 115;
                return d;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + ex.StackTrace);
            }
        }
        public DataTable IniABT()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("name");
            dt.Columns.Add("value");
            dt.Rows.Add("", "");
            dt.Rows.Add("Never", "0");
            dt.Rows.Add("Generate Exchange Differences", "1");
            dt.Rows.Add("Generate Balancing Transaction", "2");
            return dt;
        }
        public DataTable IniAPSA()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("name");
            dt.Columns.Add("value");
            dt.Rows.Add("", "");
            dt.Rows.Add("N", "N");
            dt.Rows.Add("Y", "Y");
            return dt;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public DataGridView IniGrd()
        {
            try
            {
                DataTable dt = GetUserDataFriendlyName();
                DataColumn[] keys = new DataColumn[1];
                keys[0] = dt.Columns["SunField"];
                dt.PrimaryKey = keys;
                DataTable dt2 = GetUserDataGenFriendlyName();
                DataColumn[] keys2 = new DataColumn[1];
                keys2[0] = dt2.Columns["SunField"];
                dt2.PrimaryKey = keys2;
                DataGridView d = new DataGridView();
                if (dt.Rows.Count > 0)
                {
                    d.Columns.Add("Reference", "Ref");
                    d.Columns["Reference"].DataPropertyName = "Reference";
                    d.Columns.Add("SaveReference", "Save Reference");
                    d.Columns["SaveReference"].DataPropertyName = "SaveReference";
                    DataGridViewComboBoxColumn combox = new DataGridViewComboBoxColumn();
                    combox.HeaderText = "Balance By";
                    combox.Name = "BalanceBy";
                    combox.DataPropertyName = "BalanceBy";
                    combox.SortMode = DataGridViewColumnSortMode.NotSortable;
                    combox.DataSource = IniBalanceBy(IniGrdForTransUpd());
                    combox.DisplayMember = "name";
                    combox.ValueMember = "value";
                    d.Columns.Add(combox);
                    DataGridViewComboBoxColumn combox2 = new DataGridViewComboBoxColumn();
                    combox2.HeaderText = "Allow Balancing";
                    combox2.Name = "AllowBalTrans";
                    combox2.DataPropertyName = "AllowBalTrans";
                    combox2.SortMode = DataGridViewColumnSortMode.NotSortable;
                    combox2.DataSource = IniABT();
                    combox2.DisplayMember = "name";
                    combox2.ValueMember = "value";
                    d.Columns.Add(combox2);
                    DataGridViewComboBoxColumn combox3 = new DataGridViewComboBoxColumn();
                    combox3.HeaderText = "Allow to Suspended";
                    combox3.Name = "AllowPostSuspAcco";
                    combox3.DataPropertyName = "AllowPostSuspAcco";
                    combox3.SortMode = DataGridViewColumnSortMode.NotSortable;
                    combox3.DataSource = IniAPSA();
                    combox3.DisplayMember = "name";
                    combox3.ValueMember = "value";
                    d.Columns.Add(combox3);
                    d.Columns.Add("LineIndicator", "Line Indicator");
                    d.Columns["LineIndicator"].DataPropertyName = "LineIndicator";
                    d.Columns.Add("StartinginCell", "Starting In Cell");
                    d.Columns["StartinginCell"].DataPropertyName = "StartinginCell";
                    d.Columns.Add("PopWithJNNumber", "PopulateCell With JN");
                    d.Columns["PopWithJNNumber"].DataPropertyName = "PopWithJNNumber";
                    d.Columns["PopWithJNNumber"].ToolTipText = "PopulateCell With JournalNumber";
                    d.Columns.Add("JournalNumber", "Journal Number");
                    d.Columns["JournalNumber"].DataPropertyName = "JournalNumber";
                    d.Columns.Add("JournalLineNumber", "Journal LN");
                    d.Columns["JournalLineNumber"].DataPropertyName = "JournalLineNumber";
                    d.Columns["StartinginCell"].Frozen = true;
                    d.Columns.Add("Ledger", dt.Rows.Find("Ledger")["UserFriendlyName"].ToString());
                    d.Columns["Ledger"].DataPropertyName = "Ledger";
                    d.Columns.Add("AccountCode", dt.Rows.Find("AccountCode")["UserFriendlyName"].ToString());
                    d.Columns["AccountCode"].DataPropertyName = "ft_Account";
                    d.Columns.Add("AccountingPeriod", dt.Rows.Find("AccountingPeriod")["UserFriendlyName"].ToString());
                    d.Columns["AccountingPeriod"].DataPropertyName = "Period";
                    d.Columns.Add("TransactionDate", dt.Rows.Find("TransactionDate")["UserFriendlyName"].ToString());
                    d.Columns["TransactionDate"].DataPropertyName = "TransDate";
                    d.Columns.Add("DueDate", dt.Rows.Find("DueDate")["UserFriendlyName"].ToString());
                    d.Columns["DueDate"].DataPropertyName = "DueDate";
                    d.Columns.Add("JournalType", dt.Rows.Find("JournalType")["UserFriendlyName"].ToString());
                    d.Columns["JournalType"].DataPropertyName = "JrnlType";
                    d.Columns.Add("JournalSource", dt.Rows.Find("JournalSource")["UserFriendlyName"].ToString());
                    d.Columns["JournalSource"].DataPropertyName = "JrnlSource";
                    d.Columns.Add("TransactionReference", dt.Rows.Find("TransactionReference")["UserFriendlyName"].ToString());
                    d.Columns["TransactionReference"].DataPropertyName = "TransRef";
                    d.Columns.Add("Description", dt.Rows.Find("Description")["UserFriendlyName"].ToString());
                    d.Columns["Description"].DataPropertyName = "Description";
                    d.Columns.Add("AllocationMarker", dt.Rows.Find("AllocationMarker")["UserFriendlyName"].ToString());
                    d.Columns["AllocationMarker"].DataPropertyName = "AlloctnMarker";
                    d.Columns.Add("AnalysisCode1", dt.Rows.Find("AnalysisCode1")["UserFriendlyName"].ToString());
                    d.Columns["AnalysisCode1"].DataPropertyName = "LA1";
                    d.Columns.Add("AnalysisCode2", dt.Rows.Find("AnalysisCode2")["UserFriendlyName"].ToString());
                    d.Columns["AnalysisCode2"].DataPropertyName = "LA2";
                    d.Columns.Add("AnalysisCode3", dt.Rows.Find("AnalysisCode3")["UserFriendlyName"].ToString());
                    d.Columns["AnalysisCode3"].DataPropertyName = "LA3";
                    d.Columns.Add("AnalysisCode4", dt.Rows.Find("AnalysisCode4")["UserFriendlyName"].ToString());
                    d.Columns["AnalysisCode4"].DataPropertyName = "LA4";
                    d.Columns.Add("AnalysisCode5", dt.Rows.Find("AnalysisCode5")["UserFriendlyName"].ToString());
                    d.Columns["AnalysisCode5"].DataPropertyName = "LA5";
                    d.Columns.Add("AnalysisCode6", dt.Rows.Find("AnalysisCode6")["UserFriendlyName"].ToString());
                    d.Columns["AnalysisCode6"].DataPropertyName = "LA6";
                    d.Columns.Add("AnalysisCode7", dt.Rows.Find("AnalysisCode7")["UserFriendlyName"].ToString());
                    d.Columns["AnalysisCode7"].DataPropertyName = "LA7";
                    d.Columns.Add("AnalysisCode8", dt.Rows.Find("AnalysisCode8")["UserFriendlyName"].ToString());
                    d.Columns["AnalysisCode8"].DataPropertyName = "LA8";
                    d.Columns.Add("AnalysisCode9", dt.Rows.Find("AnalysisCode9")["UserFriendlyName"].ToString());
                    d.Columns["AnalysisCode9"].DataPropertyName = "LA9";
                    d.Columns.Add("AnalysisCode10", dt.Rows.Find("AnalysisCode10")["UserFriendlyName"].ToString());
                    d.Columns["AnalysisCode10"].DataPropertyName = "LA10";
                    if (dt2.Rows.Count > 0)
                    {
                        d.Columns.Add("GeneralDescription1", dt2.Rows.Find("GeneralDescription1")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription1"].DataPropertyName = "GenDesc1";
                        d.Columns.Add("GeneralDescription2", dt2.Rows.Find("GeneralDescription2")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription2"].DataPropertyName = "GenDesc2";
                        d.Columns.Add("GeneralDescription3", dt2.Rows.Find("GeneralDescription3")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription3"].DataPropertyName = "GenDesc3";
                        d.Columns.Add("GeneralDescription4", dt2.Rows.Find("GeneralDescription4")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription4"].DataPropertyName = "GenDesc4";
                        d.Columns.Add("GeneralDescription5", dt2.Rows.Find("GeneralDescription5")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription5"].DataPropertyName = "GenDesc5";
                        d.Columns.Add("GeneralDescription6", dt2.Rows.Find("GeneralDescription6")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription6"].DataPropertyName = "GenDesc6";
                        d.Columns.Add("GeneralDescription7", dt2.Rows.Find("GeneralDescription7")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription7"].DataPropertyName = "GenDesc7";
                        d.Columns.Add("GeneralDescription8", dt2.Rows.Find("GeneralDescription8")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription8"].DataPropertyName = "GenDesc8";
                        d.Columns.Add("GeneralDescription9", dt2.Rows.Find("GeneralDescription9")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription9"].DataPropertyName = "GenDesc9";
                        d.Columns.Add("GeneralDescription10", dt2.Rows.Find("GeneralDescription10")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription10"].DataPropertyName = "GenDesc10";
                        d.Columns.Add("GeneralDescription11", dt2.Rows.Find("GeneralDescription11")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription11"].DataPropertyName = "GenDesc11";
                        d.Columns.Add("GeneralDescription12", dt2.Rows.Find("GeneralDescription12")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription12"].DataPropertyName = "GenDesc12";
                        d.Columns.Add("GeneralDescription13", dt2.Rows.Find("GeneralDescription13")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription13"].DataPropertyName = "GenDesc13";
                        d.Columns.Add("GeneralDescription14", dt2.Rows.Find("GeneralDescription14")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription14"].DataPropertyName = "GenDesc14";
                        d.Columns.Add("GeneralDescription15", dt2.Rows.Find("GeneralDescription15")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription15"].DataPropertyName = "GenDesc15";
                        d.Columns.Add("GeneralDescription16", dt2.Rows.Find("GeneralDescription16")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription16"].DataPropertyName = "GenDesc16";
                        d.Columns.Add("GeneralDescription17", dt2.Rows.Find("GeneralDescription17")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription17"].DataPropertyName = "GenDesc17";
                        d.Columns.Add("GeneralDescription18", dt2.Rows.Find("GeneralDescription18")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription18"].DataPropertyName = "GenDesc18";
                        d.Columns.Add("GeneralDescription19", dt2.Rows.Find("GeneralDescription19")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription19"].DataPropertyName = "GenDesc19";
                        d.Columns.Add("GeneralDescription20", dt2.Rows.Find("GeneralDescription20")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription20"].DataPropertyName = "GenDesc20";
                        d.Columns.Add("GeneralDescription21", dt2.Rows.Find("GeneralDescription21")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription21"].DataPropertyName = "GenDesc21";
                        d.Columns.Add("GeneralDescription22", dt2.Rows.Find("GeneralDescription22")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription22"].DataPropertyName = "GenDesc22";
                        d.Columns.Add("GeneralDescription23", dt2.Rows.Find("GeneralDescription23")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription23"].DataPropertyName = "GenDesc23";
                        d.Columns.Add("GeneralDescription24", dt2.Rows.Find("GeneralDescription24")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription24"].DataPropertyName = "GenDesc24";
                        d.Columns.Add("GeneralDescription25", dt2.Rows.Find("GeneralDescription25")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription25"].DataPropertyName = "GenDesc25";
                    }
                    else
                    {
                        d.Columns.Add("GeneralDescription1", "Gen Desc1");
                        d.Columns["GeneralDescription1"].DataPropertyName = "GenDesc1";
                        d.Columns.Add("GeneralDescription2", "Gen Desc2");
                        d.Columns["GeneralDescription2"].DataPropertyName = "GenDesc2";
                        d.Columns.Add("GeneralDescription3", "Gen Desc3");
                        d.Columns["GeneralDescription3"].DataPropertyName = "GenDesc3";
                        d.Columns.Add("GeneralDescription4", "Gen Desc4");
                        d.Columns["GeneralDescription4"].DataPropertyName = "GenDesc4";
                        d.Columns.Add("GeneralDescription5", "Gen Desc5");
                        d.Columns["GeneralDescription5"].DataPropertyName = "GenDesc5";
                        d.Columns.Add("GeneralDescription6", "Gen Desc6");
                        d.Columns["GeneralDescription6"].DataPropertyName = "GenDesc6";
                        d.Columns.Add("GeneralDescription7", "Gen Desc7");
                        d.Columns["GeneralDescription7"].DataPropertyName = "GenDesc7";
                        d.Columns.Add("GeneralDescription8", "Gen Desc8");
                        d.Columns["GeneralDescription8"].DataPropertyName = "GenDesc8";
                        d.Columns.Add("GeneralDescription9", "Gen Desc9");
                        d.Columns["GeneralDescription9"].DataPropertyName = "GenDesc9";
                        d.Columns.Add("GeneralDescription10", "Gen Desc10");
                        d.Columns["GeneralDescription10"].DataPropertyName = "GenDesc10";
                        d.Columns.Add("GeneralDescription11", "Gen Desc11");
                        d.Columns["GeneralDescription11"].DataPropertyName = "GenDesc11";
                        d.Columns.Add("GeneralDescription12", "Gen Desc12");
                        d.Columns["GeneralDescription12"].DataPropertyName = "GenDesc12";
                        d.Columns.Add("GeneralDescription13", "Gen Desc13");
                        d.Columns["GeneralDescription13"].DataPropertyName = "GenDesc13";
                        d.Columns.Add("GeneralDescription14", "Gen Desc14");
                        d.Columns["GeneralDescription14"].DataPropertyName = "GenDesc14";
                        d.Columns.Add("GeneralDescription15", "Gen Desc15");
                        d.Columns["GeneralDescription15"].DataPropertyName = "GenDesc15";
                        d.Columns.Add("GeneralDescription16", "Gen Desc16");
                        d.Columns["GeneralDescription16"].DataPropertyName = "GenDesc16";
                        d.Columns.Add("GeneralDescription17", "Gen Desc17");
                        d.Columns["GeneralDescription17"].DataPropertyName = "GenDesc17";
                        d.Columns.Add("GeneralDescription18", "Gen Desc18");
                        d.Columns["GeneralDescription18"].DataPropertyName = "GenDesc18";
                        d.Columns.Add("GeneralDescription19", "Gen Desc19");
                        d.Columns["GeneralDescription19"].DataPropertyName = "GenDesc19";
                        d.Columns.Add("GeneralDescription20", "Gen Desc20");
                        d.Columns["GeneralDescription20"].DataPropertyName = "GenDesc20";
                        d.Columns.Add("GeneralDescription21", "Gen Desc21");
                        d.Columns["GeneralDescription21"].DataPropertyName = "GenDesc21";
                        d.Columns.Add("GeneralDescription22", "Gen Desc22");
                        d.Columns["GeneralDescription22"].DataPropertyName = "GenDesc22";
                        d.Columns.Add("GeneralDescription23", "Gen Desc23");
                        d.Columns["GeneralDescription23"].DataPropertyName = "GenDesc23";
                        d.Columns.Add("GeneralDescription24", "Gen Desc24");
                        d.Columns["GeneralDescription24"].DataPropertyName = "GenDesc24";
                        d.Columns.Add("GeneralDescription25", "Gen Desc25");
                        d.Columns["GeneralDescription25"].DataPropertyName = "GenDesc25";
                    }
                    d.Columns.Add("TransactionAmount", dt.Rows.Find("TransactionAmount")["UserFriendlyName"].ToString());
                    d.Columns["TransactionAmount"].DataPropertyName = "TransAmount";
                    d.Columns.Add("CurrencyCode", dt.Rows.Find("CurrencyCode")["UserFriendlyName"].ToString());
                    d.Columns["CurrencyCode"].DataPropertyName = "Currency";
                    d.Columns.Add("BaseAmount", dt.Rows.Find("BaseAmount")["UserFriendlyName"].ToString());
                    d.Columns["BaseAmount"].DataPropertyName = "BaseAmount";
                    d.Columns.Add("Base2ReportingAmount", dt.Rows.Find("Base2ReportingAmount")["UserFriendlyName"].ToString());
                    d.Columns["Base2ReportingAmount"].DataPropertyName = "C2ndBase";
                    d.Columns.Add("Value4Amount", dt.Rows.Find("Value4Amount")["UserFriendlyName"].ToString());
                    d.Columns["Value4Amount"].DataPropertyName = "C4thAmount";
                }
                else
                {
                    d.Columns.Add("Reference", "Ref");
                    d.Columns["Reference"].DataPropertyName = "Reference";
                    d.Columns.Add("SaveReference", "Save Reference");
                    d.Columns["SaveReference"].DataPropertyName = "SaveReference";
                    DataGridViewComboBoxColumn combox = new DataGridViewComboBoxColumn();
                    combox.HeaderText = "Balance By";
                    combox.Name = "BalanceBy";
                    combox.DataPropertyName = "BalanceBy";
                    combox.SortMode = DataGridViewColumnSortMode.NotSortable;
                    combox.DataSource = IniBalanceBy(IniGrdForTransUpd());
                    combox.DisplayMember = "name";
                    combox.ValueMember = "value";
                    d.Columns.Add(combox);
                    DataGridViewComboBoxColumn combox2 = new DataGridViewComboBoxColumn();
                    combox2.HeaderText = "Allow Balancing";
                    combox2.Name = "AllowBalTrans";
                    combox2.DataPropertyName = "AllowBalTrans";
                    combox2.SortMode = DataGridViewColumnSortMode.NotSortable;
                    combox2.DataSource = IniABT();
                    combox2.DisplayMember = "name";
                    combox2.ValueMember = "value";
                    d.Columns.Add(combox2);
                    DataGridViewComboBoxColumn combox3 = new DataGridViewComboBoxColumn();
                    combox3.HeaderText = "Allow to Suspended";
                    combox3.Name = "AllowPostSuspAcco";
                    combox3.DataPropertyName = "AllowPostSuspAcco";
                    combox3.SortMode = DataGridViewColumnSortMode.NotSortable;
                    combox3.DataSource = IniAPSA();
                    combox3.DisplayMember = "name";
                    combox3.ValueMember = "value";
                    d.Columns.Add(combox3);
                    d.Columns.Add("LineIndicator", "Line Indicator");
                    d.Columns["LineIndicator"].DataPropertyName = "LineIndicator";
                    d.Columns.Add("StartinginCell", "Starting In Cell");
                    d.Columns["StartinginCell"].DataPropertyName = "StartinginCell";
                    d.Columns.Add("PopWithJNNumber", "PopulateCell With JN");
                    d.Columns["PopWithJNNumber"].DataPropertyName = "PopWithJNNumber";
                    d.Columns["PopWithJNNumber"].ToolTipText = "PopulateCell With JournalNumber";
                    d.Columns.Add("JournalNumber", "Journal Number");
                    d.Columns["JournalNumber"].DataPropertyName = "JournalNumber";
                    d.Columns.Add("JournalLineNumber", "Journal LN");
                    d.Columns["JournalLineNumber"].DataPropertyName = "JournalLineNumber";
                    d.Columns["StartinginCell"].Frozen = true;
                    d.Columns.Add("Ledger", "Ledger");
                    d.Columns["Ledger"].DataPropertyName = "Ledger";
                    d.Columns.Add("AccountCode", "Account");
                    d.Columns["AccountCode"].DataPropertyName = "ft_Account";
                    d.Columns.Add("AccountingPeriod", "Period");
                    d.Columns["AccountingPeriod"].DataPropertyName = "Period";
                    d.Columns.Add("TransactionDate", "Trans Date");
                    d.Columns["TransactionDate"].DataPropertyName = "TransDate";
                    d.Columns.Add("DueDate", "Due Date");
                    d.Columns["DueDate"].DataPropertyName = "DueDate";
                    d.Columns.Add("JournalType", "Jrnl Type");
                    d.Columns["JournalType"].DataPropertyName = "JrnlType";
                    d.Columns.Add("JournalSource", "Jrnl Source");
                    d.Columns["JournalSource"].DataPropertyName = "JrnlSource";
                    d.Columns.Add("TransactionReference", "Trans Ref");
                    d.Columns["TransactionReference"].DataPropertyName = "TransRef";
                    d.Columns.Add("Description", "Description");
                    d.Columns["Description"].DataPropertyName = "Description";
                    d.Columns.Add("AllocationMarker", "Alloctn Marker");
                    d.Columns["AllocationMarker"].DataPropertyName = "AlloctnMarker";
                    d.Columns.Add("AnalysisCode1", "LA1");
                    d.Columns["AnalysisCode1"].DataPropertyName = "LA1";
                    d.Columns.Add("AnalysisCode2", "LA2");
                    d.Columns["AnalysisCode2"].DataPropertyName = "LA2";
                    d.Columns.Add("AnalysisCode3", "LA3");
                    d.Columns["AnalysisCode3"].DataPropertyName = "LA3";
                    d.Columns.Add("AnalysisCode4", "LA4");
                    d.Columns["AnalysisCode4"].DataPropertyName = "LA4";
                    d.Columns.Add("AnalysisCode5", "LA5");
                    d.Columns["AnalysisCode5"].DataPropertyName = "LA5";
                    d.Columns.Add("AnalysisCode6", "LA6");
                    d.Columns["AnalysisCode6"].DataPropertyName = "LA6";
                    d.Columns.Add("AnalysisCode7", "LA7");
                    d.Columns["AnalysisCode7"].DataPropertyName = "LA7";
                    d.Columns.Add("AnalysisCode8", "LA8");
                    d.Columns["AnalysisCode8"].DataPropertyName = "LA8";
                    d.Columns.Add("AnalysisCode9", "LA9");
                    d.Columns["AnalysisCode9"].DataPropertyName = "LA9";
                    d.Columns.Add("AnalysisCode10", "LA10");
                    d.Columns["AnalysisCode10"].DataPropertyName = "LA10";
                    if (dt2.Rows.Count > 0)
                    {
                        d.Columns.Add("GeneralDescription1", dt2.Rows.Find("GeneralDescription1")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription1"].DataPropertyName = "GenDesc1";
                        d.Columns.Add("GeneralDescription2", dt2.Rows.Find("GeneralDescription2")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription2"].DataPropertyName = "GenDesc2";
                        d.Columns.Add("GeneralDescription3", dt2.Rows.Find("GeneralDescription3")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription3"].DataPropertyName = "GenDesc3";
                        d.Columns.Add("GeneralDescription4", dt2.Rows.Find("GeneralDescription4")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription4"].DataPropertyName = "GenDesc4";
                        d.Columns.Add("GeneralDescription5", dt2.Rows.Find("GeneralDescription5")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription5"].DataPropertyName = "GenDesc5";
                        d.Columns.Add("GeneralDescription6", dt2.Rows.Find("GeneralDescription6")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription6"].DataPropertyName = "GenDesc6";
                        d.Columns.Add("GeneralDescription7", dt2.Rows.Find("GeneralDescription7")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription7"].DataPropertyName = "GenDesc7";
                        d.Columns.Add("GeneralDescription8", dt2.Rows.Find("GeneralDescription8")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription8"].DataPropertyName = "GenDesc8";
                        d.Columns.Add("GeneralDescription9", dt2.Rows.Find("GeneralDescription9")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription9"].DataPropertyName = "GenDesc9";
                        d.Columns.Add("GeneralDescription10", dt2.Rows.Find("GeneralDescription10")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription10"].DataPropertyName = "GenDesc10";
                        d.Columns.Add("GeneralDescription11", dt2.Rows.Find("GeneralDescription11")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription11"].DataPropertyName = "GenDesc11";
                        d.Columns.Add("GeneralDescription12", dt2.Rows.Find("GeneralDescription12")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription12"].DataPropertyName = "GenDesc12";
                        d.Columns.Add("GeneralDescription13", dt2.Rows.Find("GeneralDescription13")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription13"].DataPropertyName = "GenDesc13";
                        d.Columns.Add("GeneralDescription14", dt2.Rows.Find("GeneralDescription14")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription14"].DataPropertyName = "GenDesc14";
                        d.Columns.Add("GeneralDescription15", dt2.Rows.Find("GeneralDescription15")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription15"].DataPropertyName = "GenDesc15";
                        d.Columns.Add("GeneralDescription16", dt2.Rows.Find("GeneralDescription16")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription16"].DataPropertyName = "GenDesc16";
                        d.Columns.Add("GeneralDescription17", dt2.Rows.Find("GeneralDescription17")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription17"].DataPropertyName = "GenDesc17";
                        d.Columns.Add("GeneralDescription18", dt2.Rows.Find("GeneralDescription18")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription18"].DataPropertyName = "GenDesc18";
                        d.Columns.Add("GeneralDescription19", dt2.Rows.Find("GeneralDescription19")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription19"].DataPropertyName = "GenDesc19";
                        d.Columns.Add("GeneralDescription20", dt2.Rows.Find("GeneralDescription20")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription20"].DataPropertyName = "GenDesc20";
                        d.Columns.Add("GeneralDescription21", dt2.Rows.Find("GeneralDescription21")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription21"].DataPropertyName = "GenDesc21";
                        d.Columns.Add("GeneralDescription22", dt2.Rows.Find("GeneralDescription22")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription22"].DataPropertyName = "GenDesc22";
                        d.Columns.Add("GeneralDescription23", dt2.Rows.Find("GeneralDescription23")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription23"].DataPropertyName = "GenDesc23";
                        d.Columns.Add("GeneralDescription24", dt2.Rows.Find("GeneralDescription24")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription24"].DataPropertyName = "GenDesc24";
                        d.Columns.Add("GeneralDescription25", dt2.Rows.Find("GeneralDescription25")["UserFriendlyName"].ToString());
                        d.Columns["GeneralDescription25"].DataPropertyName = "GenDesc25";
                    }
                    else
                    {
                        d.Columns.Add("GeneralDescription1", "Gen Desc1");
                        d.Columns["GeneralDescription1"].DataPropertyName = "GenDesc1";
                        d.Columns.Add("GeneralDescription2", "Gen Desc2");
                        d.Columns["GeneralDescription2"].DataPropertyName = "GenDesc2";
                        d.Columns.Add("GeneralDescription3", "Gen Desc3");
                        d.Columns["GeneralDescription3"].DataPropertyName = "GenDesc3";
                        d.Columns.Add("GeneralDescription4", "Gen Desc4");
                        d.Columns["GeneralDescription4"].DataPropertyName = "GenDesc4";
                        d.Columns.Add("GeneralDescription5", "Gen Desc5");
                        d.Columns["GeneralDescription5"].DataPropertyName = "GenDesc5";
                        d.Columns.Add("GeneralDescription6", "Gen Desc6");
                        d.Columns["GeneralDescription6"].DataPropertyName = "GenDesc6";
                        d.Columns.Add("GeneralDescription7", "Gen Desc7");
                        d.Columns["GeneralDescription7"].DataPropertyName = "GenDesc7";
                        d.Columns.Add("GeneralDescription8", "Gen Desc8");
                        d.Columns["GeneralDescription8"].DataPropertyName = "GenDesc8";
                        d.Columns.Add("GeneralDescription9", "Gen Desc9");
                        d.Columns["GeneralDescription9"].DataPropertyName = "GenDesc9";
                        d.Columns.Add("GeneralDescription10", "Gen Desc10");
                        d.Columns["GeneralDescription10"].DataPropertyName = "GenDesc10";
                        d.Columns.Add("GeneralDescription11", "Gen Desc11");
                        d.Columns["GeneralDescription11"].DataPropertyName = "GenDesc11";
                        d.Columns.Add("GeneralDescription12", "Gen Desc12");
                        d.Columns["GeneralDescription12"].DataPropertyName = "GenDesc12";
                        d.Columns.Add("GeneralDescription13", "Gen Desc13");
                        d.Columns["GeneralDescription13"].DataPropertyName = "GenDesc13";
                        d.Columns.Add("GeneralDescription14", "Gen Desc14");
                        d.Columns["GeneralDescription14"].DataPropertyName = "GenDesc14";
                        d.Columns.Add("GeneralDescription15", "Gen Desc15");
                        d.Columns["GeneralDescription15"].DataPropertyName = "GenDesc15";
                        d.Columns.Add("GeneralDescription16", "Gen Desc16");
                        d.Columns["GeneralDescription16"].DataPropertyName = "GenDesc16";
                        d.Columns.Add("GeneralDescription17", "Gen Desc17");
                        d.Columns["GeneralDescription17"].DataPropertyName = "GenDesc17";
                        d.Columns.Add("GeneralDescription18", "Gen Desc18");
                        d.Columns["GeneralDescription18"].DataPropertyName = "GenDesc18";
                        d.Columns.Add("GeneralDescription19", "Gen Desc19");
                        d.Columns["GeneralDescription19"].DataPropertyName = "GenDesc19";
                        d.Columns.Add("GeneralDescription20", "Gen Desc20");
                        d.Columns["GeneralDescription20"].DataPropertyName = "GenDesc20";
                        d.Columns.Add("GeneralDescription21", "Gen Desc21");
                        d.Columns["GeneralDescription21"].DataPropertyName = "GenDesc21";
                        d.Columns.Add("GeneralDescription22", "Gen Desc22");
                        d.Columns["GeneralDescription22"].DataPropertyName = "GenDesc22";
                        d.Columns.Add("GeneralDescription23", "Gen Desc23");
                        d.Columns["GeneralDescription23"].DataPropertyName = "GenDesc23";
                        d.Columns.Add("GeneralDescription24", "Gen Desc24");
                        d.Columns["GeneralDescription24"].DataPropertyName = "GenDesc24";
                        d.Columns.Add("GeneralDescription25", "Gen Desc25");
                        d.Columns["GeneralDescription25"].DataPropertyName = "GenDesc25";
                    }
                    d.Columns.Add("TransactionAmount", "Trans Amount");
                    d.Columns["TransactionAmount"].DataPropertyName = "TransAmount";
                    d.Columns.Add("CurrencyCode", "Currency");
                    d.Columns["CurrencyCode"].DataPropertyName = "Currency";
                    d.Columns.Add("BaseAmount", "Base Amount");
                    d.Columns["BaseAmount"].DataPropertyName = "BaseAmount";
                    d.Columns.Add("Base2ReportingAmount", "2nd Base");
                    d.Columns["Base2ReportingAmount"].DataPropertyName = "C2ndBase";
                    d.Columns.Add("Value4Amount", "4th Amount");
                    d.Columns["Value4Amount"].DataPropertyName = "C4thAmount";
                }
                d.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
                d.AutoGenerateColumns = false;
                d.ColumnHeadersHeight = 40;
                d.Dock = DockStyle.Fill;
                d.Visible = true;
                d.DataBindingComplete += new DataGridViewBindingCompleteEventHandler(dataGridView1_DataBindingComplete);
                for (int i = 0; i < d.Columns.Count; i++)
                    d.Columns[i].Width = 55;
                d.Columns["PopWithJNNumber"].Width = 65;
                d.Columns["Reference"].Width = 35;
                d.Columns["JournalLineNumber"].ToolTipText = "Journal Line Number";
                d.Columns["AllowBalTrans"].ToolTipText = "Allow Balancing Transactions";
                d.Columns["AllowPostSuspAcco"].ToolTipText = "Allow Posting to Suspended Accounts";
                return d;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + ex.StackTrace);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="i"></param>
        /// <returns></returns>
        public string getSeprStr(DataTable dt, int i)
        {
            string str = string.Empty;
            bool mandatory = dt.Rows.Count > i ? (bool)dt.Rows[i]["Mandatory"] : false;
            if (!mandatory) return getSpaceStr(str);
            string cellValue = dt.Rows.Count > i ? (dt.Rows[i]["Separator"] == null ? "" : dt.Rows[i]["Separator"].ToString()) : "";
            try
            {
                bool error = Ribbon2.wsRrigin.Cells.get_Range(cellValue).Errors.Item[1].Value;
                if (error || string.IsNullOrEmpty(cellValue))
                    str = cellValue;
                else
                    str = Ribbon2.wsRrigin.Cells.get_Range(cellValue).Value.ToString();
            }
            catch
            {
                str = cellValue;
            }
            return getSpaceStr(str);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="str"></param>
        /// <param name="dgv"></param>
        /// <param name="i"></param>
        /// <returns></returns>
        public string getFormatStr(string str, DataTable dt, int i)
        {
            string returnValue = getSpaceStr(str);
            string defaultValue = dt.Rows.Count > i ? (dt.Rows[i]["DefaultValue"] == null ? "" : dt.Rows[i]["DefaultValue"].ToString()) : "";
            if (string.IsNullOrEmpty(returnValue)) returnValue = defaultValue;
            bool mandatory = dt.Rows.Count > i ? (bool)dt.Rows[i]["Mandatory"] : false;
            if (!mandatory) return returnValue;
            string sTextLength = dt.Rows.Count > i ? (dt.Rows[i]["TextLength"] == null ? "0" : dt.Rows[i]["TextLength"].ToString()) : "";
            int iTextLength = -1;
            int.TryParse(sTextLength, out iTextLength);
            string iftrimText = dt.Rows.Count > i ? dt.Rows[i]["trimText"].ToString() : "";
            string prefix = dt.Rows.Count > i ? (dt.Rows[i]["Prefix"] == null ? "" : dt.Rows[i]["Prefix"].ToString()) : "";
            string suffix = dt.Rows.Count > i ? (dt.Rows[i]["Suffix"] == null ? "0" : dt.Rows[i]["Suffix"].ToString()) : "";
            char[] InvaidChars = dt.Rows.Count > i ? (dt.Rows[i]["RemoveCharacters"] == null ? null : dt.Rows[i]["RemoveCharacters"].ToString().ToCharArray()) : null;

            if (InvaidChars != null)
            {
                for (int k = 0; k < InvaidChars.Length; k++)
                {
                    returnValue = returnValue.Replace(InvaidChars[k].ToString(), "");
                }
            }
            if (string.IsNullOrEmpty(sTextLength)) return returnValue;
            if (returnValue.Length < iTextLength && !string.IsNullOrEmpty(prefix))
            {
                int count = iTextLength - returnValue.Length;
                for (int j = 0; j < count; j++)
                {
                    if (returnValue.Length < iTextLength)
                        returnValue = prefix + returnValue;
                }
            }
            else if (returnValue.Length < iTextLength && !string.IsNullOrEmpty(suffix))
            {
                int count = iTextLength - returnValue.Length;
                for (int j = 0; j < count; j++)
                {
                    if (returnValue.Length < iTextLength)
                        returnValue = returnValue + suffix;
                }
            }
            if ((iftrimText.ToLower() == "right") && (returnValue.Length > iTextLength) && (iTextLength != -1))
            {
                returnValue = returnValue.Remove(iTextLength);
            }
            else if ((iftrimText.ToLower() == "left") && (returnValue.Length > iTextLength) && (iTextLength != -1))
            {
                returnValue = returnValue.Substring(returnValue.Length - iTextLength, iTextLength);
            }
            else if ((iftrimText.ToLower() == "none") && (returnValue.Length > iTextLength) && (iTextLength != -1))
            {
                returnValue = "error";
            }
            return returnValue;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public string getSpaceStr(string str)
        {
            if (str.ToLower().Contains("space("))
            {
                string number = str.Substring(str.ToLower().IndexOf("space(") + 6, str.IndexOf(")", str.ToLower().IndexOf("space(") + 6) - (str.ToLower().IndexOf("space(") + 6));
                string returnValue = string.Empty;
                int outResult;
                if (int.TryParse(number, out outResult))
                    for (int i = 0; i < outResult; i++)
                    {
                        returnValue += " ";
                    }
                returnValue = str.Substring(0, str.ToLower().IndexOf("space(")) + returnValue + str.Substring(str.IndexOf(")", str.ToLower().IndexOf("space(") + 6) + 1, str.Length - str.IndexOf(")", str.ToLower().IndexOf("space(") + 6) - 1);

                return returnValue;
            }
            else
            {
                return str;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            DataTable dt = GetUserDataFriendlyName();
            DataColumn[] keys = new DataColumn[1];
            keys[0] = dt.Columns["SunField"];
            dt.PrimaryKey = keys;
            DataTable dt2 = GetUserDataGenFriendlyName();
            DataColumn[] keys2 = new DataColumn[1];
            keys2[0] = dt2.Columns["SunField"];
            dt2.PrimaryKey = keys2;
            if (dt.Rows.Count > 0)
            {
                ((DataGridView)sender).Columns["Ledger"].Visible = dt.Rows.Find("Ledger")["Output"].ToString() == "True" ? true : false;
                ((DataGridView)sender).Columns["AccountCode"].Visible = dt.Rows.Find("AccountCode")["Output"].ToString() == "True" ? true : false;
                ((DataGridView)sender).Columns["AccountingPeriod"].Visible = dt.Rows.Find("AccountingPeriod")["Output"].ToString() == "True" ? true : false;
                ((DataGridView)sender).Columns["TransactionDate"].Visible = dt.Rows.Find("TransactionDate")["Output"].ToString() == "True" ? true : false;
                ((DataGridView)sender).Columns["DueDate"].Visible = dt.Rows.Find("DueDate")["Output"].ToString() == "True" ? true : false;
                ((DataGridView)sender).Columns["JournalType"].Visible = dt.Rows.Find("JournalType")["Output"].ToString() == "True" ? true : false;
                ((DataGridView)sender).Columns["JournalSource"].Visible = dt.Rows.Find("JournalSource")["Output"].ToString() == "True" ? true : false;
                ((DataGridView)sender).Columns["TransactionReference"].Visible = dt.Rows.Find("TransactionReference")["Output"].ToString() == "True" ? true : false;
                ((DataGridView)sender).Columns["Description"].Visible = dt.Rows.Find("Description")["Output"].ToString() == "True" ? true : false;
                ((DataGridView)sender).Columns["AllocationMarker"].Visible = dt.Rows.Find("AllocationMarker")["Output"].ToString() == "True" ? true : false;
                ((DataGridView)sender).Columns["AnalysisCode1"].Visible = dt.Rows.Find("AnalysisCode1")["Output"].ToString() == "True" ? true : false;
                ((DataGridView)sender).Columns["AnalysisCode2"].Visible = dt.Rows.Find("AnalysisCode2")["Output"].ToString() == "True" ? true : false;
                ((DataGridView)sender).Columns["AnalysisCode3"].Visible = dt.Rows.Find("AnalysisCode3")["Output"].ToString() == "True" ? true : false;
                ((DataGridView)sender).Columns["AnalysisCode4"].Visible = dt.Rows.Find("AnalysisCode4")["Output"].ToString() == "True" ? true : false;
                ((DataGridView)sender).Columns["AnalysisCode5"].Visible = dt.Rows.Find("AnalysisCode5")["Output"].ToString() == "True" ? true : false;
                ((DataGridView)sender).Columns["AnalysisCode6"].Visible = dt.Rows.Find("AnalysisCode6")["Output"].ToString() == "True" ? true : false;
                ((DataGridView)sender).Columns["AnalysisCode7"].Visible = dt.Rows.Find("AnalysisCode7")["Output"].ToString() == "True" ? true : false;
                ((DataGridView)sender).Columns["AnalysisCode8"].Visible = dt.Rows.Find("AnalysisCode8")["Output"].ToString() == "True" ? true : false;
                ((DataGridView)sender).Columns["AnalysisCode9"].Visible = dt.Rows.Find("AnalysisCode9")["Output"].ToString() == "True" ? true : false;
                ((DataGridView)sender).Columns["AnalysisCode10"].Visible = dt.Rows.Find("AnalysisCode10")["Output"].ToString() == "True" ? true : false;
                if (dt2.Rows.Count > 0)
                {
                    ((DataGridView)sender).Columns["GeneralDescription1"].Visible = dt2.Rows.Find("GeneralDescription1")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription2"].Visible = dt2.Rows.Find("GeneralDescription2")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription3"].Visible = dt2.Rows.Find("GeneralDescription3")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription4"].Visible = dt2.Rows.Find("GeneralDescription4")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription5"].Visible = dt2.Rows.Find("GeneralDescription5")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription6"].Visible = dt2.Rows.Find("GeneralDescription6")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription7"].Visible = dt2.Rows.Find("GeneralDescription7")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription8"].Visible = dt2.Rows.Find("GeneralDescription8")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription9"].Visible = dt2.Rows.Find("GeneralDescription9")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription10"].Visible = dt2.Rows.Find("GeneralDescription10")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription11"].Visible = dt2.Rows.Find("GeneralDescription11")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription12"].Visible = dt2.Rows.Find("GeneralDescription12")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription13"].Visible = dt2.Rows.Find("GeneralDescription13")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription14"].Visible = dt2.Rows.Find("GeneralDescription14")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription15"].Visible = dt2.Rows.Find("GeneralDescription15")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription16"].Visible = dt2.Rows.Find("GeneralDescription16")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription17"].Visible = dt2.Rows.Find("GeneralDescription17")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription18"].Visible = dt2.Rows.Find("GeneralDescription18")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription19"].Visible = dt2.Rows.Find("GeneralDescription19")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription20"].Visible = dt2.Rows.Find("GeneralDescription20")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription21"].Visible = dt2.Rows.Find("GeneralDescription21")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription22"].Visible = dt2.Rows.Find("GeneralDescription22")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription23"].Visible = dt2.Rows.Find("GeneralDescription23")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription24"].Visible = dt2.Rows.Find("GeneralDescription24")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription25"].Visible = dt2.Rows.Find("GeneralDescription25")["Output"].ToString() == "True" ? true : false;
                }
                else
                {
                    ((DataGridView)sender).Columns["GeneralDescription1"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription2"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription3"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription4"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription5"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription6"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription7"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription8"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription9"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription10"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription11"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription12"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription13"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription14"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription15"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription16"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription17"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription18"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription19"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription20"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription21"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription22"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription23"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription24"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription25"].Visible = false;
                }
                ((DataGridView)sender).Columns["TransactionAmount"].Visible = dt.Rows.Find("TransactionAmount")["Output"].ToString() == "True" ? true : false;
                ((DataGridView)sender).Columns["CurrencyCode"].Visible = dt.Rows.Find("CurrencyCode")["Output"].ToString() == "True" ? true : false;
                ((DataGridView)sender).Columns["BaseAmount"].Visible = dt.Rows.Find("BaseAmount")["Output"].ToString() == "True" ? true : false;
                ((DataGridView)sender).Columns["Base2ReportingAmount"].Visible = dt.Rows.Find("Base2ReportingAmount")["Output"].ToString() == "True" ? true : false;
                ((DataGridView)sender).Columns["Value4Amount"].Visible = dt.Rows.Find("Value4Amount")["Output"].ToString() == "True" ? true : false;
            }
            else
            {
                if (dt2.Rows.Count > 0)
                {
                    ((DataGridView)sender).Columns["GeneralDescription1"].Visible = dt2.Rows.Find("GeneralDescription1")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription2"].Visible = dt2.Rows.Find("GeneralDescription2")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription3"].Visible = dt2.Rows.Find("GeneralDescription3")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription4"].Visible = dt2.Rows.Find("GeneralDescription4")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription5"].Visible = dt2.Rows.Find("GeneralDescription5")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription6"].Visible = dt2.Rows.Find("GeneralDescription6")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription7"].Visible = dt2.Rows.Find("GeneralDescription7")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription8"].Visible = dt2.Rows.Find("GeneralDescription8")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription9"].Visible = dt2.Rows.Find("GeneralDescription9")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription10"].Visible = dt2.Rows.Find("GeneralDescription10")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription11"].Visible = dt2.Rows.Find("GeneralDescription11")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription12"].Visible = dt2.Rows.Find("GeneralDescription12")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription13"].Visible = dt2.Rows.Find("GeneralDescription13")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription14"].Visible = dt2.Rows.Find("GeneralDescription14")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription15"].Visible = dt2.Rows.Find("GeneralDescription15")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription16"].Visible = dt2.Rows.Find("GeneralDescription16")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription17"].Visible = dt2.Rows.Find("GeneralDescription17")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription18"].Visible = dt2.Rows.Find("GeneralDescription18")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription19"].Visible = dt2.Rows.Find("GeneralDescription19")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription20"].Visible = dt2.Rows.Find("GeneralDescription20")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription21"].Visible = dt2.Rows.Find("GeneralDescription21")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription22"].Visible = dt2.Rows.Find("GeneralDescription22")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription23"].Visible = dt2.Rows.Find("GeneralDescription23")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription24"].Visible = dt2.Rows.Find("GeneralDescription24")["Output"].ToString() == "True" ? true : false;
                    ((DataGridView)sender).Columns["GeneralDescription25"].Visible = dt2.Rows.Find("GeneralDescription25")["Output"].ToString() == "True" ? true : false;
                }
                else
                {
                    ((DataGridView)sender).Columns["GeneralDescription1"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription2"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription3"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription4"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription5"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription6"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription7"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription8"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription9"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription10"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription11"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription12"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription13"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription14"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription15"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription16"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription17"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription18"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription19"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription20"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription21"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription22"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription23"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription24"].Visible = false;
                    ((DataGridView)sender).Columns["GeneralDescription25"].Visible = false;
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="StartingInCell"></param>
        /// <param name="LineIndicator"></param>
        /// <param name="list"></param>
        /// <param name="ws"></param>
        /// <returns></returns>
        public List<ExcelAddIn4.Common2.Specialist> GetEntityListFromDGVForTransUpd(string StartingInCell, string LineIndicator, List<ExcelAddIn4.Common2.Specialist> list, Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            List<ExcelAddIn4.Common2.Specialist> newlist = new List<ExcelAddIn4.Common2.Specialist>();
            try
            {
                string StartColumn = Finance_Tools.RemoveNumber(StartingInCell);
                string StartRow = Finance_Tools.RemoveNotNumber(StartingInCell);
                Microsoft.Office.Interop.Excel.Range y2;
                Microsoft.Office.Interop.Excel.Range x;

                if (StartingInCell.Replace("$", "").Contains("A"))
                {
                    y2 = Globals.ThisAddIn.Application.Rows.get_Range("A1");
                    x = Globals.ThisAddIn.Application.Rows.get_Range(StartingInCell.Replace("$", ""));
                }
                else
                {
                    y2 = Globals.ThisAddIn.Application.Rows.get_Range("A1");
                    x = Globals.ThisAddIn.Application.Rows.get_Range(StartingInCell.Replace("$", "")).Previous;
                }
                string ss = "";
                int num = 0;
                int startRowNumber = int.Parse(StartRow);
                int count = 0;
                while (ss != StartColumn || (num < startRowNumber))
                {
                    if (count > 10000)
                    {
                        return null;
                    }
                    try
                    {
                        if (ss == "" || ss == "A")
                        {
                            x = Globals.ThisAddIn.Application.Rows.Find(LineIndicator, Globals.ThisAddIn.Application.Rows.get_Range("A1"), Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, Microsoft.Office.Interop.Excel.XlLookAt.xlWhole, Type.Missing, Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, true, true, Type.Missing);
                        }
                        else
                        {
                            x = Globals.ThisAddIn.Application.Rows.Find(LineIndicator, x, Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, Microsoft.Office.Interop.Excel.XlLookAt.xlWhole, Type.Missing, Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, true, true, Type.Missing);
                        }
                        ss = Finance_Tools.RemoveNumber(x.Address.Replace("$", ""));
                        num = int.Parse(Finance_Tools.RemoveNotNumber(x.Address.Replace("$", "")));
                    }
                    catch
                    {
                        if (Ribbon2.tupf.Visible == true)
                            TransUpdPostFrm.richTextBox1.Text = "No data found for specified Line Indicator(s) -" + LineIndicator + "!";
                        else
                            SessionInfo.UserInfo.GlobalError += "Process:Journal Update(" + SessionInfo.UserInfo.CurrentRef + ") - Fail: No data found for specified Line Indicator(s) -(" + LineIndicator + ")!";
                        return null;
                    }
                    count++;
                }

                string ss2 = "";
                int num2 = 0;
                count = 0;
                Microsoft.Office.Interop.Excel.Range LastRow = Globals.ThisAddIn.Application.Range["A1", StartColumn + Ribbon2.LastRowNumber].Find(LineIndicator, y2, Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, Microsoft.Office.Interop.Excel.XlLookAt.xlWhole, Type.Missing, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, true, true, Type.Missing);
                while (ss2 != StartColumn || (num2 < startRowNumber))
                {
                    if (count > 10000)
                    {
                        return null;
                    }
                    try
                    {
                        if (ss2 == "" || ss2 == "A")
                        {
                            LastRow = Globals.ThisAddIn.Application.Range["A1", StartColumn + Ribbon2.LastRowNumber].Find(LineIndicator, y2, Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, Microsoft.Office.Interop.Excel.XlLookAt.xlWhole, Type.Missing, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, true, true, Type.Missing);
                        }
                        else
                        {
                            LastRow = Globals.ThisAddIn.Application.Range["A1", StartColumn + Ribbon2.LastRowNumber].Find(LineIndicator, LastRow, Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, Microsoft.Office.Interop.Excel.XlLookAt.xlWhole, Type.Missing, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, true, true, Type.Missing);
                        }
                        ss2 = Finance_Tools.RemoveNumber(LastRow.Address.Replace("$", ""));
                        num2 = int.Parse(Finance_Tools.RemoveNotNumber(LastRow.Address.Replace("$", "")));
                    }
                    catch
                    {
                        if (Ribbon2.tupf.Visible == true)
                            TransUpdPostFrm.richTextBox1.Text = "No data found for specified Line Indicator(s) -" + LineIndicator + "!";
                        else
                            SessionInfo.UserInfo.GlobalError += "Process:Journal Update(" + SessionInfo.UserInfo.CurrentRef + ") - Fail: No data found for specified Line Indicator(s) -(" + LineIndicator + ")!";
                        return null;
                    }
                    count++;
                }
                string lastRowNum = Finance_Tools.RemoveNotNumber(LastRow.Address);
                string s = Finance_Tools.RemoveNotNumber(x.Address);
                MyServices = new ServiceContainer();
                engine = new FormulaEngine();
                SetFormulaEngine(engine);
                this.BigCalculateTransUpd(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist);
                return newlist;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + ex.InnerException + "," + ex.StackTrace);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="StartingInCell"></param>
        /// <param name="LineIndicator"></param>
        /// <param name="list"></param>
        /// <param name="ws"></param>
        /// <returns></returns>
        public List<RowCreateTextFile> GetEntityListFromDGVForCreateTextFile(string StartingInCell, string LineIndicator, List<RowCreateTextFile> list, Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            List<RowCreateTextFile> newlist = new List<RowCreateTextFile>();
            try
            {
                string StartColumn = Finance_Tools.RemoveNumber(StartingInCell);
                string StartRow = Finance_Tools.RemoveNotNumber(StartingInCell);
                Microsoft.Office.Interop.Excel.Range y2;
                Microsoft.Office.Interop.Excel.Range x;
                if (StartingInCell.Replace("$", "").Contains("A"))
                {
                    y2 = Globals.ThisAddIn.Application.Rows.get_Range("A1");
                    x = Globals.ThisAddIn.Application.Rows.get_Range(StartingInCell.Replace("$", ""));
                }
                else
                {
                    y2 = Globals.ThisAddIn.Application.Rows.get_Range("A1");
                    x = Globals.ThisAddIn.Application.Rows.get_Range(StartingInCell.Replace("$", "")).Previous;
                }
                string ss = "";
                int num = 0;
                int startRowNumber = int.Parse(StartRow);
                int count = 0;
                while (ss != StartColumn || (num < startRowNumber))
                {
                    if (count > 10000)
                    {
                        return null;
                    }
                    try
                    {
                        if (ss == "" || ss == "A")
                        {
                            x = Globals.ThisAddIn.Application.Rows.Find(LineIndicator, Globals.ThisAddIn.Application.Rows.get_Range("A1"), Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, Microsoft.Office.Interop.Excel.XlLookAt.xlWhole, Type.Missing, Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, true, true, Type.Missing);
                        }
                        else
                        {
                            x = Globals.ThisAddIn.Application.Rows.Find(LineIndicator, x, Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, Microsoft.Office.Interop.Excel.XlLookAt.xlWhole, Type.Missing, Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, true, true, Type.Missing);
                        }
                        ss = Finance_Tools.RemoveNumber(x.Address.Replace("$", ""));
                        num = int.Parse(Finance_Tools.RemoveNotNumber(x.Address.Replace("$", "")));
                    }
                    catch
                    {
                        if (Ribbon2.ctff.Visible == true)
                            CreateTextFileForm.richTextBox1.Text = "No data found for specified Line Indicator(s) -" + LineIndicator + "!";
                        else
                            SessionInfo.UserInfo.GlobalError += "Process:" + (SessionInfo.UserInfo.ComName == "" ? SessionInfo.UserInfo.Textfilename : SessionInfo.UserInfo.ComName + " " + SessionInfo.UserInfo.MethodName) + "(" + SessionInfo.UserInfo.CurrentRef + ") - Fail: No data found for specified Line Indicator(s) -(" + LineIndicator + ")!";
                        return null;
                    }
                    count++;
                }
                string ss2 = "";
                int num2 = 0;
                count = 0;
                Microsoft.Office.Interop.Excel.Range LastRow = Globals.ThisAddIn.Application.Range["A1", StartColumn + Ribbon2.LastRowNumber].Find(LineIndicator, y2, Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, Microsoft.Office.Interop.Excel.XlLookAt.xlWhole, Type.Missing, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, true, true, Type.Missing);
                while (ss2 != StartColumn || (num2 < startRowNumber))
                {
                    if (count > 10000)
                    {
                        return null;
                    }
                    try
                    {
                        if (ss2 == "" || ss2 == "A")
                        {
                            LastRow = Globals.ThisAddIn.Application.Range["A1", StartColumn + Ribbon2.LastRowNumber].Find(LineIndicator, y2, Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, Microsoft.Office.Interop.Excel.XlLookAt.xlWhole, Type.Missing, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, true, true, Type.Missing);
                        }
                        else
                        {
                            LastRow = Globals.ThisAddIn.Application.Range["A1", StartColumn + Ribbon2.LastRowNumber].Find(LineIndicator, LastRow, Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, Microsoft.Office.Interop.Excel.XlLookAt.xlWhole, Type.Missing, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, true, true, Type.Missing);
                        }
                        ss2 = Finance_Tools.RemoveNumber(LastRow.Address.Replace("$", ""));
                        num2 = int.Parse(Finance_Tools.RemoveNotNumber(LastRow.Address.Replace("$", "")));
                    }
                    catch
                    {
                        if (Ribbon2.ctff.Visible == true)
                            CreateTextFileForm.richTextBox1.Text = "No data found for specified Line Indicator(s) -" + LineIndicator + "!";
                        else
                            SessionInfo.UserInfo.GlobalError += "Process:" + (SessionInfo.UserInfo.ComName == "" ? SessionInfo.UserInfo.Textfilename : SessionInfo.UserInfo.ComName + " " + SessionInfo.UserInfo.MethodName) + "(" + SessionInfo.UserInfo.CurrentRef + ")  - Fail : No data found for specified Line Indicator(s) -(" + LineIndicator + ")!";
                        return null;
                    }
                    count++;
                }
                string lastRowNum = Finance_Tools.RemoveNotNumber(LastRow.Address);
                string s = Finance_Tools.RemoveNotNumber(x.Address);
                MyServices = new ServiceContainer();
                engine = new FormulaEngine();
                SetFormulaEngine(engine);
                this.BigCalculateCTF(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist);
                return newlist;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + ex.InnerException + "," + ex.StackTrace);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="StartingInCell"></param>
        /// <param name="LineIndicator"></param>
        /// <param name="list"></param>
        /// <param name="ws"></param>
        /// <returns></returns>
        public List<Specialist> GetEntityListFromDGV(string StartingInCell, string LineIndicator, List<Specialist> list, Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            List<Specialist> newlist = new List<Specialist>();
            try
            {
                //first step: get row count of "Line Indicator" 
                string StartColumn = Finance_Tools.RemoveNumber(StartingInCell);
                string StartRow = Finance_Tools.RemoveNotNumber(StartingInCell);
                Microsoft.Office.Interop.Excel.Range y2;
                Microsoft.Office.Interop.Excel.Range x;

                if (StartingInCell.Replace("$", "").Contains("A"))
                {
                    y2 = Globals.ThisAddIn.Application.Rows.get_Range("A1");
                    x = Globals.ThisAddIn.Application.Rows.get_Range(StartingInCell.Replace("$", ""));
                }
                else
                {
                    y2 = Globals.ThisAddIn.Application.Rows.get_Range("A1");
                    x = Globals.ThisAddIn.Application.Rows.get_Range(StartingInCell.Replace("$", "")).Previous;
                }
                string ss = "";
                int num = 0;
                int startRowNumber = int.Parse(StartRow);
                int count = 0;
                while (ss != StartColumn || (num < startRowNumber))
                {
                    if (count > 10000)
                    {
                        return null;
                    }
                    try
                    {
                        if (ss == "" || ss == "A")
                        {
                            x = Globals.ThisAddIn.Application.Rows.Find(LineIndicator, Globals.ThisAddIn.Application.Rows.get_Range("A1"), Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, Microsoft.Office.Interop.Excel.XlLookAt.xlWhole, Type.Missing, Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, true, true, Type.Missing);
                        }
                        else
                        {
                            x = Globals.ThisAddIn.Application.Rows.Find(LineIndicator, x, Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, Microsoft.Office.Interop.Excel.XlLookAt.xlWhole, Type.Missing, Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext, true, true, Type.Missing);
                        }
                        ss = Finance_Tools.RemoveNumber(x.Address.Replace("$", ""));
                        num = int.Parse(Finance_Tools.RemoveNotNumber(x.Address.Replace("$", "")));
                    }
                    catch
                    {
                        if (Ribbon2.xpf.Visible == true)
                            XMLPostFrm.richTextBox1.Text = "No data found for specified Line Indicator(s) -" + LineIndicator + "!";
                        else
                            SessionInfo.UserInfo.GlobalError += "Process:Journal Post(" + SessionInfo.UserInfo.CurrentRef + ") - Fail: No data found for specified Line Indicator(s) -(" + LineIndicator + ")!";

                        return null;
                    }
                    count++;
                }
                string ss2 = "";
                int num2 = 0;
                count = 0;
                Microsoft.Office.Interop.Excel.Range LastRow = Globals.ThisAddIn.Application.Range["A1", StartColumn + Ribbon2.LastRowNumber].Find(LineIndicator, y2, Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, Microsoft.Office.Interop.Excel.XlLookAt.xlWhole, Type.Missing, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, true, true, Type.Missing);
                while (ss2 != StartColumn || (num2 < startRowNumber))
                {
                    if (count > 10000)
                    {
                        return null;
                    }
                    try
                    {
                        if (ss2 == "" || ss2 == "A")
                        {
                            LastRow = Globals.ThisAddIn.Application.Range["A1", StartColumn + Ribbon2.LastRowNumber].Find(LineIndicator, y2, Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, Microsoft.Office.Interop.Excel.XlLookAt.xlWhole, Type.Missing, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, true, true, Type.Missing);
                        }
                        else
                        {
                            LastRow = Globals.ThisAddIn.Application.Range["A1", StartColumn + Ribbon2.LastRowNumber].Find(LineIndicator, LastRow, Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues, Microsoft.Office.Interop.Excel.XlLookAt.xlWhole, Type.Missing, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, true, true, Type.Missing);
                        }
                        ss2 = Finance_Tools.RemoveNumber(LastRow.Address.Replace("$", ""));
                        num2 = int.Parse(Finance_Tools.RemoveNotNumber(LastRow.Address.Replace("$", "")));
                    }
                    catch
                    {
                        if (Ribbon2.xpf.Visible == true)
                            XMLPostFrm.richTextBox1.Text = "No data found for specified Line Indicator(s) -" + LineIndicator + "!";
                        else
                            SessionInfo.UserInfo.GlobalError += "Process:Journal Post(" + SessionInfo.UserInfo.CurrentRef + ") - Fail: No data found for specified Line Indicator(s) -(" + LineIndicator + ")!";

                        return null;
                    }
                    count++;
                }
                string lastRowNum = Finance_Tools.RemoveNotNumber(LastRow.Address);
                //second step: match DGV Row to Workbook Row into Entity.
                string s = Finance_Tools.RemoveNotNumber(x.Address);
                MyServices = new ServiceContainer();
                engine = new FormulaEngine();
                SetFormulaEngine(engine);
                this.BigCalculate(lastRowNum, list, StartColumn, LineIndicator, s, ws, ref newlist);
                return newlist;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + ex.InnerException + "," + ex.StackTrace);
            }
        }
        ///// <summary>
        ///// distinct F15 TO $F$15 
        ///// </summary>
        ///// <param name="input"></param>
        ///// <param name="RowNum"></param>
        ///// <returns></returns>
        //private string GetActualEntityValue(string input, int RowNum)
        //{
        //    string str2 = ReplaceOperator(input);
        //    string str3 = input;
        //    string[] arr = str2.Split(',');

        //    foreach (string arrStr in arr)
        //    {
        //        string rArrStr = Finance_Tools.RemoveNumber(arrStr);
        //        string rnArrStr = Finance_Tools.RemoveNotNumber(arrStr);
        //        if (Finance_Tools.IsLetterOrDigit(arrStr) && rArrStr.Length > 0 && rnArrStr.Length > 0 && rArrStr.Length < 3)
        //        {
        //            string tmp2 = rArrStr + (int.Parse(rnArrStr) + RowNum).ToString();
        //            str3 = str3.Replace(arrStr, tmp2);
        //        }
        //    }
        //    return str3.Replace("$", "");
        //}
        ///// <summary>
        ///// 
        ///// </summary>
        ///// <param name="str"></param>
        ///// <returns></returns>
        //private string ReplaceOperator(string str)
        //{
        //    return str.Replace("(", ",").Replace(")", ",").Replace("+", ",").Replace("-", ",").Replace("*", ",").Replace("/", ",");
        //}
        ///// <summary>
        ///// 
        ///// </summary>
        ///// <param name="Char"></param>
        ///// <returns></returns>
        //private bool isOperator(string Char)
        //{
        //    bool returnValue = false;

        //    switch (Char)
        //    {
        //        case "+":
        //            returnValue = true;
        //            break;
        //        case "-":
        //            returnValue = true;
        //            break;
        //        case "*":
        //            returnValue = true;
        //            break;
        //        case "/":
        //            returnValue = true;
        //            break;
        //        default:
        //            break;
        //    }
        //    return returnValue;
        //}
        /// <summary>
        /// 
        /// </summary>
        /// <param name="input"></param>
        /// <param name="RowNum"></param>
        /// <param name="ws"></param>
        /// <returns></returns>
        public string GetEntityValueInExcel(string input, int RowNum, Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            string outPut = string.Empty;
            //Globals.ThisAddIn.Application.Cells.get_Range("BT1000").Clear();
            //Microsoft.Office.Interop.Excel.WorksheetClass workSheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value) as Microsoft.Office.Interop.Excel.WorksheetClass;
            ws.Cells.get_Range("BZ1000").Clear();

            //string str2 = GetActualEntityValue(input, RowNum);
            //GetActualEntityValue(input, RowNum); start write here for high up efficiency
            //string str2 = ReplaceOperator(input);
            string str2 = input.Replace("(", ",").Replace(")", ",").Replace("+", ",").Replace("-", ",").Replace("*", ",").Replace("/", ",");
            string str3 = input;
            string[] arr = str2.Split(',');

            foreach (string arrStr in arr)
            {
                string rArrStr = Regex.Replace(arrStr, @"\d", "");
                string rnArrStr = Regex.Replace(arrStr, @"[^\d]*", "");
                if (Finance_Tools.IsLetterOrDigit(arrStr) && rArrStr.Length > 0 && rnArrStr.Length > 0 && rArrStr.Length < 3)
                {
                    string tmp2 = rArrStr + (int.Parse(rnArrStr) + RowNum).ToString();
                    str3 = str3.Replace(arrStr, tmp2);
                }
            }
            str2 = str3.Replace("$", "");
            //end

            try
            {
                ws.Cells[1000, 78] = (str2.IndexOf("=") != -1 ? str2 : "=" + str2);
                bool error = ws.Cells.get_Range("BZ1000").Errors.Item[1].Value;
                if (error || string.IsNullOrEmpty(str2))
                {
                    outPut = input;
                }
                else
                {
                    outPut = ws.Cells[1000, 78] == null ? str2 : ws.Cells.get_Range("BZ1000").Value.ToString();
                }
            }
            catch
            {
                outPut = input;
            }

            return outPut;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private bool isFormulaPeriod(string str)
        {
            try
            {
                bool result = Finance_Tools.IsFormulaPeriodString(str);
                if (result)
                {
                    string y3 = str.Replace("/", "").Substring(0, 4);
                    string m = str.Replace("/", "").Substring(4, 3);
                    int a = 0;
                    if ((int.TryParse(y3, out a) == false) || (int.TryParse(m, out a) == false))
                        return false;
                    else
                        return true;
                }
                else
                {
                    return false;
                }
            }
            catch { return false; }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="input"></param>
        /// <param name="startRowNum"></param>
        /// <param name="RowNum"></param>
        /// <param name="ws"></param>
        /// <returns></returns>
        public string GetEntityValueInExcel(string input, string startRowNum, int RowNum, Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            string outPut = string.Empty;
            //Globals.ThisAddIn.Application.Cells.get_Range("BT1000").Clear();
            //Microsoft.Office.Interop.Excel.WorksheetClass workSheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value) as Microsoft.Office.Interop.Excel.WorksheetClass;
            //ws.Cells.get_Range("BZ1000").Clear();

            //string str2 = GetActualEntityValue(input, RowNum);
            //GetActualEntityValue(input, RowNum); start write here for high up efficiency
            //string str2 = ReplaceOperator(input);
            string str2 = input.Replace("(", ",").Replace(")", ",").Replace("+", ",").Replace("-", ",").Replace("*", ",").Replace("/", ",").Replace("^", ",").Replace("&", ",").Replace(">=", ",").Replace("<=", ",").Replace("<>", ",").Replace("%", ",").Replace("<", ",").Replace("=", ",").Replace(">", ",").Replace("\'", ",");
            string str3 = input;
            string[] arr = str2.Split(',');

            foreach (string arrStr in arr)
            {
                string rArrStr = Regex.Replace(arrStr, @"\d", "");
                string rnArrStr = Regex.Replace(arrStr, @"[^\d]*", "");
                string tmp2 = rArrStr + (int.Parse(startRowNum) + RowNum).ToString();
                if (Finance_Tools.IsLetterOrDigit(arrStr) && rArrStr.Length > 0 && rnArrStr.Length > 0 && rArrStr.Length < 3)
                {
                    //string tmp2 = rArrStr + (int.Parse(rnArrStr) + RowNum).ToString();
                    try
                    {
                        tmp2 = ws.Cells.get_Range(tmp2).Value.ToString();
                    }
                    catch { tmp2 = ""; }
                    str3 = str3.Replace(arrStr, tmp2);
                }
                else if (arrStr.Contains("$"))
                {
                    try
                    {
                        tmp2 = ws.Cells.get_Range(arrStr).Value.ToString();
                    }
                    catch { tmp2 = ""; }
                    str3 = str3.Replace(arrStr, tmp2);
                }
            }
            str2 = str3;
            //end
            if (!str2.Contains("(") && !str2.Contains(")") && !str2.Contains("+") && !str2.Contains("-") && !str2.Contains("*") && !str2.Contains("/") && !str2.Contains("^") && !str2.Contains("&") && !str2.Contains(">=") && !str2.Contains("<=") && !str2.Contains("<>") && !str2.Contains("%") && !str2.Contains("<") && !str2.Contains("=") && !str2.Contains(">"))
            {
                outPut = str2;
            }
            else
            {
                if (str2.Contains("/") && isFormulaPeriod(str2))
                    outPut = str2;
                else
                {
                    try
                    {
                        string expression = (str2.IndexOf("=") != -1 ? str2 : str2).Replace("\'", "\"");
                        Formula f = CreateFormula(expression);
                        object result = null;
                        if (f != null)
                        {
                            result = f.Evaluate();
                            outPut = result.ToString();
                        }
                        else
                        {
                            outPut = str2;
                        }
                        //ws.Cells[1000, 78] = (str2.IndexOf("=") != -1 ? str2 : "=" + str2);
                        //bool error = ws.Cells.get_Range("BZ1000").Errors.Item[1].Value;
                        //if (error || string.IsNullOrEmpty(str2))
                        //{
                        //    outPut = input;
                        //}
                        //else
                        //{
                        //    outPut = ws.Cells[1000, 78] == null ? str2 : ws.Cells.get_Range("BZ1000").Value.ToString();
                        //}
                    }
                    catch
                    {
                        outPut = input;
                    }
                }
            }
            return outPut;
            //string outPut = string.Empty;
            ////Globals.ThisAddIn.Application.Cells.get_Range("BT1000").Clear();
            ////Microsoft.Office.Interop.Excel.WorksheetClass workSheet = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value) as Microsoft.Office.Interop.Excel.WorksheetClass;
            //ws.Cells.get_Range("BZ1000").Clear();

            ////string str2 = GetActualEntityValue(input, RowNum);
            ////GetActualEntityValue(input, RowNum); start write here for high up efficiency
            ////string str2 = ReplaceOperator(input);
            //string str2 = input.Replace("(", ",").Replace(")", ",").Replace("+", ",").Replace("-", ",").Replace("*", ",").Replace("/", ",");
            //string str3 = input;
            //string[] arr = str2.Split(',');

            //foreach (string arrStr in arr)
            //{
            //    string rArrStr = Regex.Replace(arrStr, @"\d", "");
            //    string rnArrStr = Regex.Replace(arrStr, @"[^\d]*", "");
            //    if (Finance_Tools.IsLetterOrDigit(arrStr) && rArrStr.Length > 0 && rnArrStr.Length > 0 && rArrStr.Length < 3)
            //    {
            //        //string tmp2 = rArrStr + (int.Parse(rnArrStr) + RowNum).ToString();
            //        string tmp2 = rArrStr + (int.Parse(startRowNum) + RowNum).ToString();
            //        str3 = str3.Replace(arrStr, tmp2);
            //    }
            //}
            //str2 = str3.Replace("$", "");
            ////end

            //try
            //{
            //    ws.Cells[1000, 78] = (str2.IndexOf("=") != -1 ? str2 : "=" + str2);
            //    bool error = ws.Cells.get_Range("BZ1000").Errors.Item[1].Value;
            //    if (error || string.IsNullOrEmpty(str2))
            //    {
            //        outPut = input;
            //    }
            //    else
            //    {
            //        outPut = ws.Cells[1000, 78] == null ? str2 : ws.Cells.get_Range("BZ1000").Value.ToString();
            //    }
            //}
            //catch
            //{
            //    outPut = input;
            //}

            //return outPut;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="startRow"></param>
        /// <param name="currentRow"></param>
        /// <returns></returns>
        private bool IsGreaterThanStartRow(string startRow, int currentRow)
        {
            try
            {
                int defaultRow = int.Parse(startRow);
                if (currentRow >= defaultRow)
                    return true;
                else
                    return false;
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="list"></param>
        /// <param name="str"></param>
        /// <returns></returns>
        private bool IsInColumns(int column, int str)
        {
            try
            {
                if (column == str)
                    return true;
                else
                    return false;
            }
            catch
            {
                throw new Exception("Starting in Cell or Line Indicator Error!");
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private bool IsExcelRowContainCharacter(string str)
        {
            try
            {
                string iniAddress = str;
                string tmp = string.Empty;
                for (int i = 0; i < Finance_Tools.MaxColumnCount; i++)
                {
                    var x = Globals.ThisAddIn.Application.Rows.get_Range(str.Replace("$", "")).Next;
                    str = x.Address;
                    tmp += x.Value;
                }
                if (string.IsNullOrEmpty(tmp))
                {
                    for (int i = 0; i < Finance_Tools.MaxColumnCount; i++)
                    {
                        if (!iniAddress.Replace("$", "").Contains("A"))
                        {
                            var x = Globals.ThisAddIn.Application.Rows.get_Range(iniAddress.Replace("$", "")).Previous;
                            iniAddress = x.Address;
                            tmp += x.Value;
                        }
                    }
                    if (string.IsNullOrEmpty(tmp)) return false;
                    else return true;
                }
                else return true;
            }
            catch
            {
                throw new Exception("Starting in Cell or Line Indicator Error!");
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="StartingCell"></param>
        /// <param name="LineIndicator"></param>
        /// <returns></returns>
        private bool IsStartingCellRowContainLineIndicator(string StartingCell, string LineIndicator)
        {
            try
            {
                string tmp = string.Empty;
                var x = Globals.ThisAddIn.Application.Rows.get_Range(StartingCell.Replace("$", ""));
                tmp += x.Value;
                if (tmp == LineIndicator) return true;
                else return false;
            }
            catch
            {
                throw new Exception("Starting in Cell or Line Indicator Error!");
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public string GetValueOfAddress(string address)
        {
            string tmp = string.Empty;
            var x = Globals.ThisAddIn.Application.Rows.get_Range(address.Replace("$", ""));
            tmp += x.Value;
            return tmp;
        }
        /// <summary>
        ///     //                         <ErrorContext>" +
        //                   "       <CompatibilityMode>0</CompatibilityMode>" +
        //                   "       <ErrorOutput>1</ErrorOutput>" +
        //                   "       <ErrorThreshold>1</ErrorThreshold>" +
        //                   "   </ErrorContext>" +
        //                   "   <SunSystemsContext>" +
        //                   "       <BusinessUnit>PK1</BusinessUnit>" +
        //                   "       <BudgetCode>A</BudgetCode>" +
        //                   "   </SunSystemsContext>" +
        //                   "   <MethodContext>" +
        //                   "       <LedgerPostingParameters>" +
        //                   "           <DefaultPeriod>0122012</DefaultPeriod>" +
        //                   "           <Description>TestYB</Description>" +
        //                   "           <JournalType>PI</JournalType>" +
        //                   "           <LoadOnly>N</LoadOnly>" +
        //                   "           <PostProvisional>N</PostProvisional>" +
        //                   "           <PostToHold>N</PostToHold>" +
        //                   "           <PostingType>2</PostingType>" +
        //                   "           <ReportErrorsOnly>Y</ReportErrorsOnly>" +
        //                   "           <ReportingAccount>999</ReportingAccount>" +
        //                   "           <SuppressSubstitutedMessages>Y</SuppressSubstitutedMessages>" +
        //                   "           <SuspenseAccount>999</SuspenseAccount>" +
        //                   "           <TransactionAmountAccount>999</TransactionAmountAccount>" +
        //                   "       </LedgerPostingParameters>" +
        //                   "   </MethodContext>" +
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public string GetXMLScript(List<Line> list)
        {
            try
            {
                SSC de = new SSC();
                de.Payload = new Payload();
                de.ErrorContext = new ErrorContext();
                de.ErrorContext.CompatibilityMode = "1";
                de.ErrorContext.ErrorOutput = "1";
                de.ErrorContext.ErrorThreshold = "1";
                de.SunSystemsContext = new SunSystemsContext();
                de.SunSystemsContext.BusinessUnit = Finance_Tools.GetAppConfig("BusinessUnit");

                if (list.Count != 0 && !string.IsNullOrEmpty(list[0].Ledger))
                {
                    de.SunSystemsContext.BudgetCode = list[0].Ledger;
                }
                else
                {
                    de.SunSystemsContext.BudgetCode = "";
                }
                de.MethodContext = new MethodContext();
                de.MethodContext.LedgerPostingParamrters = new LedgerPostingParameters();
                de.MethodContext.LedgerPostingParamrters.AllowBalTran = SessionInfo.UserInfo.AllowBalTran;
                de.MethodContext.LedgerPostingParamrters.AllowPostToSuspended = SessionInfo.UserInfo.AllowPostToSuspended;
                de.MethodContext.LedgerPostingParamrters.DefaultPeriod = list[0].AccountingPeriod;
                de.MethodContext.LedgerPostingParamrters.Description = "Finance Tools " + DateTime.Now.ToString("dd/MM/yyyy");

                if (list.Count != 0 && !string.IsNullOrEmpty(list[0].JournalType))
                {
                    de.MethodContext.LedgerPostingParamrters.JournalType = list[0].JournalType;
                }
                else
                {
                    de.MethodContext.LedgerPostingParamrters.JournalType = "";
                }
                de.MethodContext.LedgerPostingParamrters.LoadOnly = "N";
                de.MethodContext.LedgerPostingParamrters.PostProvisional = "N";
                de.MethodContext.LedgerPostingParamrters.PostToHold = "N";
                de.MethodContext.LedgerPostingParamrters.PostingType = "2";
                de.MethodContext.LedgerPostingParamrters.ReportErrorsOnly = "Y";
                de.MethodContext.LedgerPostingParamrters.ReportingAccount = Finance_Tools.GetAppConfig("SuspenseAccount");
                de.MethodContext.LedgerPostingParamrters.SuppressSubstitutedMessages = "Y";
                de.MethodContext.LedgerPostingParamrters.SuspenseAccount = Finance_Tools.GetAppConfig("SuspenseAccount");
                de.MethodContext.LedgerPostingParamrters.TransactionAmountAccount = Finance_Tools.GetAppConfig("SuspenseAccount");
                de.Payload.Ledger = list;
                de.ErrorMessages = ".";
                string str = XmlSerialization<SSC>.Serialize(de).Replace("<?xml version=\"1.0\"?>", "<?xml version=\"1.0\" encoding=\"UTF-8\"?>").Replace("<SSC xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">", "<SSC>");
                return str;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + "," + ex.InnerException + "," + ex.StackTrace);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="elem"></param>
        /// <returns></returns>
        private bool EquaWithName(KeyValuePair<string, int> elem)
        {
            if (elem.Value == tmpStr)
                return true;
            return false;
        }
        int tmpStr;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="elem"></param>
        /// <returns></returns>
        private bool EquaWithHeaderName(KeyValuePair<string, string> elem)
        {
            if (elem.Key == tmpStrHeader)
                return true;
            return false;
        }
        /// <summary>
        /// 
        /// </summary>
        string tmpStrHeader;
        private int invisibleFieldExist(string field, int id, string com, string method)
        {
            int count = (from XMLorTextFile in db.rsTemplateCreateXMLTextProfiles
                         where XMLorTextFile.SunComponentName == com && XMLorTextFile.SunMethod == method
                         && XMLorTextFile.Section == "Line" && XMLorTextFile.Field == field && XMLorTextFile.ID < id && (XMLorTextFile.Mandatory == true && XMLorTextFile.Visible == false)
                         select XMLorTextFile).Count();
            return count;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="comName"></param>
        /// <param name="methodName"></param>
        /// <param name="field"></param>
        /// <returns></returns>
        public int GetXMLorTextFileSameFieldCount(string comName, string methodName, string field)
        {
            int count = (from XMLorTextFile in db.rsTemplateCreateXMLTextProfiles
                         where XMLorTextFile.SunComponentName == comName && XMLorTextFile.SunMethod == methodName
                         && XMLorTextFile.Section == "Line" && XMLorTextFile.Field == field && !(XMLorTextFile.Mandatory == false && XMLorTextFile.Visible == false)
                         select XMLorTextFile).Count();
            return count;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="payload"></param>
        /// <param name="outXML"></param>
        /// <param name="item"></param>
        /// <param name="list"></param>
        /// <param name="dtheader"></param>
        /// <param name="dtLine"></param>
        /// <param name="bu"></param>
        /// <param name="sa"></param>
        /// <param name="objlist"></param>
        /// <param name="firstLine"></param>
        /// <returns></returns>
        private string GeneLinesByHeader(string payload, string outXML, KeyValuePair<string, int> item, List<RowCreateTextFile> list, DataTable dtheader, DataTable dtLine, string bu, string sa, List<KeyValuePair<string, int>> objlist, string firstLine, string com, string method)
        {
            int count = 1;
            XmlDocument xdoc = new XmlDocument();
            xdoc.LoadXml(payload);
            string headerXML = string.Empty;
            string wholeXML = "";
            int x = 0;
            List<RowCreateTextFile> Linelist = new List<RowCreateTextFile>();
            foreach (KeyValuePair<string, int> obj in objlist)
                if (obj.Key == item.Key)
                    Linelist.Add(list[obj.Value]);

            List<string> replist = new List<string>();
            for (int k = 0; k < Linelist.Count; k++)
            {
                for (int j = 0; j < dtLine.Rows.Count; j++)
                {
                    int existCount = invisibleFieldExist(dtLine.Rows[j]["Field"].ToString(), int.Parse(dtLine.Rows[j]["ID"].ToString()), com, method);
                    if (existCount > 0)
                        x = existCount + replist.Where(a => a.ToString() == dtLine.Rows[j]["Field"].ToString()).Count();
                    else
                        x = replist.Where(a => a.ToString() == dtLine.Rows[j]["Field"].ToString()).Count();
                    XmlNode Nodes = xdoc.GetElementsByTagName(dtLine.Rows[j]["Field"].ToString())[x];
                    bool Mandatory = bool.Parse(dtLine.Rows[j]["Mandatory"].ToString());
                    if (Mandatory == false && CreateTextFileForm.removedColumns.Contains(dtLine.Rows[j]["Field"].ToString() + "Column" + (count + dtheader.Rows.Count) + ","))
                    {
                        Nodes.ParentNode.RemoveChild(Nodes);
                    }
                    else if (Nodes != null && (Nodes.ChildNodes.Count == 0 || Nodes.LastChild.NodeType == XmlNodeType.Text))
                    {
                        Nodes.InnerText = DataConversionTools.GetPropertyValue("Column" + (count + dtheader.Rows.Count), Linelist[k]);
                        replist.Add(dtLine.Rows[j]["Field"].ToString());
                    }
                    count++;
                }
                count = 1;
                x = 0;
                replist.Clear();
                wholeXML += xdoc.InnerXml;
                xdoc = new XmlDocument();
                xdoc.LoadXml(payload);
            }
            xdoc = new XmlDocument();
            xdoc.LoadXml(outXML);
            count = 1;
            for (int j = 0; j < dtheader.Rows.Count; j++)
            {
                XmlNode Nodes = xdoc.GetElementsByTagName(dtheader.Rows[j]["Field"].ToString())[0];
                bool Mandatory = bool.Parse(dtheader.Rows[j]["Mandatory"].ToString());
                if (Mandatory == false && CreateTextFileForm.removedColumns.Contains(dtheader.Rows[j]["Field"].ToString() + "Column" + count + ","))
                {
                    Nodes.ParentNode.RemoveChild(Nodes);
                }
                else if (Nodes.ChildNodes.Count == 0 || Nodes.LastChild.NodeType == XmlNodeType.Text)
                {
                    Nodes.InnerText = DataConversionTools.GetPropertyValue("Column" + count, Linelist[0]);
                }
                count++;
            }
            headerXML += xdoc.InnerXml;
            xdoc = new XmlDocument();
            xdoc.LoadXml(headerXML);
            XmlNode Nodes2 = xdoc.GetElementsByTagName(firstLine)[0];
            Nodes2.InnerXml = wholeXML;
            return FormatXml(xdoc.InnerXml);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public string GetXMLProfileScript(List<RowCreateTextFile> list)
        {
            try
            {
                string bu = Finance_Tools.GetAppConfig("BusinessUnit");
                string sa = Finance_Tools.GetAppConfig("SuspenseAccount");
                List<KeyValuePair<string, int>> objlist = new List<KeyValuePair<string, int>>();
                List<KeyValuePair<string, string>> headerlist = new List<KeyValuePair<string, string>>();
                string xml = GetXMLFileTemplate(list[0].SunComponent + "," + list[0].SunMethod);
                DataTable dt = GetXMLorTextFileVFieldsByComName(list[0].SunComponent, list[0].SunMethod);
                DataTable dtHeader = GetXMLorTextFileHeaderByComName(list[0].SunComponent, list[0].SunMethod);
                DataTable dtLine = GetXMLorTextFileLineByComName(list[0].SunComponent, list[0].SunMethod);
                string FirstLine = GetFirstLineByComName(list[0].SunComponent, list[0].SunMethod);
                Predicate<KeyValuePair<string, int>> pred = EquaWithName;
                Predicate<KeyValuePair<string, string>> pred2 = EquaWithHeaderName;
                for (int k = 0; k < list.Count; k++)                                //get all header value lines into dictionary
                    for (int j = 0; j < dt.Rows.Count; j++)
                        if (dt.Rows[j]["Section"].ToString() == "Header")
                        {
                            tmpStr = k;
                            if (objlist.Exists(pred))
                            {
                                string key = objlist[objlist.Count - 1].Key + "," + DataConversionTools.GetPropertyValue("Column" + (j + 1), list[k]) + "," + "Column" + (j + 1);
                                objlist.RemoveAt(objlist.Count - 1);
                                objlist.Add(new KeyValuePair<string, int>(key, k));
                            }
                            else
                                objlist.Add(new KeyValuePair<string, int>(DataConversionTools.GetPropertyValue("Column" + (j + 1), list[k]) + "," + "Column" + (j + 1), k));
                        }

                if (objlist.Count == 0)                                             //if 0 means there is no header column values , add all lines to generate xml into one
                    for (int k = 0; k < list.Count; k++)
                        objlist.Add(new KeyValuePair<string, int>("," + "Column" + (0), k));

                string wholeXML = string.Empty;
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(xml);
                string payload = xmlDoc.GetElementsByTagName(FirstLine)[0].InnerXml; //get payload innerxml
                XmlNode xn = xmlDoc.GetElementsByTagName(FirstLine)[0];              //get xml out of payload 
                xn.RemoveAll();
                string outXML = xmlDoc.InnerXml;
                foreach (KeyValuePair<string, int> item in objlist)                 //generate xml by different header column values
                {
                    tmpStrHeader = item.Key;
                    if (!headerlist.Exists(pred2))
                    {
                        wholeXML += GeneLinesByHeader(payload, outXML, item, list, dtHeader, dtLine, bu, sa, objlist, FirstLine, list[0].SunComponent, list[0].SunMethod);
                        headerlist.Add(new KeyValuePair<string, string>(tmpStrHeader, ""));
                        wholeXML += "\r\n***\r\n";
                    }
                }
                return wholeXML.Replace("[BU]", bu).Replace("[SUSPENSE]", sa).Replace("[SUNID]", SessionInfo.UserInfo.SunUserIP);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + "," + ex.InnerException + "," + ex.StackTrace);
            }
        }
        /// <summary>
        /// //xtw.Indentation = 2;//xtw.IndentChar = ' ';
        /// </summary>
        /// <param name="sUnformattedXml"></param>
        /// <returns></returns>
        public string FormatXml(string sUnformattedXml)
        {
            XmlDocument xd = new XmlDocument();
            xd.LoadXml(sUnformattedXml);
            StringBuilder sb = new StringBuilder();
            StringWriter sw = new StringWriter(sb);
            XmlTextWriter xtw = null;
            try
            {
                xtw = new XmlTextWriter(sw);
                xtw.Formatting = Formatting.Indented;
                xd.WriteTo(xtw);
            }
            finally
            {
                if (xtw != null)
                    xtw.Close();
            }
            return sb.ToString();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public string GetTransUpdXMLScript(List<ExcelAddIn4.Common2.LedgerUpdate> list)
        {
            try
            {
                ExcelAddIn4.Common2.SSC de = new ExcelAddIn4.Common2.SSC();
                de.Payload = new List<Common2.LedgerUpdate>();
                de.ErrorContext = new ExcelAddIn4.Common2.ErrorContext();
                de.ErrorContext.CompatibilityMode = "0";
                de.ErrorContext.ErrorOutput = "1";
                de.ErrorContext.ErrorThreshold = "1";
                de.SunSystemsContext = new ExcelAddIn4.Common2.SunSystemsContext();
                de.SunSystemsContext.BusinessUnit = Finance_Tools.GetAppConfig("BusinessUnit");
                if (list.Count != 0 && !string.IsNullOrEmpty(list[0].Ledger))
                {
                    de.SunSystemsContext.BudgetCode = list[0].Ledger;
                }
                else
                {
                    de.SunSystemsContext.BudgetCode = "";
                }
                de.Payload = list;
                de.ErrorMessages = ".";
                string str = XmlSerialization<ExcelAddIn4.Common2.SSC>.Serialize(de).Replace("<?xml version=\"1.0\"?>", "<?xml version=\"1.0\" encoding=\"UTF-8\"?>").Replace("<SSC xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">", "<SSC>");
                return str;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + "," + ex.InnerException + "," + ex.StackTrace);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="prefix"></param>
        /// <param name="invName"></param>
        /// <param name="i"></param>
        public void GetInvoiceInfo(ref string prefix, ref string invName, ref int? i)
        {
            if (SessionInfo.UserInfo.UseSequenceNumbering == "1")
            {
                try
                {
                    var maxnum = ProcessMaxNumber();
                    if (string.IsNullOrEmpty(SessionInfo.UserInfo.SequencePrefix))
                        prefix = ProcessPrefix();
                    else
                        prefix = SessionInfo.UserInfo.SequencePrefix;

                    string IDvalue = string.Empty;
                    if (SessionInfo.UserInfo.Dictionary.dict.Count != 0 && SessionInfo.UserInfo.Dictionary.dict.ContainsKey(SessionInfo.UserInfo.CachePath))
                    {
                        IDvalue = SessionInfo.UserInfo.Dictionary.dict[SessionInfo.UserInfo.CachePath];
                        string[] sArray = Regex.Split(IDvalue, ",");
                        IDvalue = sArray[0];
                        int InvNumber;
                        int.TryParse(sArray[1], out InvNumber);
                        SessionInfo.UserInfo.InvNumber = InvNumber;
                    }
                    if (!IsGUID(Path.GetFileNameWithoutExtension(SessionInfo.UserInfo.CachePath)) || string.IsNullOrEmpty(IDvalue))
                        i = maxnum.Value + 1;
                    else
                        i = SessionInfo.UserInfo.InvNumber;

                    string number = i.ToString();
                    //if (number.Length < 5)
                    //{
                    //int padNum = 6 - number.Length;
                    number = number.PadLeft(5, '0');
                    //}
                    invName = prefix.Trim() + number;
                }
                catch
                {
                    //prefix = dgv.Rows[i].Cells["TransactionReference"] == null ? "" : dgv.Rows[i].Cells["TransactionReference"].Value.ToString().Substring(0, 3);
                    i = 1;
                    invName = prefix.Trim() + "00001";
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="list"></param>
        /// <param name="ledger"></param>
        /// <returns></returns>
        public string GetXMLScriptQuery(List<Line> list, string ledger)
        {
            try
            {
                ExcelAddIn4.Common.SSC de = new ExcelAddIn4.Common.SSC();
                de.Payload = new ExcelAddIn4.Common.Payload();
                de.ErrorContext = new ExcelAddIn4.Common.ErrorContext();
                de.ErrorContext.CompatibilityMode = "0";
                de.ErrorContext.ErrorOutput = "1";
                de.ErrorContext.ErrorThreshold = "1";
                de.SunSystemsContext = new ExcelAddIn4.Common.SunSystemsContext();
                de.SunSystemsContext.BusinessUnit = Finance_Tools.GetAppConfig("BusinessUnit");
                if (list.Count != 0 && !string.IsNullOrEmpty(ledger))//list[0].Ledger
                {
                    de.SunSystemsContext.BudgetCode = ledger;//list[0].Ledger;
                }
                else
                {
                    de.SunSystemsContext.BudgetCode = "";
                }
                de.Payload.Filter = OutputContainer.filter;
                de.Payload.Select = new Common.Select();
                de.Payload.Select.Ledger = list;
                //de.ErrorMessages = ".";
                string str = XmlSerialization<ExcelAddIn4.Common.SSC>.Serialize(de).Replace("<?xml version=\"1.0\"?>", "<?xml version=\"1.0\" encoding=\"UTF-8\"?>").Replace("<SSC xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">", "<SSC>");
                return str;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + "," + ex.InnerException + "," + ex.StackTrace);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cb"></param>
        /// <param name="dgv"></param>
        public void BindDropdowns(DataGridViewComboBoxColumn cb, DataGridView dgv)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("name");
            dt.Columns.Add("value");
            dt.Rows.Add("", "");
            cb.DisplayMember = "name";
            cb.ValueMember = "value";
            foreach (DataGridViewColumn dc in dgv.Columns)
                dt.Rows.Add(dc.HeaderText, dc.Name);
            cb.DataSource = dt;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ConnectionStringsName"></param>
        /// <param name="elementValue"></param>
        public static void ConnectionStringsSave(string ConnectionStringsName, string elementValue)
        {
            System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.ConnectionStrings.ConnectionStrings[ConnectionStringsName].ConnectionString = elementValue;
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("connectionStrings");
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ConnectionStringsName"></param>
        /// <param name="elementValue"></param>
        public static void AppSettingSave(string AppSettingName, string elementValue)
        {
            System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings[AppSettingName].Value = elementValue;
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static bool IsDateString(string str)
        {
            return Regex.IsMatch(str, @"(\d{4})-(\d{1,2})-(\d{1,2})");
        }
        /// <summary>
        /// //return Regex.IsMatch(str, @"(\d{6})");
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static bool IsPeriodString(string str)
        {
            if (str.Length == 6 || str.Length == 7 || str.Length == 8)
                return true;
            else
                return false;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static bool IsFormulaPeriodString(string str)
        {
            if (str.Length == 7 || str.Length == 8)
            {
                try
                {
                    string[] str2 = str.Split('/');
                    if ((str2.Length == 2) && (str2[0].Length == 4))
                        return true;
                    else
                        return false;
                }
                catch { return false; }
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="fDateTime"></param>
        /// <param name="formatStr"></param>
        /// <returns></returns>
        public static string GetStandardDateTime(string fDateTime, string formatStr)
        {
            if (fDateTime == "0000-0-0 0:00:00")
                return fDateTime;
            DateTime time = new DateTime(1900, 1, 1, 0, 0, 0, 0);
            if (DateTime.TryParse(fDateTime, out time))
                return time.ToString(formatStr);
            else
                return "N/A";
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="fDateTime"></param>
        /// <returns></returns>
        public static string GetStandardDateTime(string fDateTime)
        {
            return GetStandardDateTime(fDateTime, "yyyy-MM-dd HH:mm:ss");
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="fDate"></param>
        /// <returns></returns>
        public static string GetStandardDate(string fDate)
        {
            return GetStandardDateTime(fDate, "yyyy-MM-dd");
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="keyData"></param>
        /// <param name="dgv"></param>
        /// <returns></returns>
        public bool EditingControlWantsInputKey(Keys keyData, DataGridView dgv)
        {
            switch (keyData)
            {
                case Keys.Tab:
                    return true;
                case Keys.Home:
                case Keys.End:
                case Keys.Left:
                case Keys.Right:
                    return true;
                case Keys.Delete:
                    dgv.CurrentCell.Value = "";
                    return true;
                case Keys.Enter:
                    dgv.NotifyCurrentCellDirty(true);
                    return false;
                default:
                    return false;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public DataGridView IniGroups()
        {
            DataGridView d = new DataGridView();
            d.Columns.Add("ID", "ID");
            d.Columns["ID"].DataPropertyName = "ID";
            d.Columns["ID"].Visible = false;
            d.Columns.Add("GroupName", "Group Name");
            d.Columns["GroupName"].DataPropertyName = "GroupName";
            d.Columns["GroupName"].FillWeight = 10;
            DataGridViewCheckBoxColumn column = new DataGridViewCheckBoxColumn();
            column.DataPropertyName = "GroupDisable";
            column.Name = "GroupDisable";
            column.HeaderText = "Disable";
            column.FillWeight = 10;
            d.Columns.Add(column);
            d.Columns.Add("FolderMaskName", "Folder Display Name");
            d.Columns["FolderMaskName"].DataPropertyName = "AddInTabName";
            d.Columns["FolderMaskName"].FillWeight = 30;
            d.Columns["FolderMaskName"].Visible = false;
            d.Columns.Add("Remark", "Remark");
            d.Columns["Remark"].DataPropertyName = "Remark";
            d.Columns["Remark"].FillWeight = 30;
            return d;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public DataGridView IniUsers()
        {
            DataGridView d = new DataGridView();
            d.Columns.Add("UserID", "UserID");
            d.Columns["UserID"].DataPropertyName = "ft_id";
            d.Columns["UserID"].Visible = false;
            DataGridViewCheckBoxColumn colCB = new DataGridViewCheckBoxColumn();
            ExcelAddIn4.Component.DatagridViewCheckBoxHeaderCell cbHeader = new ExcelAddIn4.Component.DatagridViewCheckBoxHeaderCell();
            colCB.HeaderCell = cbHeader;
            colCB.HeaderText = "";
            colCB.FillWeight = 10;
            cbHeader.OnCheckBoxClicked += (a) => { d.Rows.OfType<DataGridViewRow>().ToList().ForEach(t => t.Cells[1].Value = a); };
            d.Columns.Add(colCB);
            d.Columns.Add("WindowsUserID", "Windows User ID");
            d.Columns["WindowsUserID"].DataPropertyName = "WindowsUserID";
            d.Columns["WindowsUserID"].FillWeight = 40;
            d.Columns["WindowsUserID"].ReadOnly = true;
            d.Columns["WindowsUserID"].DefaultCellStyle.BackColor = System.Drawing.Color.LightGray;
            d.Columns.Add("AddInTabName", "AddIn Tab Display Name");
            d.Columns["AddInTabName"].DataPropertyName = "AddInTabName";
            d.Columns["AddInTabName"].FillWeight = 50;
            return d;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public DataGridView IniPermissions()
        {
            DataGridView d = new DataGridView();
            d.Columns.Add("ID", "ID");
            d.Columns["ID"].DataPropertyName = "ID";
            d.Columns["ID"].Visible = false;
            DataGridViewCheckBoxColumn colCB = new DataGridViewCheckBoxColumn();
            ExcelAddIn4.Component.DatagridViewCheckBoxHeaderCell cbHeader = new ExcelAddIn4.Component.DatagridViewCheckBoxHeaderCell();
            colCB.HeaderCell = cbHeader;
            colCB.HeaderText = "";
            colCB.FillWeight = 5;
            cbHeader.OnCheckBoxClicked += (a) => { d.Rows.OfType<DataGridViewRow>().Where(t => t.Visible == true).ToList().ForEach(t => t.Cells[1].Value = a); };
            d.Columns.Add(colCB);
            d.Columns.Add("PermissionName", "Permission Name");
            d.Columns["PermissionName"].DataPropertyName = "PermissionName";
            d.Columns["PermissionName"].FillWeight = 25;
            d.Columns.Add("TemplateName", "Belongs to");
            d.Columns["TemplateName"].DataPropertyName = "TemplateName";
            d.Columns["TemplateName"].FillWeight = 15;
            d.Columns["TemplateName"].ReadOnly = true;
            d.Columns["TemplateName"].DefaultCellStyle.BackColor = System.Drawing.Color.LightGray;
            d.Columns.Add("PerType", "Type");
            d.Columns["PerType"].DataPropertyName = "PerType";
            d.Columns["PerType"].FillWeight = 15;
            d.Columns["PerType"].ReadOnly = true;
            d.Columns["PerType"].DefaultCellStyle.BackColor = System.Drawing.Color.LightGray;
            d.Columns.Add("Remark", "Remark");
            d.Columns["Remark"].DataPropertyName = "Remark";
            d.Columns["Remark"].FillWeight = 20;
            d.Columns.Add("TemplateID", "TemplateID");
            d.Columns["TemplateID"].DataPropertyName = "TemplateID";
            d.Columns["TemplateID"].Visible = false;
            ExcelAddIn4.Component.MyDataGridViewColumn colTB = new ExcelAddIn4.Component.MyDataGridViewColumn();
            colTB.DataGridViewTextChanged += new Component.EventCellChangeEvent(colTB_DataGridViewTextChanged);
            colTB.HeaderText = "Folder";
            colTB.DataPropertyName = "Folder";
            colTB.Name = "Folder";
            colTB.FillWeight = 15;
            d.Columns.Add(colTB);
            return d;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void colTB_DataGridViewTextChanged(object sender, ExcelAddIn4.Component.EventCellChangeArgs e)
        {
            object tempid = ((System.Windows.Forms.DataGridViewTextBoxEditingControl)sender).EditingControlDataGridView.Rows[e.rowIndex].Cells["TemplateID"].FormattedValue;
            ((System.Windows.Forms.DataGridViewTextBoxEditingControl)sender).EditingControlDataGridView.Rows.OfType<DataGridViewRow>().Where(t => t.Cells[6].Value != null).Where(t => t.Cells[6].Value.ToString() == tempid.ToString()).ToList().ForEach(t => t.Cells[7].Value = e.value);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public DataGridView IniXMLFormGrd()
        {
            DataTable dt = GetUserDataFriendlyName();
            DataColumn[] keys = new DataColumn[1];
            keys[0] = dt.Columns["SunField"];
            dt.PrimaryKey = keys;
            DataTable dt2 = GetUserDataGenFriendlyName();
            DataColumn[] keys2 = new DataColumn[1];
            keys2[0] = dt2.Columns["SunField"];
            dt2.PrimaryKey = keys2;
            DataGridView d = new DataGridView();
            if (dt.Rows.Count > 0)
            {
                d.Columns.Add("LineIndicator", "Line Indicator");
                d.Columns["LineIndicator"].DataPropertyName = "LineIndicator";
                d.Columns.Add("Ledger", dt.Rows.Find("Ledger")["UserFriendlyName"].ToString());
                d.Columns["Ledger"].DataPropertyName = "Ledger";
                d.Columns.Add("AccountCode", dt.Rows.Find("AccountCode")["UserFriendlyName"].ToString());
                d.Columns["AccountCode"].DataPropertyName = "AccountCode";
                d.Columns.Add("AccountingPeriod", dt.Rows.Find("AccountingPeriod")["UserFriendlyName"].ToString());
                d.Columns["AccountingPeriod"].DataPropertyName = "AccountingPeriod";
                d.Columns.Add("TransactionDate", dt.Rows.Find("TransactionDate")["UserFriendlyName"].ToString());
                d.Columns["TransactionDate"].DataPropertyName = "TransactionDate";
                d.Columns.Add("DueDate", dt.Rows.Find("DueDate")["UserFriendlyName"].ToString());
                d.Columns["DueDate"].DataPropertyName = "DueDate";
                d.Columns.Add("JournalType", dt.Rows.Find("JournalType")["UserFriendlyName"].ToString());
                d.Columns["JournalType"].DataPropertyName = "JournalType";
                d.Columns.Add("JournalSource", dt.Rows.Find("JournalSource")["UserFriendlyName"].ToString());
                d.Columns["JournalSource"].DataPropertyName = "JournalSource";
                d.Columns.Add("TransactionReference", dt.Rows.Find("TransactionReference")["UserFriendlyName"].ToString());
                d.Columns["TransactionReference"].DataPropertyName = "TransactionReference";
                d.Columns.Add("Description", dt.Rows.Find("Description")["UserFriendlyName"].ToString());
                d.Columns["Description"].DataPropertyName = "Description";
                d.Columns.Add("AllocationMarker", dt.Rows.Find("AllocationMarker")["UserFriendlyName"].ToString());
                d.Columns["AllocationMarker"].DataPropertyName = "AllocationMarker";
                d.Columns.Add("AnalysisCode1", dt.Rows.Find("AnalysisCode1")["UserFriendlyName"].ToString());
                d.Columns["AnalysisCode1"].DataPropertyName = "AnalysisCode1";
                d.Columns.Add("AnalysisCode2", dt.Rows.Find("AnalysisCode2")["UserFriendlyName"].ToString());
                d.Columns["AnalysisCode2"].DataPropertyName = "AnalysisCode2";
                d.Columns.Add("AnalysisCode3", dt.Rows.Find("AnalysisCode3")["UserFriendlyName"].ToString());
                d.Columns["AnalysisCode3"].DataPropertyName = "AnalysisCode3";
                d.Columns.Add("AnalysisCode4", dt.Rows.Find("AnalysisCode4")["UserFriendlyName"].ToString());
                d.Columns["AnalysisCode4"].DataPropertyName = "AnalysisCode4";
                d.Columns.Add("AnalysisCode5", dt.Rows.Find("AnalysisCode5")["UserFriendlyName"].ToString());
                d.Columns["AnalysisCode5"].DataPropertyName = "AnalysisCode5";
                d.Columns.Add("AnalysisCode6", dt.Rows.Find("AnalysisCode6")["UserFriendlyName"].ToString());
                d.Columns["AnalysisCode6"].DataPropertyName = "AnalysisCode6";
                d.Columns.Add("AnalysisCode7", dt.Rows.Find("AnalysisCode7")["UserFriendlyName"].ToString());
                d.Columns["AnalysisCode7"].DataPropertyName = "AnalysisCode7";
                d.Columns.Add("AnalysisCode8", dt.Rows.Find("AnalysisCode8")["UserFriendlyName"].ToString());
                d.Columns["AnalysisCode8"].DataPropertyName = "AnalysisCode8";
                d.Columns.Add("AnalysisCode9", dt.Rows.Find("AnalysisCode9")["UserFriendlyName"].ToString());
                d.Columns["AnalysisCode9"].DataPropertyName = "AnalysisCode9";
                d.Columns.Add("AnalysisCode10", dt.Rows.Find("AnalysisCode10")["UserFriendlyName"].ToString());
                d.Columns["AnalysisCode10"].DataPropertyName = "AnalysisCode10";
                if (dt2.Rows.Count > 0)
                {
                    d.Columns.Add("GeneralDescription1", dt2.Rows.Find("GeneralDescription1")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription1"].DataPropertyName = "GenDesc1";
                    d.Columns.Add("GeneralDescription2", dt2.Rows.Find("GeneralDescription2")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription2"].DataPropertyName = "GenDesc2";
                    d.Columns.Add("GeneralDescription3", dt2.Rows.Find("GeneralDescription3")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription3"].DataPropertyName = "GenDesc3";
                    d.Columns.Add("GeneralDescription4", dt2.Rows.Find("GeneralDescription4")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription4"].DataPropertyName = "GenDesc4";
                    d.Columns.Add("GeneralDescription5", dt2.Rows.Find("GeneralDescription5")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription5"].DataPropertyName = "GenDesc5";
                    d.Columns.Add("GeneralDescription6", dt2.Rows.Find("GeneralDescription6")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription6"].DataPropertyName = "GenDesc6";
                    d.Columns.Add("GeneralDescription7", dt2.Rows.Find("GeneralDescription7")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription7"].DataPropertyName = "GenDesc7";
                    d.Columns.Add("GeneralDescription8", dt2.Rows.Find("GeneralDescription8")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription8"].DataPropertyName = "GenDesc8";
                    d.Columns.Add("GeneralDescription9", dt2.Rows.Find("GeneralDescription9")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription9"].DataPropertyName = "GenDesc9";
                    d.Columns.Add("GeneralDescription10", dt2.Rows.Find("GeneralDescription10")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription10"].DataPropertyName = "GenDesc10";
                    d.Columns.Add("GeneralDescription11", dt2.Rows.Find("GeneralDescription11")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription11"].DataPropertyName = "GenDesc11";
                    d.Columns.Add("GeneralDescription12", dt2.Rows.Find("GeneralDescription12")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription12"].DataPropertyName = "GenDesc12";
                    d.Columns.Add("GeneralDescription13", dt2.Rows.Find("GeneralDescription13")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription13"].DataPropertyName = "GenDesc13";
                    d.Columns.Add("GeneralDescription14", dt2.Rows.Find("GeneralDescription14")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription14"].DataPropertyName = "GenDesc14";
                    d.Columns.Add("GeneralDescription15", dt2.Rows.Find("GeneralDescription15")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription15"].DataPropertyName = "GenDesc15";
                    d.Columns.Add("GeneralDescription16", dt2.Rows.Find("GeneralDescription16")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription16"].DataPropertyName = "GenDesc16";
                    d.Columns.Add("GeneralDescription17", dt2.Rows.Find("GeneralDescription17")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription17"].DataPropertyName = "GenDesc17";
                    d.Columns.Add("GeneralDescription18", dt2.Rows.Find("GeneralDescription18")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription18"].DataPropertyName = "GenDesc18";
                    d.Columns.Add("GeneralDescription19", dt2.Rows.Find("GeneralDescription19")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription19"].DataPropertyName = "GenDesc19";
                    d.Columns.Add("GeneralDescription20", dt2.Rows.Find("GeneralDescription20")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription20"].DataPropertyName = "GenDesc20";
                    d.Columns.Add("GeneralDescription21", dt2.Rows.Find("GeneralDescription21")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription21"].DataPropertyName = "GenDesc21";
                    d.Columns.Add("GeneralDescription22", dt2.Rows.Find("GeneralDescription22")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription22"].DataPropertyName = "GenDesc22";
                    d.Columns.Add("GeneralDescription23", dt2.Rows.Find("GeneralDescription23")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription23"].DataPropertyName = "GenDesc23";
                    d.Columns.Add("GeneralDescription24", dt2.Rows.Find("GeneralDescription24")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription24"].DataPropertyName = "GenDesc24";
                    d.Columns.Add("GeneralDescription25", dt2.Rows.Find("GeneralDescription25")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription25"].DataPropertyName = "GenDesc25";
                }
                else
                {
                    d.Columns.Add("GeneralDescription1", "Gen Desc1");
                    d.Columns["GeneralDescription1"].DataPropertyName = "GenDesc1";
                    d.Columns.Add("GeneralDescription2", "Gen Desc2");
                    d.Columns["GeneralDescription2"].DataPropertyName = "GenDesc2";
                    d.Columns.Add("GeneralDescription3", "Gen Desc3");
                    d.Columns["GeneralDescription3"].DataPropertyName = "GenDesc3";
                    d.Columns.Add("GeneralDescription4", "Gen Desc4");
                    d.Columns["GeneralDescription4"].DataPropertyName = "GenDesc4";
                    d.Columns.Add("GeneralDescription5", "Gen Desc5");
                    d.Columns["GeneralDescription5"].DataPropertyName = "GenDesc5";
                    d.Columns.Add("GeneralDescription6", "Gen Desc6");
                    d.Columns["GeneralDescription6"].DataPropertyName = "GenDesc6";
                    d.Columns.Add("GeneralDescription7", "Gen Desc7");
                    d.Columns["GeneralDescription7"].DataPropertyName = "GenDesc7";
                    d.Columns.Add("GeneralDescription8", "Gen Desc8");
                    d.Columns["GeneralDescription8"].DataPropertyName = "GenDesc8";
                    d.Columns.Add("GeneralDescription9", "Gen Desc9");
                    d.Columns["GeneralDescription9"].DataPropertyName = "GenDesc9";
                    d.Columns.Add("GeneralDescription10", "Gen Desc10");
                    d.Columns["GeneralDescription10"].DataPropertyName = "GenDesc10";
                    d.Columns.Add("GeneralDescription11", "Gen Desc11");
                    d.Columns["GeneralDescription11"].DataPropertyName = "GenDesc11";
                    d.Columns.Add("GeneralDescription12", "Gen Desc12");
                    d.Columns["GeneralDescription12"].DataPropertyName = "GenDesc12";
                    d.Columns.Add("GeneralDescription13", "Gen Desc13");
                    d.Columns["GeneralDescription13"].DataPropertyName = "GenDesc13";
                    d.Columns.Add("GeneralDescription14", "Gen Desc14");
                    d.Columns["GeneralDescription14"].DataPropertyName = "GenDesc14";
                    d.Columns.Add("GeneralDescription15", "Gen Desc15");
                    d.Columns["GeneralDescription15"].DataPropertyName = "GenDesc15";
                    d.Columns.Add("GeneralDescription16", "Gen Desc16");
                    d.Columns["GeneralDescription16"].DataPropertyName = "GenDesc16";
                    d.Columns.Add("GeneralDescription17", "Gen Desc17");
                    d.Columns["GeneralDescription17"].DataPropertyName = "GenDesc17";
                    d.Columns.Add("GeneralDescription18", "Gen Desc18");
                    d.Columns["GeneralDescription18"].DataPropertyName = "GenDesc18";
                    d.Columns.Add("GeneralDescription19", "Gen Desc19");
                    d.Columns["GeneralDescription19"].DataPropertyName = "GenDesc19";
                    d.Columns.Add("GeneralDescription20", "Gen Desc20");
                    d.Columns["GeneralDescription20"].DataPropertyName = "GenDesc20";
                    d.Columns.Add("GeneralDescription21", "Gen Desc21");
                    d.Columns["GeneralDescription21"].DataPropertyName = "GenDesc21";
                    d.Columns.Add("GeneralDescription22", "Gen Desc22");
                    d.Columns["GeneralDescription22"].DataPropertyName = "GenDesc22";
                    d.Columns.Add("GeneralDescription23", "Gen Desc23");
                    d.Columns["GeneralDescription23"].DataPropertyName = "GenDesc23";
                    d.Columns.Add("GeneralDescription24", "Gen Desc24");
                    d.Columns["GeneralDescription24"].DataPropertyName = "GenDesc24";
                    d.Columns.Add("GeneralDescription25", "Gen Desc25");
                    d.Columns["GeneralDescription25"].DataPropertyName = "GenDesc25";
                }
                d.Columns.Add("TransactionAmount", dt.Rows.Find("TransactionAmount")["UserFriendlyName"].ToString());
                d.Columns["TransactionAmount"].DataPropertyName = "TransactionAmount";
                d.Columns.Add("CurrencyCode", dt.Rows.Find("CurrencyCode")["UserFriendlyName"].ToString());
                d.Columns["CurrencyCode"].DataPropertyName = "CurrencyCode";
                d.Columns.Add("BaseAmount", dt.Rows.Find("BaseAmount")["UserFriendlyName"].ToString());
                d.Columns["BaseAmount"].DataPropertyName = "BaseAmount";
                d.Columns.Add("Base2ReportingAmount", dt.Rows.Find("Base2ReportingAmount")["UserFriendlyName"].ToString());
                d.Columns["Base2ReportingAmount"].DataPropertyName = "Base2ReportingAmount";
                d.Columns.Add("Value4Amount", dt.Rows.Find("Value4Amount")["UserFriendlyName"].ToString());
                d.Columns["Value4Amount"].DataPropertyName = "Value4Amount";
                d.Columns.Add("Reference", "Ref");
                d.Columns["Reference"].DataPropertyName = "Reference";
                d.Columns["Reference"].Visible = false;
                d.Columns.Add("SaveReference", "SaveReference");
                d.Columns["SaveReference"].DataPropertyName = "SaveReference";
                d.Columns["SaveReference"].Visible = false;
                d.Columns.Add("populatecellwithJN", "populatecellwithJN");
                d.Columns["populatecellwithJN"].DataPropertyName = "populatecellwithJN";
                d.Columns["populatecellwithJN"].Visible = false;
                d.Columns.Add("BalanceBy", "BalanceBy");
                d.Columns["BalanceBy"].DataPropertyName = "BalanceBy";
                d.Columns["BalanceBy"].Visible = false;
                d.Columns.Add("StartInCell", "StartInCell");
                d.Columns["StartInCell"].DataPropertyName = "StartInCell";
                d.Columns["StartInCell"].Visible = false;
                d.Columns.Add("AllowBalTrans", "AllowBalTrans");
                d.Columns["AllowBalTrans"].DataPropertyName = "AllowBalTrans";
                d.Columns["AllowBalTrans"].Visible = false;
                d.Columns.Add("AllowPostSuspAcco", "AllowPostSuspAcco");
                d.Columns["AllowPostSuspAcco"].DataPropertyName = "AllowPostSuspAcco";
                d.Columns["AllowPostSuspAcco"].Visible = false;
            }
            else
            {
                d.Columns.Add("LineIndicator", "Line Indicator");
                d.Columns["LineIndicator"].DataPropertyName = "LineIndicator";
                d.Columns.Add("Ledger", "Ledger");
                d.Columns["Ledger"].DataPropertyName = "Ledger";
                d.Columns.Add("AccountCode", "Account");
                d.Columns["AccountCode"].DataPropertyName = "AccountCode";
                d.Columns.Add("AccountingPeriod", "Period");
                d.Columns["AccountingPeriod"].DataPropertyName = "AccountingPeriod";
                d.Columns.Add("TransactionDate", "Trans Date");
                d.Columns["TransactionDate"].DataPropertyName = "TransactionDate";
                d.Columns.Add("DueDate", "Due Date");
                d.Columns["DueDate"].DataPropertyName = "DueDate";
                d.Columns.Add("JournalType", "Jrnl Type");
                d.Columns["JournalType"].DataPropertyName = "JournalType";
                d.Columns.Add("JournalSource", "Jrnl Source");
                d.Columns["JournalSource"].DataPropertyName = "JournalSource";
                d.Columns.Add("TransactionReference", "Trans Ref");
                d.Columns["TransactionReference"].DataPropertyName = "TransactionReference";
                d.Columns.Add("Description", "Description");
                d.Columns["Description"].DataPropertyName = "Description";
                d.Columns.Add("AllocationMarker", "Alloctn Marker");
                d.Columns["AllocationMarker"].DataPropertyName = "AllocationMarker";
                d.Columns.Add("AnalysisCode1", "LA1");
                d.Columns["AnalysisCode1"].DataPropertyName = "AnalysisCode1";
                d.Columns.Add("AnalysisCode2", "LA2");
                d.Columns["AnalysisCode2"].DataPropertyName = "AnalysisCode2";
                d.Columns.Add("AnalysisCode3", "LA3");
                d.Columns["AnalysisCode3"].DataPropertyName = "AnalysisCode3";
                d.Columns.Add("AnalysisCode4", "LA4");
                d.Columns["AnalysisCode4"].DataPropertyName = "AnalysisCode4";
                d.Columns.Add("AnalysisCode5", "LA5");
                d.Columns["AnalysisCode5"].DataPropertyName = "AnalysisCode5";
                d.Columns.Add("AnalysisCode6", "LA6");
                d.Columns["AnalysisCode6"].DataPropertyName = "AnalysisCode6";
                d.Columns.Add("AnalysisCode7", "LA7");
                d.Columns["AnalysisCode7"].DataPropertyName = "AnalysisCode7";
                d.Columns.Add("AnalysisCode8", "LA8");
                d.Columns["AnalysisCode8"].DataPropertyName = "AnalysisCode8";
                d.Columns.Add("AnalysisCode9", "LA9");
                d.Columns["AnalysisCode9"].DataPropertyName = "AnalysisCode9";
                d.Columns.Add("AnalysisCode10", "LA10");
                d.Columns["AnalysisCode10"].DataPropertyName = "AnalysisCode10";
                if (dt2.Rows.Count > 0)
                {
                    d.Columns.Add("GeneralDescription1", dt2.Rows.Find("GeneralDescription1")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription1"].DataPropertyName = "GenDesc1";
                    d.Columns.Add("GeneralDescription2", dt2.Rows.Find("GeneralDescription2")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription2"].DataPropertyName = "GenDesc2";
                    d.Columns.Add("GeneralDescription3", dt2.Rows.Find("GeneralDescription3")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription3"].DataPropertyName = "GenDesc3";
                    d.Columns.Add("GeneralDescription4", dt2.Rows.Find("GeneralDescription4")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription4"].DataPropertyName = "GenDesc4";
                    d.Columns.Add("GeneralDescription5", dt2.Rows.Find("GeneralDescription5")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription5"].DataPropertyName = "GenDesc5";
                    d.Columns.Add("GeneralDescription6", dt2.Rows.Find("GeneralDescription6")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription6"].DataPropertyName = "GenDesc6";
                    d.Columns.Add("GeneralDescription7", dt2.Rows.Find("GeneralDescription7")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription7"].DataPropertyName = "GenDesc7";
                    d.Columns.Add("GeneralDescription8", dt2.Rows.Find("GeneralDescription8")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription8"].DataPropertyName = "GenDesc8";
                    d.Columns.Add("GeneralDescription9", dt2.Rows.Find("GeneralDescription9")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription9"].DataPropertyName = "GenDesc9";
                    d.Columns.Add("GeneralDescription10", dt2.Rows.Find("GeneralDescription10")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription10"].DataPropertyName = "GenDesc10";
                    d.Columns.Add("GeneralDescription11", dt2.Rows.Find("GeneralDescription11")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription11"].DataPropertyName = "GenDesc11";
                    d.Columns.Add("GeneralDescription12", dt2.Rows.Find("GeneralDescription12")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription12"].DataPropertyName = "GenDesc12";
                    d.Columns.Add("GeneralDescription13", dt2.Rows.Find("GeneralDescription13")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription13"].DataPropertyName = "GenDesc13";
                    d.Columns.Add("GeneralDescription14", dt2.Rows.Find("GeneralDescription14")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription14"].DataPropertyName = "GenDesc14";
                    d.Columns.Add("GeneralDescription15", dt2.Rows.Find("GeneralDescription15")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription15"].DataPropertyName = "GenDesc15";
                    d.Columns.Add("GeneralDescription16", dt2.Rows.Find("GeneralDescription16")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription16"].DataPropertyName = "GenDesc16";
                    d.Columns.Add("GeneralDescription17", dt2.Rows.Find("GeneralDescription17")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription17"].DataPropertyName = "GenDesc17";
                    d.Columns.Add("GeneralDescription18", dt2.Rows.Find("GeneralDescription18")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription18"].DataPropertyName = "GenDesc18";
                    d.Columns.Add("GeneralDescription19", dt2.Rows.Find("GeneralDescription19")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription19"].DataPropertyName = "GenDesc19";
                    d.Columns.Add("GeneralDescription20", dt2.Rows.Find("GeneralDescription20")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription20"].DataPropertyName = "GenDesc20";
                    d.Columns.Add("GeneralDescription21", dt2.Rows.Find("GeneralDescription21")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription21"].DataPropertyName = "GenDesc21";
                    d.Columns.Add("GeneralDescription22", dt2.Rows.Find("GeneralDescription22")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription22"].DataPropertyName = "GenDesc22";
                    d.Columns.Add("GeneralDescription23", dt2.Rows.Find("GeneralDescription23")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription23"].DataPropertyName = "GenDesc23";
                    d.Columns.Add("GeneralDescription24", dt2.Rows.Find("GeneralDescription24")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription24"].DataPropertyName = "GenDesc24";
                    d.Columns.Add("GeneralDescription25", dt2.Rows.Find("GeneralDescription25")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription25"].DataPropertyName = "GenDesc25";
                }
                else
                {
                    d.Columns.Add("GeneralDescription1", "Gen Desc1");
                    d.Columns["GeneralDescription1"].DataPropertyName = "GenDesc1";
                    d.Columns.Add("GeneralDescription2", "Gen Desc2");
                    d.Columns["GeneralDescription2"].DataPropertyName = "GenDesc2";
                    d.Columns.Add("GeneralDescription3", "Gen Desc3");
                    d.Columns["GeneralDescription3"].DataPropertyName = "GenDesc3";
                    d.Columns.Add("GeneralDescription4", "Gen Desc4");
                    d.Columns["GeneralDescription4"].DataPropertyName = "GenDesc4";
                    d.Columns.Add("GeneralDescription5", "Gen Desc5");
                    d.Columns["GeneralDescription5"].DataPropertyName = "GenDesc5";
                    d.Columns.Add("GeneralDescription6", "Gen Desc6");
                    d.Columns["GeneralDescription6"].DataPropertyName = "GenDesc6";
                    d.Columns.Add("GeneralDescription7", "Gen Desc7");
                    d.Columns["GeneralDescription7"].DataPropertyName = "GenDesc7";
                    d.Columns.Add("GeneralDescription8", "Gen Desc8");
                    d.Columns["GeneralDescription8"].DataPropertyName = "GenDesc8";
                    d.Columns.Add("GeneralDescription9", "Gen Desc9");
                    d.Columns["GeneralDescription9"].DataPropertyName = "GenDesc9";
                    d.Columns.Add("GeneralDescription10", "Gen Desc10");
                    d.Columns["GeneralDescription10"].DataPropertyName = "GenDesc10";
                    d.Columns.Add("GeneralDescription11", "Gen Desc11");
                    d.Columns["GeneralDescription11"].DataPropertyName = "GenDesc11";
                    d.Columns.Add("GeneralDescription12", "Gen Desc12");
                    d.Columns["GeneralDescription12"].DataPropertyName = "GenDesc12";
                    d.Columns.Add("GeneralDescription13", "Gen Desc13");
                    d.Columns["GeneralDescription13"].DataPropertyName = "GenDesc13";
                    d.Columns.Add("GeneralDescription14", "Gen Desc14");
                    d.Columns["GeneralDescription14"].DataPropertyName = "GenDesc14";
                    d.Columns.Add("GeneralDescription15", "Gen Desc15");
                    d.Columns["GeneralDescription15"].DataPropertyName = "GenDesc15";
                    d.Columns.Add("GeneralDescription16", "Gen Desc16");
                    d.Columns["GeneralDescription16"].DataPropertyName = "GenDesc16";
                    d.Columns.Add("GeneralDescription17", "Gen Desc17");
                    d.Columns["GeneralDescription17"].DataPropertyName = "GenDesc17";
                    d.Columns.Add("GeneralDescription18", "Gen Desc18");
                    d.Columns["GeneralDescription18"].DataPropertyName = "GenDesc18";
                    d.Columns.Add("GeneralDescription19", "Gen Desc19");
                    d.Columns["GeneralDescription19"].DataPropertyName = "GenDesc19";
                    d.Columns.Add("GeneralDescription20", "Gen Desc20");
                    d.Columns["GeneralDescription20"].DataPropertyName = "GenDesc20";
                    d.Columns.Add("GeneralDescription21", "Gen Desc21");
                    d.Columns["GeneralDescription21"].DataPropertyName = "GenDesc21";
                    d.Columns.Add("GeneralDescription22", "Gen Desc22");
                    d.Columns["GeneralDescription22"].DataPropertyName = "GenDesc22";
                    d.Columns.Add("GeneralDescription23", "Gen Desc23");
                    d.Columns["GeneralDescription23"].DataPropertyName = "GenDesc23";
                    d.Columns.Add("GeneralDescription24", "Gen Desc24");
                    d.Columns["GeneralDescription24"].DataPropertyName = "GenDesc24";
                    d.Columns.Add("GeneralDescription25", "Gen Desc25");
                    d.Columns["GeneralDescription25"].DataPropertyName = "GenDesc25";
                }
                d.Columns.Add("TransactionAmount", "Trans Amount");
                d.Columns["TransactionAmount"].DataPropertyName = "TransactionAmount";
                d.Columns.Add("CurrencyCode", "Currency");
                d.Columns["CurrencyCode"].DataPropertyName = "CurrencyCode";
                d.Columns.Add("BaseAmount", "Base Amount");
                d.Columns["BaseAmount"].DataPropertyName = "BaseAmount";
                d.Columns.Add("Base2ReportingAmount", "2nd Base");
                d.Columns["Base2ReportingAmount"].DataPropertyName = "Base2ReportingAmount";
                d.Columns.Add("Value4Amount", "4th Amount");
                d.Columns["Value4Amount"].DataPropertyName = "Value4Amount";
                d.Columns.Add("Reference", "Ref");
                d.Columns["Reference"].DataPropertyName = "Reference";
                d.Columns["Reference"].Visible = false;
                d.Columns.Add("SaveReference", "SaveReference");
                d.Columns["SaveReference"].DataPropertyName = "SaveReference";
                d.Columns["SaveReference"].Visible = false;
                d.Columns.Add("populatecellwithJN", "populatecellwithJN");
                d.Columns["populatecellwithJN"].DataPropertyName = "populatecellwithJN";
                d.Columns["populatecellwithJN"].Visible = false;
                d.Columns.Add("BalanceBy", "BalanceBy");
                d.Columns["BalanceBy"].DataPropertyName = "BalanceBy";
                d.Columns["BalanceBy"].Visible = false;
                d.Columns.Add("StartInCell", "StartInCell");
                d.Columns["StartInCell"].DataPropertyName = "StartInCell";
                d.Columns["StartInCell"].Visible = false;
                d.Columns.Add("AllowBalTrans", "AllowBalTrans");
                d.Columns["AllowBalTrans"].DataPropertyName = "AllowBalTrans";
                d.Columns["AllowBalTrans"].Visible = false;
                d.Columns.Add("AllowPostSuspAcco", "AllowPostSuspAcco");
                d.Columns["AllowPostSuspAcco"].DataPropertyName = "AllowPostSuspAcco";
                d.Columns["AllowPostSuspAcco"].Visible = false;
            }
            return d;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public DataGridView IniXMLFormGrdForTransUpd()
        {
            DataTable dt = GetUserDataFriendlyName();
            DataColumn[] keys = new DataColumn[1];
            keys[0] = dt.Columns["SunField"];
            dt.PrimaryKey = keys;
            DataTable dt2 = GetUserDataGenFriendlyName();
            DataColumn[] keys2 = new DataColumn[1];
            keys2[0] = dt2.Columns["SunField"];
            dt2.PrimaryKey = keys2;
            DataGridView d = new DataGridView();
            if (dt.Rows.Count > 0)
            {
                d.Columns.Add("LineIndicator", "Line Indicator");
                d.Columns["LineIndicator"].DataPropertyName = "LineIndicator";
                d.Columns.Add("JournalNumber", "JournalNumber");
                d.Columns["JournalNumber"].DataPropertyName = "JournalNumber";
                d.Columns.Add("JournalLineNumber", "JournalLineNumber");
                d.Columns["JournalLineNumber"].DataPropertyName = "JournalLineNumber";
                d.Columns.Add("Ledger", dt.Rows.Find("Ledger")["UserFriendlyName"].ToString());
                d.Columns["Ledger"].DataPropertyName = "Ledger";
                d.Columns.Add("AccountCode", dt.Rows.Find("AccountCode")["UserFriendlyName"].ToString());
                d.Columns["AccountCode"].DataPropertyName = "AccountCode";
                d.Columns.Add("AccountingPeriod", dt.Rows.Find("AccountingPeriod")["UserFriendlyName"].ToString());
                d.Columns["AccountingPeriod"].DataPropertyName = "AccountingPeriod";
                d.Columns.Add("TransactionDate", dt.Rows.Find("TransactionDate")["UserFriendlyName"].ToString());
                d.Columns["TransactionDate"].DataPropertyName = "TransactionDate";
                d.Columns.Add("DueDate", dt.Rows.Find("DueDate")["UserFriendlyName"].ToString());
                d.Columns["DueDate"].DataPropertyName = "DueDate";
                d.Columns.Add("JournalType", dt.Rows.Find("JournalType")["UserFriendlyName"].ToString());
                d.Columns["JournalType"].DataPropertyName = "JournalType";
                d.Columns.Add("JournalSource", dt.Rows.Find("JournalSource")["UserFriendlyName"].ToString());
                d.Columns["JournalSource"].DataPropertyName = "JournalSource";
                d.Columns.Add("TransactionReference", dt.Rows.Find("TransactionReference")["UserFriendlyName"].ToString());
                d.Columns["TransactionReference"].DataPropertyName = "TransactionReference";
                d.Columns.Add("Description", dt.Rows.Find("Description")["UserFriendlyName"].ToString());
                d.Columns["Description"].DataPropertyName = "Description";
                d.Columns.Add("AllocationMarker", dt.Rows.Find("AllocationMarker")["UserFriendlyName"].ToString());
                d.Columns["AllocationMarker"].DataPropertyName = "AllocationMarker";
                d.Columns.Add("AnalysisCode1", dt.Rows.Find("AnalysisCode1")["UserFriendlyName"].ToString());
                d.Columns["AnalysisCode1"].DataPropertyName = "AnalysisCode1";
                d.Columns.Add("AnalysisCode2", dt.Rows.Find("AnalysisCode2")["UserFriendlyName"].ToString());
                d.Columns["AnalysisCode2"].DataPropertyName = "AnalysisCode2";
                d.Columns.Add("AnalysisCode3", dt.Rows.Find("AnalysisCode3")["UserFriendlyName"].ToString());
                d.Columns["AnalysisCode3"].DataPropertyName = "AnalysisCode3";
                d.Columns.Add("AnalysisCode4", dt.Rows.Find("AnalysisCode4")["UserFriendlyName"].ToString());
                d.Columns["AnalysisCode4"].DataPropertyName = "AnalysisCode4";
                d.Columns.Add("AnalysisCode5", dt.Rows.Find("AnalysisCode5")["UserFriendlyName"].ToString());
                d.Columns["AnalysisCode5"].DataPropertyName = "AnalysisCode5";
                d.Columns.Add("AnalysisCode6", dt.Rows.Find("AnalysisCode6")["UserFriendlyName"].ToString());
                d.Columns["AnalysisCode6"].DataPropertyName = "AnalysisCode6";
                d.Columns.Add("AnalysisCode7", dt.Rows.Find("AnalysisCode7")["UserFriendlyName"].ToString());
                d.Columns["AnalysisCode7"].DataPropertyName = "AnalysisCode7";
                d.Columns.Add("AnalysisCode8", dt.Rows.Find("AnalysisCode8")["UserFriendlyName"].ToString());
                d.Columns["AnalysisCode8"].DataPropertyName = "AnalysisCode8";
                d.Columns.Add("AnalysisCode9", dt.Rows.Find("AnalysisCode9")["UserFriendlyName"].ToString());
                d.Columns["AnalysisCode9"].DataPropertyName = "AnalysisCode9";
                d.Columns.Add("AnalysisCode10", dt.Rows.Find("AnalysisCode10")["UserFriendlyName"].ToString());
                d.Columns["AnalysisCode10"].DataPropertyName = "AnalysisCode10";
                if (dt2.Rows.Count > 0)
                {
                    d.Columns.Add("GeneralDescription1", dt2.Rows.Find("GeneralDescription1")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription1"].DataPropertyName = "GenDesc1";
                    d.Columns.Add("GeneralDescription2", dt2.Rows.Find("GeneralDescription2")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription2"].DataPropertyName = "GenDesc2";
                    d.Columns.Add("GeneralDescription3", dt2.Rows.Find("GeneralDescription3")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription3"].DataPropertyName = "GenDesc3";
                    d.Columns.Add("GeneralDescription4", dt2.Rows.Find("GeneralDescription4")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription4"].DataPropertyName = "GenDesc4";
                    d.Columns.Add("GeneralDescription5", dt2.Rows.Find("GeneralDescription5")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription5"].DataPropertyName = "GenDesc5";
                    d.Columns.Add("GeneralDescription6", dt2.Rows.Find("GeneralDescription6")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription6"].DataPropertyName = "GenDesc6";
                    d.Columns.Add("GeneralDescription7", dt2.Rows.Find("GeneralDescription7")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription7"].DataPropertyName = "GenDesc7";
                    d.Columns.Add("GeneralDescription8", dt2.Rows.Find("GeneralDescription8")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription8"].DataPropertyName = "GenDesc8";
                    d.Columns.Add("GeneralDescription9", dt2.Rows.Find("GeneralDescription9")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription9"].DataPropertyName = "GenDesc9";
                    d.Columns.Add("GeneralDescription10", dt2.Rows.Find("GeneralDescription10")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription10"].DataPropertyName = "GenDesc10";
                    d.Columns.Add("GeneralDescription11", dt2.Rows.Find("GeneralDescription11")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription11"].DataPropertyName = "GenDesc11";
                    d.Columns.Add("GeneralDescription12", dt2.Rows.Find("GeneralDescription12")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription12"].DataPropertyName = "GenDesc12";
                    d.Columns.Add("GeneralDescription13", dt2.Rows.Find("GeneralDescription13")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription13"].DataPropertyName = "GenDesc13";
                    d.Columns.Add("GeneralDescription14", dt2.Rows.Find("GeneralDescription14")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription14"].DataPropertyName = "GenDesc14";
                    d.Columns.Add("GeneralDescription15", dt2.Rows.Find("GeneralDescription15")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription15"].DataPropertyName = "GenDesc15";
                    d.Columns.Add("GeneralDescription16", dt2.Rows.Find("GeneralDescription16")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription16"].DataPropertyName = "GenDesc16";
                    d.Columns.Add("GeneralDescription17", dt2.Rows.Find("GeneralDescription17")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription17"].DataPropertyName = "GenDesc17";
                    d.Columns.Add("GeneralDescription18", dt2.Rows.Find("GeneralDescription18")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription18"].DataPropertyName = "GenDesc18";
                    d.Columns.Add("GeneralDescription19", dt2.Rows.Find("GeneralDescription19")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription19"].DataPropertyName = "GenDesc19";
                    d.Columns.Add("GeneralDescription20", dt2.Rows.Find("GeneralDescription20")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription20"].DataPropertyName = "GenDesc20";
                    d.Columns.Add("GeneralDescription21", dt2.Rows.Find("GeneralDescription21")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription21"].DataPropertyName = "GenDesc21";
                    d.Columns.Add("GeneralDescription22", dt2.Rows.Find("GeneralDescription22")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription22"].DataPropertyName = "GenDesc22";
                    d.Columns.Add("GeneralDescription23", dt2.Rows.Find("GeneralDescription23")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription23"].DataPropertyName = "GenDesc23";
                    d.Columns.Add("GeneralDescription24", dt2.Rows.Find("GeneralDescription24")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription24"].DataPropertyName = "GenDesc24";
                    d.Columns.Add("GeneralDescription25", dt2.Rows.Find("GeneralDescription25")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription25"].DataPropertyName = "GenDesc25";
                }
                else
                {
                    d.Columns.Add("GeneralDescription1", "Gen Desc1");
                    d.Columns["GeneralDescription1"].DataPropertyName = "GenDesc1";
                    d.Columns.Add("GeneralDescription2", "Gen Desc2");
                    d.Columns["GeneralDescription2"].DataPropertyName = "GenDesc2";
                    d.Columns.Add("GeneralDescription3", "Gen Desc3");
                    d.Columns["GeneralDescription3"].DataPropertyName = "GenDesc3";
                    d.Columns.Add("GeneralDescription4", "Gen Desc4");
                    d.Columns["GeneralDescription4"].DataPropertyName = "GenDesc4";
                    d.Columns.Add("GeneralDescription5", "Gen Desc5");
                    d.Columns["GeneralDescription5"].DataPropertyName = "GenDesc5";
                    d.Columns.Add("GeneralDescription6", "Gen Desc6");
                    d.Columns["GeneralDescription6"].DataPropertyName = "GenDesc6";
                    d.Columns.Add("GeneralDescription7", "Gen Desc7");
                    d.Columns["GeneralDescription7"].DataPropertyName = "GenDesc7";
                    d.Columns.Add("GeneralDescription8", "Gen Desc8");
                    d.Columns["GeneralDescription8"].DataPropertyName = "GenDesc8";
                    d.Columns.Add("GeneralDescription9", "Gen Desc9");
                    d.Columns["GeneralDescription9"].DataPropertyName = "GenDesc9";
                    d.Columns.Add("GeneralDescription10", "Gen Desc10");
                    d.Columns["GeneralDescription10"].DataPropertyName = "GenDesc10";
                    d.Columns.Add("GeneralDescription11", "Gen Desc11");
                    d.Columns["GeneralDescription11"].DataPropertyName = "GenDesc11";
                    d.Columns.Add("GeneralDescription12", "Gen Desc12");
                    d.Columns["GeneralDescription12"].DataPropertyName = "GenDesc12";
                    d.Columns.Add("GeneralDescription13", "Gen Desc13");
                    d.Columns["GeneralDescription13"].DataPropertyName = "GenDesc13";
                    d.Columns.Add("GeneralDescription14", "Gen Desc14");
                    d.Columns["GeneralDescription14"].DataPropertyName = "GenDesc14";
                    d.Columns.Add("GeneralDescription15", "Gen Desc15");
                    d.Columns["GeneralDescription15"].DataPropertyName = "GenDesc15";
                    d.Columns.Add("GeneralDescription16", "Gen Desc16");
                    d.Columns["GeneralDescription16"].DataPropertyName = "GenDesc16";
                    d.Columns.Add("GeneralDescription17", "Gen Desc17");
                    d.Columns["GeneralDescription17"].DataPropertyName = "GenDesc17";
                    d.Columns.Add("GeneralDescription18", "Gen Desc18");
                    d.Columns["GeneralDescription18"].DataPropertyName = "GenDesc18";
                    d.Columns.Add("GeneralDescription19", "Gen Desc19");
                    d.Columns["GeneralDescription19"].DataPropertyName = "GenDesc19";
                    d.Columns.Add("GeneralDescription20", "Gen Desc20");
                    d.Columns["GeneralDescription20"].DataPropertyName = "GenDesc20";
                    d.Columns.Add("GeneralDescription21", "Gen Desc21");
                    d.Columns["GeneralDescription21"].DataPropertyName = "GenDesc21";
                    d.Columns.Add("GeneralDescription22", "Gen Desc22");
                    d.Columns["GeneralDescription22"].DataPropertyName = "GenDesc22";
                    d.Columns.Add("GeneralDescription23", "Gen Desc23");
                    d.Columns["GeneralDescription23"].DataPropertyName = "GenDesc23";
                    d.Columns.Add("GeneralDescription24", "Gen Desc24");
                    d.Columns["GeneralDescription24"].DataPropertyName = "GenDesc24";
                    d.Columns.Add("GeneralDescription25", "Gen Desc25");
                    d.Columns["GeneralDescription25"].DataPropertyName = "GenDesc25";
                }
                d.Columns.Add("TransactionAmount", dt.Rows.Find("TransactionAmount")["UserFriendlyName"].ToString());
                d.Columns["TransactionAmount"].DataPropertyName = "TransactionAmount";
                d.Columns.Add("CurrencyCode", dt.Rows.Find("CurrencyCode")["UserFriendlyName"].ToString());
                d.Columns["CurrencyCode"].DataPropertyName = "CurrencyCode";
                d.Columns.Add("BaseAmount", dt.Rows.Find("BaseAmount")["UserFriendlyName"].ToString());
                d.Columns["BaseAmount"].DataPropertyName = "BaseAmount";
                d.Columns.Add("Base2ReportingAmount", dt.Rows.Find("Base2ReportingAmount")["UserFriendlyName"].ToString());
                d.Columns["Base2ReportingAmount"].DataPropertyName = "Base2ReportingAmount";
                d.Columns.Add("Value4Amount", dt.Rows.Find("Value4Amount")["UserFriendlyName"].ToString());
                d.Columns["Value4Amount"].DataPropertyName = "Value4Amount";
                d.Columns.Add("Actions", "Actions");
                d.Columns["Actions"].DataPropertyName = "Actions";
                d.Columns.Add("Messages", "Messages");
                d.Columns["Messages"].DataPropertyName = "Messages";
                d.Columns.Add("Reference", "Ref");
                d.Columns["Reference"].DataPropertyName = "Reference";
                d.Columns["Reference"].Visible = false;
                d.Columns.Add("SaveReference", "SaveReference");
                d.Columns["SaveReference"].DataPropertyName = "SaveReference";
                d.Columns["SaveReference"].Visible = false;
                d.Columns.Add("populatecellwithJN", "populatecellwithJN");
                d.Columns["populatecellwithJN"].DataPropertyName = "populatecellwithJN";
                d.Columns["populatecellwithJN"].Visible = false;
                d.Columns.Add("BalanceBy", "BalanceBy");
                d.Columns["BalanceBy"].DataPropertyName = "BalanceBy";
                d.Columns["BalanceBy"].Visible = false;
                d.Columns.Add("AccountRange", "AccountRange");
                d.Columns["AccountRange"].DataPropertyName = "AccountRange";
                d.Columns["AccountRange"].Visible = false;
                d.Columns.Add("StartInCell", "StartInCell");
                d.Columns["StartInCell"].DataPropertyName = "StartInCell";
                d.Columns["StartInCell"].Visible = false;
            }
            else
            {
                d.Columns.Add("LineIndicator", "Line Indicator");
                d.Columns["LineIndicator"].DataPropertyName = "LineIndicator";
                d.Columns.Add("JournalNumber", "JournalNumber");
                d.Columns["JournalNumber"].DataPropertyName = "JournalNumber";
                d.Columns.Add("JournalLineNumber", "JournalLineNumber");
                d.Columns["JournalLineNumber"].DataPropertyName = "JournalLineNumber";
                d.Columns.Add("Ledger", "Ledger");
                d.Columns["Ledger"].DataPropertyName = "Ledger";
                d.Columns.Add("AccountCode", "Account");
                d.Columns["AccountCode"].DataPropertyName = "AccountCode";
                d.Columns.Add("AccountingPeriod", "Period");
                d.Columns["AccountingPeriod"].DataPropertyName = "AccountingPeriod";
                d.Columns.Add("TransactionDate", "Trans Date");
                d.Columns["TransactionDate"].DataPropertyName = "TransactionDate";
                d.Columns.Add("DueDate", "Due Date");
                d.Columns["DueDate"].DataPropertyName = "DueDate";
                d.Columns.Add("JournalType", "Jrnl Type");
                d.Columns["JournalType"].DataPropertyName = "JournalType";
                d.Columns.Add("JournalSource", "Jrnl Source");
                d.Columns["JournalSource"].DataPropertyName = "JournalSource";
                d.Columns.Add("TransactionReference", "Trans Ref");
                d.Columns["TransactionReference"].DataPropertyName = "TransactionReference";
                d.Columns.Add("Description", "Description");
                d.Columns["Description"].DataPropertyName = "Description";
                d.Columns.Add("AllocationMarker", "Alloctn Marker");
                d.Columns["AllocationMarker"].DataPropertyName = "AllocationMarker";
                d.Columns.Add("AnalysisCode1", "LA1");
                d.Columns["AnalysisCode1"].DataPropertyName = "AnalysisCode1";
                d.Columns.Add("AnalysisCode2", "LA2");
                d.Columns["AnalysisCode2"].DataPropertyName = "AnalysisCode2";
                d.Columns.Add("AnalysisCode3", "LA3");
                d.Columns["AnalysisCode3"].DataPropertyName = "AnalysisCode3";
                d.Columns.Add("AnalysisCode4", "LA4");
                d.Columns["AnalysisCode4"].DataPropertyName = "AnalysisCode4";
                d.Columns.Add("AnalysisCode5", "LA5");
                d.Columns["AnalysisCode5"].DataPropertyName = "AnalysisCode5";
                d.Columns.Add("AnalysisCode6", "LA6");
                d.Columns["AnalysisCode6"].DataPropertyName = "AnalysisCode6";
                d.Columns.Add("AnalysisCode7", "LA7");
                d.Columns["AnalysisCode7"].DataPropertyName = "AnalysisCode7";
                d.Columns.Add("AnalysisCode8", "LA8");
                d.Columns["AnalysisCode8"].DataPropertyName = "AnalysisCode8";
                d.Columns.Add("AnalysisCode9", "LA9");
                d.Columns["AnalysisCode9"].DataPropertyName = "AnalysisCode9";
                d.Columns.Add("AnalysisCode10", "LA10");
                d.Columns["AnalysisCode10"].DataPropertyName = "AnalysisCode10";
                if (dt2.Rows.Count > 0)
                {
                    d.Columns.Add("GeneralDescription1", dt2.Rows.Find("GeneralDescription1")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription1"].DataPropertyName = "GenDesc1";
                    d.Columns.Add("GeneralDescription2", dt2.Rows.Find("GeneralDescription2")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription2"].DataPropertyName = "GenDesc2";
                    d.Columns.Add("GeneralDescription3", dt2.Rows.Find("GeneralDescription3")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription3"].DataPropertyName = "GenDesc3";
                    d.Columns.Add("GeneralDescription4", dt2.Rows.Find("GeneralDescription4")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription4"].DataPropertyName = "GenDesc4";
                    d.Columns.Add("GeneralDescription5", dt2.Rows.Find("GeneralDescription5")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription5"].DataPropertyName = "GenDesc5";
                    d.Columns.Add("GeneralDescription6", dt2.Rows.Find("GeneralDescription6")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription6"].DataPropertyName = "GenDesc6";
                    d.Columns.Add("GeneralDescription7", dt2.Rows.Find("GeneralDescription7")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription7"].DataPropertyName = "GenDesc7";
                    d.Columns.Add("GeneralDescription8", dt2.Rows.Find("GeneralDescription8")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription8"].DataPropertyName = "GenDesc8";
                    d.Columns.Add("GeneralDescription9", dt2.Rows.Find("GeneralDescription9")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription9"].DataPropertyName = "GenDesc9";
                    d.Columns.Add("GeneralDescription10", dt2.Rows.Find("GeneralDescription10")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription10"].DataPropertyName = "GenDesc10";
                    d.Columns.Add("GeneralDescription11", dt2.Rows.Find("GeneralDescription11")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription11"].DataPropertyName = "GenDesc11";
                    d.Columns.Add("GeneralDescription12", dt2.Rows.Find("GeneralDescription12")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription12"].DataPropertyName = "GenDesc12";
                    d.Columns.Add("GeneralDescription13", dt2.Rows.Find("GeneralDescription13")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription13"].DataPropertyName = "GenDesc13";
                    d.Columns.Add("GeneralDescription14", dt2.Rows.Find("GeneralDescription14")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription14"].DataPropertyName = "GenDesc14";
                    d.Columns.Add("GeneralDescription15", dt2.Rows.Find("GeneralDescription15")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription15"].DataPropertyName = "GenDesc15";
                    d.Columns.Add("GeneralDescription16", dt2.Rows.Find("GeneralDescription16")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription16"].DataPropertyName = "GenDesc16";
                    d.Columns.Add("GeneralDescription17", dt2.Rows.Find("GeneralDescription17")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription17"].DataPropertyName = "GenDesc17";
                    d.Columns.Add("GeneralDescription18", dt2.Rows.Find("GeneralDescription18")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription18"].DataPropertyName = "GenDesc18";
                    d.Columns.Add("GeneralDescription19", dt2.Rows.Find("GeneralDescription19")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription19"].DataPropertyName = "GenDesc19";
                    d.Columns.Add("GeneralDescription20", dt2.Rows.Find("GeneralDescription20")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription20"].DataPropertyName = "GenDesc20";
                    d.Columns.Add("GeneralDescription21", dt2.Rows.Find("GeneralDescription21")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription21"].DataPropertyName = "GenDesc21";
                    d.Columns.Add("GeneralDescription22", dt2.Rows.Find("GeneralDescription22")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription22"].DataPropertyName = "GenDesc22";
                    d.Columns.Add("GeneralDescription23", dt2.Rows.Find("GeneralDescription23")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription23"].DataPropertyName = "GenDesc23";
                    d.Columns.Add("GeneralDescription24", dt2.Rows.Find("GeneralDescription24")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription24"].DataPropertyName = "GenDesc24";
                    d.Columns.Add("GeneralDescription25", dt2.Rows.Find("GeneralDescription25")["UserFriendlyName"].ToString());
                    d.Columns["GeneralDescription25"].DataPropertyName = "GenDesc25";
                }
                else
                {
                    d.Columns.Add("GeneralDescription1", "Gen Desc1");
                    d.Columns["GeneralDescription1"].DataPropertyName = "GenDesc1";
                    d.Columns.Add("GeneralDescription2", "Gen Desc2");
                    d.Columns["GeneralDescription2"].DataPropertyName = "GenDesc2";
                    d.Columns.Add("GeneralDescription3", "Gen Desc3");
                    d.Columns["GeneralDescription3"].DataPropertyName = "GenDesc3";
                    d.Columns.Add("GeneralDescription4", "Gen Desc4");
                    d.Columns["GeneralDescription4"].DataPropertyName = "GenDesc4";
                    d.Columns.Add("GeneralDescription5", "Gen Desc5");
                    d.Columns["GeneralDescription5"].DataPropertyName = "GenDesc5";
                    d.Columns.Add("GeneralDescription6", "Gen Desc6");
                    d.Columns["GeneralDescription6"].DataPropertyName = "GenDesc6";
                    d.Columns.Add("GeneralDescription7", "Gen Desc7");
                    d.Columns["GeneralDescription7"].DataPropertyName = "GenDesc7";
                    d.Columns.Add("GeneralDescription8", "Gen Desc8");
                    d.Columns["GeneralDescription8"].DataPropertyName = "GenDesc8";
                    d.Columns.Add("GeneralDescription9", "Gen Desc9");
                    d.Columns["GeneralDescription9"].DataPropertyName = "GenDesc9";
                    d.Columns.Add("GeneralDescription10", "Gen Desc10");
                    d.Columns["GeneralDescription10"].DataPropertyName = "GenDesc10";
                    d.Columns.Add("GeneralDescription11", "Gen Desc11");
                    d.Columns["GeneralDescription11"].DataPropertyName = "GenDesc11";
                    d.Columns.Add("GeneralDescription12", "Gen Desc12");
                    d.Columns["GeneralDescription12"].DataPropertyName = "GenDesc12";
                    d.Columns.Add("GeneralDescription13", "Gen Desc13");
                    d.Columns["GeneralDescription13"].DataPropertyName = "GenDesc13";
                    d.Columns.Add("GeneralDescription14", "Gen Desc14");
                    d.Columns["GeneralDescription14"].DataPropertyName = "GenDesc14";
                    d.Columns.Add("GeneralDescription15", "Gen Desc15");
                    d.Columns["GeneralDescription15"].DataPropertyName = "GenDesc15";
                    d.Columns.Add("GeneralDescription16", "Gen Desc16");
                    d.Columns["GeneralDescription16"].DataPropertyName = "GenDesc16";
                    d.Columns.Add("GeneralDescription17", "Gen Desc17");
                    d.Columns["GeneralDescription17"].DataPropertyName = "GenDesc17";
                    d.Columns.Add("GeneralDescription18", "Gen Desc18");
                    d.Columns["GeneralDescription18"].DataPropertyName = "GenDesc18";
                    d.Columns.Add("GeneralDescription19", "Gen Desc19");
                    d.Columns["GeneralDescription19"].DataPropertyName = "GenDesc19";
                    d.Columns.Add("GeneralDescription20", "Gen Desc20");
                    d.Columns["GeneralDescription20"].DataPropertyName = "GenDesc20";
                    d.Columns.Add("GeneralDescription21", "Gen Desc21");
                    d.Columns["GeneralDescription21"].DataPropertyName = "GenDesc21";
                    d.Columns.Add("GeneralDescription22", "Gen Desc22");
                    d.Columns["GeneralDescription22"].DataPropertyName = "GenDesc22";
                    d.Columns.Add("GeneralDescription23", "Gen Desc23");
                    d.Columns["GeneralDescription23"].DataPropertyName = "GenDesc23";
                    d.Columns.Add("GeneralDescription24", "Gen Desc24");
                    d.Columns["GeneralDescription24"].DataPropertyName = "GenDesc24";
                    d.Columns.Add("GeneralDescription25", "Gen Desc25");
                    d.Columns["GeneralDescription25"].DataPropertyName = "GenDesc25";
                }
                d.Columns.Add("TransactionAmount", "Trans Amount");
                d.Columns["TransactionAmount"].DataPropertyName = "TransactionAmount";
                d.Columns.Add("CurrencyCode", "Currency");
                d.Columns["CurrencyCode"].DataPropertyName = "CurrencyCode";
                d.Columns.Add("BaseAmount", "Base Amount");
                d.Columns["BaseAmount"].DataPropertyName = "BaseAmount";
                d.Columns.Add("Base2ReportingAmount", "2nd Base");
                d.Columns["Base2ReportingAmount"].DataPropertyName = "Base2ReportingAmount";
                d.Columns.Add("Value4Amount", "4th Amount");
                d.Columns["Value4Amount"].DataPropertyName = "Value4Amount";
                d.Columns.Add("Actions", "Actions");
                d.Columns["Actions"].DataPropertyName = "Actions";
                d.Columns.Add("Messages", "Messages");
                d.Columns["Messages"].DataPropertyName = "Messages";
                d.Columns.Add("Reference", "Ref");
                d.Columns["Reference"].DataPropertyName = "Reference";
                d.Columns["Reference"].Visible = false;
                d.Columns.Add("SaveReference", "SaveReference");
                d.Columns["SaveReference"].DataPropertyName = "SaveReference";
                d.Columns["SaveReference"].Visible = false;
                d.Columns.Add("populatecellwithJN", "populatecellwithJN");
                d.Columns["populatecellwithJN"].DataPropertyName = "populatecellwithJN";
                d.Columns["populatecellwithJN"].Visible = false;
                d.Columns.Add("BalanceBy", "BalanceBy");
                d.Columns["BalanceBy"].DataPropertyName = "BalanceBy";
                d.Columns["BalanceBy"].Visible = false;
                d.Columns.Add("AccountRange", "AccountRange");
                d.Columns["AccountRange"].DataPropertyName = "AccountRange";
                d.Columns["AccountRange"].Visible = false;
                d.Columns.Add("StartInCell", "StartInCell");
                d.Columns["StartInCell"].DataPropertyName = "StartInCell";
                d.Columns["StartInCell"].Visible = false;
            }
            return d;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void XMLFormdataGridView_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            DataTable dt = GetUserDataFriendlyName();
            DataColumn[] keys = new DataColumn[1];
            keys[0] = dt.Columns["SunField"];
            dt.PrimaryKey = keys;
            DataTable dt2 = GetUserDataGenFriendlyName();
            DataColumn[] keys2 = new DataColumn[1];
            keys2[0] = dt2.Columns["SunField"];
            dt2.PrimaryKey = keys2;
            if (dt.Rows.Count > 0)
            {
                if (((DataGridView)sender).Columns["Ledger"] != null)
                    ((DataGridView)sender).Columns["Ledger"].Visible = dt.Rows.Find("Ledger")["Output"].ToString() == "True" ? true : false;
                if (((DataGridView)sender).Columns["AccountCode"] != null)
                    ((DataGridView)sender).Columns["AccountCode"].Visible = dt.Rows.Find("AccountCode")["Output"].ToString() == "True" ? true : false;
                if (((DataGridView)sender).Columns["AccountingPeriod"] != null)
                    ((DataGridView)sender).Columns["AccountingPeriod"].Visible = dt.Rows.Find("AccountingPeriod")["Output"].ToString() == "True" ? true : false;
                if (((DataGridView)sender).Columns["TransactionDate"] != null)
                    ((DataGridView)sender).Columns["TransactionDate"].Visible = dt.Rows.Find("TransactionDate")["Output"].ToString() == "True" ? true : false;
                if (((DataGridView)sender).Columns["DueDate"] != null)
                    ((DataGridView)sender).Columns["DueDate"].Visible = dt.Rows.Find("DueDate")["Output"].ToString() == "True" ? true : false;
                if (((DataGridView)sender).Columns["JournalType"] != null)
                    ((DataGridView)sender).Columns["JournalType"].Visible = dt.Rows.Find("JournalType")["Output"].ToString() == "True" ? true : false;
                if (((DataGridView)sender).Columns["JournalSource"] != null)
                    ((DataGridView)sender).Columns["JournalSource"].Visible = dt.Rows.Find("JournalSource")["Output"].ToString() == "True" ? true : false;
                if (((DataGridView)sender).Columns["TransactionReference"] != null)
                    ((DataGridView)sender).Columns["TransactionReference"].Visible = dt.Rows.Find("TransactionReference")["Output"].ToString() == "True" ? true : false;
                if (((DataGridView)sender).Columns["Description"] != null)
                    ((DataGridView)sender).Columns["Description"].Visible = dt.Rows.Find("Description")["Output"].ToString() == "True" ? true : false;
                if (((DataGridView)sender).Columns["AllocationMarker"] != null)
                    ((DataGridView)sender).Columns["AllocationMarker"].Visible = dt.Rows.Find("AllocationMarker")["Output"].ToString() == "True" ? true : false;
                if (((DataGridView)sender).Columns["AnalysisCode1"] != null)
                    ((DataGridView)sender).Columns["AnalysisCode1"].Visible = dt.Rows.Find("AnalysisCode1")["Output"].ToString() == "True" ? true : false;
                if (((DataGridView)sender).Columns["AnalysisCode2"] != null)
                    ((DataGridView)sender).Columns["AnalysisCode2"].Visible = dt.Rows.Find("AnalysisCode2")["Output"].ToString() == "True" ? true : false;
                if (((DataGridView)sender).Columns["AnalysisCode3"] != null)
                    ((DataGridView)sender).Columns["AnalysisCode3"].Visible = dt.Rows.Find("AnalysisCode3")["Output"].ToString() == "True" ? true : false;
                if (((DataGridView)sender).Columns["AnalysisCode4"] != null)
                    ((DataGridView)sender).Columns["AnalysisCode4"].Visible = dt.Rows.Find("AnalysisCode4")["Output"].ToString() == "True" ? true : false;
                if (((DataGridView)sender).Columns["AnalysisCode5"] != null)
                    ((DataGridView)sender).Columns["AnalysisCode5"].Visible = dt.Rows.Find("AnalysisCode5")["Output"].ToString() == "True" ? true : false;
                if (((DataGridView)sender).Columns["AnalysisCode6"] != null)
                    ((DataGridView)sender).Columns["AnalysisCode6"].Visible = dt.Rows.Find("AnalysisCode6")["Output"].ToString() == "True" ? true : false;
                if (((DataGridView)sender).Columns["AnalysisCode7"] != null)
                    ((DataGridView)sender).Columns["AnalysisCode7"].Visible = dt.Rows.Find("AnalysisCode7")["Output"].ToString() == "True" ? true : false;
                if (((DataGridView)sender).Columns["AnalysisCode8"] != null)
                    ((DataGridView)sender).Columns["AnalysisCode8"].Visible = dt.Rows.Find("AnalysisCode8")["Output"].ToString() == "True" ? true : false;
                if (((DataGridView)sender).Columns["AnalysisCode9"] != null)
                    ((DataGridView)sender).Columns["AnalysisCode9"].Visible = dt.Rows.Find("AnalysisCode9")["Output"].ToString() == "True" ? true : false;
                if (((DataGridView)sender).Columns["AnalysisCode10"] != null)
                    ((DataGridView)sender).Columns["AnalysisCode10"].Visible = dt.Rows.Find("AnalysisCode10")["Output"].ToString() == "True" ? true : false;
                if (dt2.Rows.Count > 0)
                {
                    if (((DataGridView)sender).Columns["GeneralDescription1"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription1"].Visible = dt2.Rows.Find("GeneralDescription1")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription2"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription2"].Visible = dt2.Rows.Find("GeneralDescription2")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription3"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription3"].Visible = dt2.Rows.Find("GeneralDescription3")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription4"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription4"].Visible = dt2.Rows.Find("GeneralDescription4")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription5"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription5"].Visible = dt2.Rows.Find("GeneralDescription5")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription6"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription6"].Visible = dt2.Rows.Find("GeneralDescription6")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription7"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription7"].Visible = dt2.Rows.Find("GeneralDescription7")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription8"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription8"].Visible = dt2.Rows.Find("GeneralDescription8")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription9"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription9"].Visible = dt2.Rows.Find("GeneralDescription9")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription10"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription10"].Visible = dt2.Rows.Find("GeneralDescription10")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription11"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription11"].Visible = dt2.Rows.Find("GeneralDescription11")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription12"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription12"].Visible = dt2.Rows.Find("GeneralDescription12")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription13"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription13"].Visible = dt2.Rows.Find("GeneralDescription13")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription14"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription14"].Visible = dt2.Rows.Find("GeneralDescription14")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription15"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription15"].Visible = dt2.Rows.Find("GeneralDescription15")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription16"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription16"].Visible = dt2.Rows.Find("GeneralDescription16")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription17"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription17"].Visible = dt2.Rows.Find("GeneralDescription17")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription18"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription18"].Visible = dt2.Rows.Find("GeneralDescription18")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription19"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription19"].Visible = dt2.Rows.Find("GeneralDescription19")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription20"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription20"].Visible = dt2.Rows.Find("GeneralDescription20")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription21"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription21"].Visible = dt2.Rows.Find("GeneralDescription21")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription22"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription22"].Visible = dt2.Rows.Find("GeneralDescription22")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription23"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription23"].Visible = dt2.Rows.Find("GeneralDescription23")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription24"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription24"].Visible = dt2.Rows.Find("GeneralDescription24")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription25"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription25"].Visible = dt2.Rows.Find("GeneralDescription25")["Output"].ToString() == "True" ? true : false;
                }
                else
                {
                    if (((DataGridView)sender).Columns["GeneralDescription1"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription1"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription2"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription2"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription3"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription3"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription4"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription4"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription5"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription5"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription6"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription6"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription7"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription7"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription8"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription8"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription9"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription9"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription10"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription10"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription11"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription11"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription12"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription12"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription13"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription13"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription14"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription14"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription15"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription15"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription16"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription16"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription17"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription17"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription18"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription18"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription19"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription19"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription20"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription20"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription21"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription21"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription22"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription22"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription23"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription23"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription24"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription24"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription25"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription25"].Visible = false;
                }
                if (((DataGridView)sender).Columns["TransactionAmount"] != null)
                    ((DataGridView)sender).Columns["TransactionAmount"].Visible = dt.Rows.Find("TransactionAmount")["Output"].ToString() == "True" ? true : false;
                if (((DataGridView)sender).Columns["CurrencyCode"] != null)
                    ((DataGridView)sender).Columns["CurrencyCode"].Visible = dt.Rows.Find("CurrencyCode")["Output"].ToString() == "True" ? true : false;
                if (((DataGridView)sender).Columns["BaseAmount"] != null)
                    ((DataGridView)sender).Columns["BaseAmount"].Visible = dt.Rows.Find("BaseAmount")["Output"].ToString() == "True" ? true : false;
                if (((DataGridView)sender).Columns["Base2ReportingAmount"] != null)
                    ((DataGridView)sender).Columns["Base2ReportingAmount"].Visible = dt.Rows.Find("Base2ReportingAmount")["Output"].ToString() == "True" ? true : false;
                if (((DataGridView)sender).Columns["Value4Amount"] != null)
                    ((DataGridView)sender).Columns["Value4Amount"].Visible = dt.Rows.Find("Value4Amount")["Output"].ToString() == "True" ? true : false;
            }
            else
            {
                if (dt2.Rows.Count > 0)
                {
                    if (((DataGridView)sender).Columns["GeneralDescription1"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription1"].Visible = dt2.Rows.Find("GeneralDescription1")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription2"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription2"].Visible = dt2.Rows.Find("GeneralDescription2")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription3"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription3"].Visible = dt2.Rows.Find("GeneralDescription3")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription4"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription4"].Visible = dt2.Rows.Find("GeneralDescription4")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription5"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription5"].Visible = dt2.Rows.Find("GeneralDescription5")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription6"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription6"].Visible = dt2.Rows.Find("GeneralDescription6")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription7"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription7"].Visible = dt2.Rows.Find("GeneralDescription7")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription8"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription8"].Visible = dt2.Rows.Find("GeneralDescription8")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription9"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription9"].Visible = dt2.Rows.Find("GeneralDescription9")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription10"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription10"].Visible = dt2.Rows.Find("GeneralDescription10")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription11"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription11"].Visible = dt2.Rows.Find("GeneralDescription11")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription12"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription12"].Visible = dt2.Rows.Find("GeneralDescription12")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription13"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription13"].Visible = dt2.Rows.Find("GeneralDescription13")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription14"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription14"].Visible = dt2.Rows.Find("GeneralDescription14")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription15"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription15"].Visible = dt2.Rows.Find("GeneralDescription15")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription16"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription16"].Visible = dt2.Rows.Find("GeneralDescription16")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription17"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription17"].Visible = dt2.Rows.Find("GeneralDescription17")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription18"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription18"].Visible = dt2.Rows.Find("GeneralDescription18")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription19"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription19"].Visible = dt2.Rows.Find("GeneralDescription19")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription20"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription20"].Visible = dt2.Rows.Find("GeneralDescription20")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription21"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription21"].Visible = dt2.Rows.Find("GeneralDescription21")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription22"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription22"].Visible = dt2.Rows.Find("GeneralDescription22")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription23"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription23"].Visible = dt2.Rows.Find("GeneralDescription23")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription24"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription24"].Visible = dt2.Rows.Find("GeneralDescription24")["Output"].ToString() == "True" ? true : false;
                    if (((DataGridView)sender).Columns["GeneralDescription25"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription25"].Visible = dt2.Rows.Find("GeneralDescription25")["Output"].ToString() == "True" ? true : false;
                }
                else
                {
                    if (((DataGridView)sender).Columns["GeneralDescription1"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription1"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription2"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription2"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription3"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription3"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription4"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription4"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription5"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription5"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription6"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription6"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription7"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription7"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription8"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription8"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription9"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription9"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription10"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription10"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription11"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription11"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription12"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription12"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription13"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription13"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription14"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription14"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription15"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription15"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription16"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription16"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription17"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription17"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription18"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription18"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription19"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription19"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription20"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription20"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription21"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription21"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription22"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription22"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription23"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription23"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription24"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription24"].Visible = false;
                    if (((DataGridView)sender).Columns["GeneralDescription25"] != null)
                        ((DataGridView)sender).Columns["GeneralDescription25"].Visible = false;
                }
            }
        }
    }
    /// <summary>
    /// power by peter 20141112
    /// </summary>
    public class DataConversionTools
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="PropertyName"></param>
        /// <returns></returns>
        public static bool IsPropertyInClassProperties<T>(string PropertyName)
        {
            bool returnValue = false;
            try
            {
                System.Reflection.PropertyInfo[] myPropertyInfo = typeof(T).GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                for (int i = 0, j = myPropertyInfo.Length; i < j; i++)
                {
                    System.Reflection.PropertyInfo pi = myPropertyInfo[i];
                    string name = pi.Name;
                    if (name == PropertyName)
                    {
                        return true;
                    }
                }
                return returnValue;
            }
            catch { return false; }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="PropertyName"></param>
        /// <param name="t"></param>
        /// <returns></returns>
        public static string GetPropertyValue<T>(string PropertyName, T t)
        {
            string returnValue = string.Empty;
            try
            {
                System.Reflection.PropertyInfo[] myPropertyInfo = typeof(T).GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                for (int i = 0, j = myPropertyInfo.Length; i < j; i++)
                {
                    System.Reflection.PropertyInfo pi = myPropertyInfo[i];
                    string name = pi.Name;
                    if (name == PropertyName)
                    {
                        returnValue = pi.GetValue(t, null).ToString();
                    }
                }
                return returnValue;
            }
            catch { return ""; }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="PropertyName"></param>
        /// <param name="value"></param>
        /// <param name="t"></param>
        public static void SetPropertyValue<T>(string PropertyName, string value, ref T t)
        {
            System.Reflection.PropertyInfo[] myPropertyInfo = typeof(T).GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
            for (int i = 0, j = myPropertyInfo.Length; i < j; i++)
            {
                System.Reflection.PropertyInfo pi = myPropertyInfo[i];
                string name = pi.Name;
                if (name == PropertyName)
                {
                    pi.SetValue(t, value, null);
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <returns></returns>
        public static DataSet ConvertToDataSet<T>(IList<T> list)
        {
            if (list == null || list.Count <= 0)
            {
                return null;
            }
            DataSet ds = new DataSet();
            DataTable dt = new DataTable(typeof(T).Name);
            DataColumn column;
            DataRow row;
            System.Reflection.PropertyInfo[] myPropertyInfo = typeof(T).GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
            foreach (T t in list)
            {
                if (t == null)
                {
                    continue;
                }
                row = dt.NewRow();
                for (int i = 0, j = myPropertyInfo.Length; i < j; i++)
                {
                    System.Reflection.PropertyInfo pi = myPropertyInfo[i];
                    string name = pi.Name;
                    if (dt.Columns[name] == null)
                    {
                        column = new DataColumn(name, pi.PropertyType);
                        dt.Columns.Add(column);
                    }
                    row[name] = pi.GetValue(t, null);
                }
                dt.Rows.Add(row);
            }
            ds.Tables.Add(dt);
            return ds;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static DataTable ConvertToDataTableStructure<T>()
        {
            DataTable dt = new DataTable(typeof(T).Name);
            DataColumn column;
            System.Reflection.PropertyInfo[] myPropertyInfo = typeof(T).GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
            for (int i = 0, j = myPropertyInfo.Length; i < j; i++)
            {
                System.Reflection.PropertyInfo pi = myPropertyInfo[i];
                string name = pi.Name;
                if (dt.Columns[name] == null)
                {
                    column = new DataColumn(name, pi.PropertyType);
                    dt.Columns.Add(column);
                }
            }
            return dt;
        }
    }
    ///// <summary>
    ///// 
    ///// </summary>
    //public static class Pathing
    //{
    //    [DllImport("mpr.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    //    public static extern int WNetGetConnection(
    //        [MarshalAs(UnmanagedType.LPTStr)] string localName,
    //        [MarshalAs(UnmanagedType.LPTStr)] StringBuilder remoteName,
    //        ref int length);
    //    /// <summary>   
    //    /// Given a path, the network path back or the original path.
    //    /// For example: the given path P:/2008 year in February 29th (P: as a mapped network drive name), may return: "//networkserver/ photo /2008 year in February 9th"
    //    /// </summary>   
    //    /// <param name="originalPath">The specified path</param>   
    //    /// <returns>If it is a local path, the return value and the incoming parameters value; if it is a local mapped network drives</returns>   
    //    public static string GetUNCPath(string originalPath)
    //    {
    //        StringBuilder sb = new StringBuilder(512);
    //        int size = sb.Capacity;
    //        if (originalPath.Length > 2 && originalPath[1] == ':')
    //        {
    //            char c = originalPath[0];
    //            if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z'))
    //            {
    //                int error = WNetGetConnection(originalPath.Substring(0, 2),
    //                    sb, ref size);

    //                //if (error == 0)
    //                //{
    //                DirectoryInfo dir = new DirectoryInfo(originalPath);
    //                string path = Path.GetFullPath(originalPath)
    //                    .Substring(Path.GetPathRoot(originalPath).Length);
    //                return Path.Combine(sb.ToString().TrimEnd(), path);
    //                //}
    //            }
    //        }
    //        return originalPath;
    //    }
    //}
}
