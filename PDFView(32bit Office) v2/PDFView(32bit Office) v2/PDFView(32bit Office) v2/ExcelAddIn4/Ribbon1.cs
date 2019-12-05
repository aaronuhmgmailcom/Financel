/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<Ribbon2>   
 * Author：Peter.uhm  (yanb@shinetechchina.com)
 * Modify date：2016.04
 * Modify date：2016.09
 * Version : 2.0.0.2
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;
using Microsoft.Office.Core;
using System.Security.Principal;
using ExcelAddIn4.Common;
using System.Xml;
using System.Text.RegularExpressions;
namespace ExcelAddIn4
{
    public partial class Ribbon2
    {
        /// <summary>
        /// 
        /// </summary>
        internal static Microsoft.Office.Tools.CustomTaskPane _MyCustomTaskPane = null;
        /// <summary>
        /// 
        /// </summary>
        internal static Microsoft.Office.Tools.CustomTaskPane _MyOutputCustomTaskPane = null;
        /// <summary>
        /// 
        /// </summary>
        internal static Microsoft.Office.Tools.CustomTaskPane _MyHelpCustomTaskPane = null;
        /// <summary>
        /// 
        /// </summary>
        internal static Microsoft.Office.Interop.Excel.Worksheet wsRrigin = null;
        /// <summary>
        /// 
        /// </summary>
        private UCForTaskPane taskPane = null;
        /// <summary>
        /// 
        /// </summary>
        internal static OutputContainer outputPane = null;
        /// <summary>
        /// 
        /// </summary>
        private HelpContainer helpPane = null;
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
        internal static System.Windows.Forms.NotifyIcon notifyIcon1 = null;
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
        internal static string LastColumnName
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        internal static string LastRowNumber
        {
            get;
            set;
        }
        /// <summary>
        /// 
        /// </summary>
        internal static XMLPostFrm xpf;
        /// <summary>
        /// 
        /// </summary>
        internal static TransUpdPostFrm tupf;
        /// <summary>
        /// 
        /// </summary>
        internal static CreateTextFileForm ctff;
        /// <summary>
        /// 
        /// </summary>
        internal static List<KeyValuePair<string, string>> TemplateAndPath = new List<KeyValuePair<string, string>>();
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Ribbon2_Load(object sender, RibbonUIEventArgs e)
        {
            InitializeCustomTaskPanes();
            Globals.ThisAddIn.Application.SheetSelectionChange += new Microsoft.Office.Interop.Excel.AppEvents_SheetSelectionChangeEventHandler(ThisWorkbook_SheetSelectionChange);
            Globals.ThisAddIn.Application.WorkbookActivate += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookActivateEventHandler(Application_WorkbookActivate);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileid"></param>
        /// <param name="filename"></param>
        public static void InitializeNewButtons(string fileid, string filename)
        {
            for (int i = 0; i < ButtonViewCount; i++)
            {
                try
                {
                    RibbonControl rb = (RibbonButton)Globals.Ribbons[0].Tabs[0].Groups[2].Items[i];
                    rb.Visible = false;
                }
                catch { }
            }
            System.Collections.Generic.List<string> list = ft.GetUserGroups(SessionInfo.UserInfo.ID);
            for (int j = 0; j < list.Count; j++)
            {
                string groupid = list[j];
                DataTable dt = ft.GetGroupButtonsView(groupid, fileid);
                for (int i = 0; i < ButtonViewCount; i++)
                {
                    try
                    {
                        RibbonButton rb = (RibbonButton)Globals.Ribbons[0].Tabs[0].Groups[2].Items[i];
                        string[] sArray = Regex.Split(rb.Tag.ToString(), ",");
                        string templateid = sArray[0];
                        string actionid = sArray[1];
                        DataRow[] dr = dt.Select(" TemplateID = '" + templateid + "' and  ActionID = '" + actionid + "'");
                        if (dr.Length > 0)
                            rb.Visible = true;
                    }
                    catch { }
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        private void InitializeCustomTaskPanes()
        {
            if (Globals.ThisAddIn.CustomTaskPanes.Count != 0)
            {
                int num = Globals.ThisAddIn.CustomTaskPanes.Count;
                for (int i = 0; i < num; i++)
                    Globals.ThisAddIn.CustomTaskPanes.RemoveAt(0);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileName">Current file clicked in RDDI</param>
        /// <param name="filePath">Current file path clicked in RDDI</param>
        private void InitializeUserSessionInfo(string fileName, string filePath)
        {
            if (string.IsNullOrEmpty(SessionInfo.UserInfo.FilePath))
            {
                SessionInfo.UserInfo.FilePath = filePath;
                SessionInfo.UserInfo.CachePath = filePath;
                SessionInfo.UserInfo.FileName = Path.GetFileNameWithoutExtension(fileName);
                if (!ft.IsGUID(Path.GetFileNameWithoutExtension(filePath)))
                    SessionInfo.UserInfo.File_ftid = ft.GetTemplateIDByName(Path.GetFileNameWithoutExtension(filePath)).ToString();
            }
            else if (!filePath.Contains("RSDataCache") && !ft.IsGUID(Path.GetFileNameWithoutExtension(filePath)))
            {
                SessionInfo.UserInfo.FilePath = filePath;
                SessionInfo.UserInfo.CachePath = filePath;
                SessionInfo.UserInfo.FileName = Path.GetFileNameWithoutExtension(fileName);
                if (!ft.IsGUID(Path.GetFileNameWithoutExtension(filePath)))
                    SessionInfo.UserInfo.File_ftid = ft.GetTemplateIDByName(Path.GetFileNameWithoutExtension(filePath)).ToString();
            }
            else
            {
                SessionInfo.UserInfo.CachePath = filePath;
                if (!SessionInfo.UserInfo.Dictionary.dict.ContainsKey(SessionInfo.UserInfo.CachePath))
                    SessionInfo.UserInfo.Dictionary.dict.Add(SessionInfo.UserInfo.CachePath, SessionInfo.UserInfo.File_ftid + "," + SessionInfo.UserInfo.InvNumber);
                else
                {
                    SessionInfo.UserInfo.File_ftid = SessionInfo.UserInfo.Dictionary.dict[SessionInfo.UserInfo.CachePath];
                    string[] sArray = Regex.Split(SessionInfo.UserInfo.File_ftid, ",");
                    SessionInfo.UserInfo.File_ftid = sArray[0];
                    SessionInfo.UserInfo.InvNumber = int.Parse(sArray[1]);
                    SessionInfo.UserInfo.FilePath = ft.getFilePath(SessionInfo.UserInfo.File_ftid);
                }
                SessionInfo.UserInfo.FileName = Path.GetFileNameWithoutExtension(SessionInfo.UserInfo.FilePath);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="flag"></param>
        private void InitializePDFContainerPane(bool flag)
        {
            _MyCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(taskPane, "PDF Container");
            int w = Convert.ToInt32(Globals.ThisAddIn.Application.ActiveWindow.Width) / 2;
            _MyCustomTaskPane.Control.Dock = System.Windows.Forms.DockStyle.Fill;
            _MyCustomTaskPane.Width = w;
            _MyCustomTaskPane.Visible = flag;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="flag"></param>
        private void InitializeOutPutContainerPane(bool flag)
        {
            try
            {
                outputPane = new OutputContainer();
                _MyOutputCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(outputPane, "Output settings");
                int x = Convert.ToInt32(Globals.ThisAddIn.Application.ActiveWindow.Height) / 2;
                _MyOutputCustomTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionBottom;
                _MyOutputCustomTaskPane.Height = x;
                _MyOutputCustomTaskPane.Visible = flag;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Output settings Error");
                LogHelper.WriteLog(typeof(Ribbon2), ex.Message + "Ribbon Error");
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="flag"></param>
        private void InitializeHelpContainerPane(bool flag)
        {
            try
            {
                helpPane = new HelpContainer();
                _MyHelpCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(helpPane, "Help Container");
                int w = Convert.ToInt32(Globals.ThisAddIn.Application.ActiveWindow.Width) / 2;
                _MyHelpCustomTaskPane.Control.Dock = System.Windows.Forms.DockStyle.Fill;
                _MyHelpCustomTaskPane.Width = w;
                _MyHelpCustomTaskPane.Visible = flag;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Help Container Error");
                LogHelper.WriteLog(typeof(Ribbon2), ex.Message + "Ribbon Error");
            }
        }
        /// <summary>
        /// RibbonToggle Button Click event (open workbooks)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rt_click(object sender, RibbonControlEventArgs e)
        {
            RibbonToggleButton btn = sender as RibbonToggleButton;
            RibbonControl rc = btn.Parent as RibbonControl;
            var xlapp = Globals.ThisAddIn.Application;
            var path = Finance_Tools.RootPath;
            if (path != null)
            {
                string btnname = path + "\\" + rc.Tag + "\\" + btn.Label + ".xlsm";
                SessionInfo.UserInfo.File_ftid = btn.Tag.ToString();
                try
                {
                    xlapp.Workbooks.Open(btnname);
                }
                catch { xlapp.Workbooks.Open(btnname.Replace(".xlsm", ".xlsx")); btnname = path + "\\" + rc.Tag + "\\" + btn.Label + ".xlsx"; }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="filepath"></param>
        private void SyncPanes(string filename, string filepath)
        {
            try
            {
                #region Initialize the CustomTaskPanes
                InitializeCustomTaskPanes();
                #endregion
                #region Initialize the user information and set them into Session
                InitializeUserSessionInfo(filename, filepath);
                #endregion
                #region Initialize the NewButtons
                InitializeNewButtons(SessionInfo.UserInfo.File_ftid, SessionInfo.UserInfo.FileName);
                #endregion
                taskPane = new UCForTaskPane();
                #region Initialize PDF Container Pane
                InitializePDFContainerPane(false);
                this.btnATVP_View.Checked = false;
                #endregion
                #region Initialize Setting Container Pane
                if (ft.IsOutPutPaneVisiable(SessionInfo.UserInfo.File_ftid))
                    InitializeOutPutContainerPane(true);
                else
                    InitializeOutPutContainerPane(false);
                #endregion
                #region Initialize Help Container Pane
                InitializeHelpContainerPane(false);
                #endregion
            }
            catch
            {
                Microsoft.Office.Tools.Ribbon.RibbonCheckBox rcb = (Microsoft.Office.Tools.Ribbon.RibbonCheckBox)Globals.Ribbons[0].Tabs[0].Groups[5].Items[0];
                rcb.Checked = false;
                rcb = (Microsoft.Office.Tools.Ribbon.RibbonCheckBox)Globals.Ribbons[0].Tabs[0].Groups[5].Items[1];
                rcb.Checked = false;
                #region Initialize Setting Container Pane
                if (ft.IsOutPutPaneVisiable(SessionInfo.UserInfo.File_ftid))
                    InitializeOutPutContainerPane(true);
                else
                    InitializeOutPutContainerPane(false);
                #endregion
                #region Initialize Help Container Pane
                InitializeHelpContainerPane(false);
                #endregion
            }
            finally
            {
                _MyOutputCustomTaskPane.VisibleChanged += new EventHandler(OutputContainer_VisibleChanged);
                OutputContainer_VisibleChanged(null, null);
                _MyHelpCustomTaskPane.VisibleChanged += new EventHandler(HelpContainer_VisibleChanged);
                HelpContainer_VisibleChanged(null, null);
                _MyCustomTaskPane.VisibleChanged += new EventHandler(UCForTaskPane_VisibleChanged);
                UCForTaskPane_VisibleChanged(null, null);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="containerPath"></param>
        private void SetPDFViewer(string containerPath)
        {
            taskPane.pdfViewer1.AllowBookmarks = true;
            taskPane.pdfViewer1.AutoSize = true;
            taskPane.pdfViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            taskPane.pdfViewer1.Location = new System.Drawing.Point(0, 30);
            taskPane.pdfViewer1.Name = "pdfViewer1";
            taskPane.pdfViewer1.TabIndex = 0;
            taskPane.pdfViewer1.UseXPDF = true;
            taskPane.pdfViewer1.FileName = containerPath;
            SessionInfo.UserInfo.Containerpath = containerPath;
        }
        /// <summary>
        /// Save Template Button Click event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTemplate_Save_Click(object sender, RibbonControlEventArgs e)
        {
            FileNameForm frm = new FileNameForm();
            frm.ShowDialog();
        }
        /// <summary>
        /// Data Fields Setting Button Click event 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAd_DataFieldsSetting_Click(object sender, RibbonControlEventArgs e)
        {
            DataFieldsSetting dfs = new DataFieldsSetting("");
            dfs.ShowDialog();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OutputContainer_VisibleChanged(object sender, EventArgs e)
        {
            Microsoft.Office.Tools.Ribbon.RibbonCheckBox rcb = (Microsoft.Office.Tools.Ribbon.RibbonCheckBox)Globals.Ribbons[0].Tabs[0].Groups[5].Items[1];
            try
            {
                if (_MyOutputCustomTaskPane.Visible == false)
                    rcb.Checked = false;
                else
                    rcb.Checked = true;
            }
            catch
            {
                rcb.Checked = false;
            }
            finally
            {
                if (sender != null)
                    SaveToRemember();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void HelpContainer_VisibleChanged(object sender, EventArgs e)
        {
            Microsoft.Office.Tools.Ribbon.RibbonCheckBox rcb = (Microsoft.Office.Tools.Ribbon.RibbonCheckBox)Globals.Ribbons[0].Tabs[0].Groups[5].Items[2];
            try
            {
                if (_MyHelpCustomTaskPane.Visible == false)
                    rcb.Checked = false;
                else
                {
                    rcb.Checked = true;
                    if (helpPane.webBrowser1.Url == null || helpPane.webBrowser1.Url.AbsolutePath == "blank")
                    {
                        string Url = ft.OpenHelpFile().Replace("\\\\", "\\");
                        helpPane.webBrowser1.Navigate(Url);
                        helpPane.webBrowser1.Document.ExecCommand("EditMode", true, "");
                    }
                }
            }
            catch
            {
                rcb.Checked = false;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        private void SaveToRemember()
        {
            if (!string.IsNullOrEmpty(SessionInfo.UserInfo.File_ftid))
            {
                Microsoft.Office.Tools.Ribbon.RibbonCheckBox rcbOutput = (Microsoft.Office.Tools.Ribbon.RibbonCheckBox)Globals.Ribbons[0].Tabs[0].Groups[5].Items[1];
                Microsoft.Office.Tools.Ribbon.RibbonCheckBox rcbPDF = (Microsoft.Office.Tools.Ribbon.RibbonCheckBox)Globals.Ribbons[0].Tabs[0].Groups[5].Items[0];
                SqlConnection conn = null;
                try
                {
                    conn = new
                        SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                    conn.Open();
                    SqlCommand cmd = new SqlCommand("rsTemplateVisible_Del", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@UserID", SessionInfo.UserInfo.ID));
                    cmd.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                    cmd.ExecuteNonQuery();
                    SqlCommand cmd2 = new SqlCommand("rsTemplateVisible_Ins", conn);
                    cmd2.CommandType = CommandType.StoredProcedure;
                    cmd2.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                    cmd2.Parameters.Add(new SqlParameter("@OutputPaneVisiable", rcbOutput.Checked));
                    cmd2.Parameters.Add(new SqlParameter("@UserID", SessionInfo.UserInfo.ID));
                    cmd2.ExecuteNonQuery();
                }
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
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UCForTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            Microsoft.Office.Tools.Ribbon.RibbonCheckBox rcb = (Microsoft.Office.Tools.Ribbon.RibbonCheckBox)Globals.Ribbons[0].Tabs[0].Groups[5].Items[0];
            try
            {
                if (_MyCustomTaskPane.Visible == false)
                    rcb.Checked = false;
                else
                    rcb.Checked = true;
            }
            catch
            {
                rcb.Checked = false;
            }
            finally
            {
                if (sender != null)
                    SaveToRemember();
            }
        }
        /// <summary>
        /// 'View' CheckBox Click event (Hide or show PDF Container) 
        ///  //IEnumerable<Microsoft.Office.Tools.CustomTaskPane> filtered2 = Globals.ThisAddIn.CustomTaskPanes.Where(s => s.Title.Contains("PDF Container"));
        //   foreach (var p in filtered2)
        //      {filtered2.Visible = true;}
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnATVP_View_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Microsoft.Office.Tools.CustomTaskPane filtered2 = Globals.ThisAddIn.CustomTaskPanes.First(s => s.Title.Contains("PDF Container"));
                if (this.btnATVP_View.Checked == true)
                {
                    filtered2.Visible = true;
                    bool IsWebBrowser = false; bool IsDataBaseQuery = false; string ConnectionString = string.Empty; string UserID = string.Empty; string Password = string.Empty; string SQLString = string.Empty;
                    GetPDFDetails(ref IsWebBrowser, ref IsDataBaseQuery, ref ConnectionString, ref UserID, ref Password, ref SQLString);
                    var xlapp = Globals.ThisAddIn.Application;
                    string cellvalue = string.Empty;
                    try
                    {
                        cellvalue = xlapp.ActiveCell.Value.ToString();
                    }
                    catch { return; }
                    if (IsDataBaseQuery)
                        OpenDataBaseQuery(IsWebBrowser, ConnectionString, UserID, Password, SQLString, cellvalue);
                    else
                    {
                        if (IsWebBrowser)
                            BroswerNavigate(cellvalue);
                        else
                        {
                            taskPane.Panel1.Visible = true;
                            taskPane.Panel2.Visible = true;
                            taskPane.pdfViewer1.Visible = true;
                            taskPane.panel3.Visible = false;
                            ThisWorkbook_SheetSelectionChange(null, Globals.ThisAddIn.Application.ActiveCell);
                        }
                    }
                }
                else
                {
                    filtered2.Visible = false;
                }
            }
            catch (Exception ex)
            { MessageBox.Show(ex.ToString(), "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void suspendToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("To abort the process may bring risk, are you sure to continue this operation?", "Alert - RSystems FinanceTools", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                == DialogResult.Yes)
            {
                if (!Process.GetCurrentProcess().HasExited)
                {
                    Globals.ThisAddIn.ThisAddIn_Shutdown(null, null);
                    Process.GetCurrentProcess().Kill();
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        private void ChangeXPFFormStatus()
        {
            try
            {
                if (Ribbon2.xpf != null)
                    if (Ribbon2.xpf.Visible == false)
                        if (Ribbon2.xpf.WindowState == FormWindowState.Minimized)
                            Ribbon2.xpf.WindowState = FormWindowState.Normal;
                        else
                            Ribbon2.xpf.WindowState = FormWindowState.Minimized;
                if (Ribbon2.ctff != null)
                    if (Ribbon2.ctff.Visible == false)
                        if (Ribbon2.ctff.WindowState == FormWindowState.Minimized)
                            Ribbon2.ctff.WindowState = FormWindowState.Normal;
                        else
                            Ribbon2.ctff.WindowState = FormWindowState.Minimized;
                if (Ribbon2.tupf != null)
                    if (Ribbon2.tupf.Visible == false)
                        if (Ribbon2.tupf.WindowState == FormWindowState.Minimized)
                            Ribbon2.tupf.WindowState = FormWindowState.Normal;
                        else
                            Ribbon2.tupf.WindowState = FormWindowState.Minimized;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog(typeof(Ribbon2), ex.Message + "Ribbon Error");
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellvalue"></param>
        private void BroswerNavigate(string cellvalue)
        {
            taskPane.Panel1.Visible = false;
            taskPane.Panel2.Visible = false;
            taskPane.pdfViewer1.Visible = false;
            taskPane.panel3.Visible = true;
            string Url = SessionInfo.UserInfo.Containerpath.Replace("\\\\", "\\");
            if (Url.EndsWith("//") || Url.EndsWith("/"))
                Url += cellvalue + ".html";
            else
                Url += "/" + cellvalue + ".html";

            taskPane.webBrowser1.Navigate(Url);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="IsWebBrowser"></param>
        /// <param name="ConnectionString"></param>
        /// <param name="UserID"></param>
        /// <param name="Password"></param>
        /// <param name="SQLString"></param>
        /// <param name="cellvalue"></param>
        private void OpenDataBaseQuery(bool IsWebBrowser, string ConnectionString, string UserID, string Password, string SQLString, string cellvalue)
        {
            string connString = ConnectionString.Replace("[UserID]", UserID).Replace("[Password]", Password);
            SqlConnection sqlConn = new SqlConnection(connString);
            DataSet ds = new DataSet();
            try
            {
                sqlConn.Open();
                SqlCommand cmd = new SqlCommand(SQLString.Replace("~", cellvalue), sqlConn);
                cmd.CommandType = CommandType.Text;
                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                {
                    try
                    {
                        da.Fill(ds, "ds");
                        cmd.Parameters.Clear();
                    }
                    catch (System.Data.SqlClient.SqlException ex)
                    {
                        throw new Exception(ex.Message);
                    }
                }
                if (IsWebBrowser && ds.Tables[0].Rows.Count > 0)
                {
                    SessionInfo.UserInfo.Containerpath = ds.Tables[0].Rows[0][0].ToString();
                    BroswerNavigate(cellvalue);
                }
                else if (!IsWebBrowser && ds.Tables[0].Rows.Count > 0)
                {
                    var data = Finance_Tools.Serialize(ds.Tables[0].Rows[0][0]);
                    var fileType = ".pdf";
                    string tmp = Guid.NewGuid().ToString();
                    var file = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType, FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
                    var bw = new BinaryWriter(file);
                    bw.Write(data);
                    bw.Close();
                    file.Close();
                    if (taskPane.pdfViewer1.FileName != (AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType))
                    {
                        SetPDFViewer(AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType);
                        SessionInfo.UserInfo.Containerpath = AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType;
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (sqlConn != null)
                    sqlConn.Close();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void notifyIcon1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
                ChangeXPFFormStatus();
            else
                contextMenuStrip2.Show();
        }
        /// <summary>
        /// 
        /// </summary>
        delegate void MyDelegate();
        /// <summary>
        /// 
        /// </summary>
        private void BindNotify()
        {
            if (notifyIcon1 == null)
                notifyIcon1 = new System.Windows.Forms.NotifyIcon();
            notifyIcon1.Icon = ExcelAddIn4.Properties.Resources.x_office_spreadsheet;
            notifyIcon1.Text = "Finance Tool";
            notifyIcon1.MouseClick += new MouseEventHandler(notifyIcon1_MouseClick);
            notifyIcon1.ContextMenuStrip = this.contextMenuStrip2;
            notifyIcon1.Visible = true;
            notifyIcon1.ShowBalloonTip(10000, "Tip of FinanceTool", "Post journal is processing!", System.Windows.Forms.ToolTipIcon.Info);
            StripMenuItemSetEnable();
        }
        /// <summary>
        /// 
        /// </summary>
        private void BindNotifyTrans()
        {
            if (notifyIcon1 == null)
                notifyIcon1 = new System.Windows.Forms.NotifyIcon();
            notifyIcon1.Icon = ExcelAddIn4.Properties.Resources.x_office_spreadsheet;
            notifyIcon1.Text = "Finance Tool";
            notifyIcon1.MouseClick += new MouseEventHandler(notifyIcon1_MouseClick);
            notifyIcon1.ContextMenuStrip = this.contextMenuStrip2;
            notifyIcon1.Visible = true;
            notifyIcon1.ShowBalloonTip(10000, "Tip of FinanceTool", "Journal update is processing!", System.Windows.Forms.ToolTipIcon.Info);
            StripMenuItemSetEnable();
        }
        /// <summary>
        /// 
        /// </summary>
        private void BindNotifyCreateTextFile()
        {
            if (notifyIcon1 == null)
                notifyIcon1 = new System.Windows.Forms.NotifyIcon();
            notifyIcon1.Icon = ExcelAddIn4.Properties.Resources.x_office_spreadsheet;
            notifyIcon1.Text = "Finance Tool";
            notifyIcon1.MouseClick += new MouseEventHandler(notifyIcon1_MouseClick);
            notifyIcon1.ContextMenuStrip = this.contextMenuStrip2;
            notifyIcon1.Visible = true;
            notifyIcon1.ShowBalloonTip(10000, "Tip of FinanceTool", "Create XML/Text File is processing!", System.Windows.Forms.ToolTipIcon.Info);
            StripMenuItemSetEnable();
        }
        /// <summary>
        /// 
        /// </summary>
        private void StripMenuItemSetEnable()
        {
            suspendToolStripMenuItem.Enabled = true;
        }
        /// <summary>
        /// 
        /// </summary>
        private void StripMenuItemSetDisable()
        {
            suspendToolStripMenuItem.Enabled = false;
        }
        /// <summary>
        /// Journal update process event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnATC_TransUpd_Click(object sender, RibbonControlEventArgs e)
        {
            #region 2. Show notify
            MyDelegate myDelegate = new MyDelegate(BindNotifyTrans);
            myDelegate.Invoke();
            #endregion
            #region 3. Initialize Data Form
            if (tupf != null)
                tupf.Dispose();
            DateTime starttime = DateTime.Now;
            tupf = new TransUpdPostFrm();
            tupf.panel3.Visible = true;
            tupf.panel1.Visible = false;
            tupf.panel2.Visible = false;
            tupf.panel3.Dock = DockStyle.Fill;
            tupf.ControlBox = false;
            #endregion
            OutputContainer.isTransUpdFlag = true;
            try
            {
                #region 4. Data save OUTPUT PANE DGV and extraction data
                Globals.ThisAddIn.Application.DisplayAlerts = false;
                wsRrigin = (Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                var lastColumn = wsRrigin.Cells.Find("*", wsRrigin.Cells[1, 1], Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, Microsoft.Office.Interop.Excel.XlLookAt.xlPart, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, Type.Missing, Type.Missing);
                LastColumnName = Finance_Tools.RemoveNumber(lastColumn.Address).Replace("$", "");
                Finance_Tools.MaxColumnCount = lastColumn.Column;
                var lastrow = wsRrigin.Cells.Find("*", wsRrigin.Cells[1, 1], Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, Microsoft.Office.Interop.Excel.XlLookAt.xlPart, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, Type.Missing, Type.Missing);
                LastRowNumber = Finance_Tools.RemoveNotNumber(lastrow.Address);
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                if (sender != null)
                    tupf.Show();
                TransUpdPostFrm.richTextBox1.Text += "Error List :\r\n";
                outputPane.SaveTransUpd(null, wsRrigin);
                outputPane.SetSession();
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                Globals.ThisAddIn.Application.DisplayAlerts = true;
                if (TransUpdPostFrm.richTextBox1.Text.Length > 21)
                {
                    LogHelper.WriteLog(typeof(Ribbon2), TransUpdPostFrm.richTextBox1.Text + " - Journal update processing error , Template:" + SessionInfo.UserInfo.FileName);
                }
                #endregion
            }
            catch (Exception ex)
            {
                if (ex.ToString().Contains("Clipboard"))
                {
                    MessageBox.Show("Clipboard not ready, please try again.", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    if (Ribbon2.tupf.Visible == true)
                        MessageBox.Show(ex.ToString(), "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else
                        SessionInfo.UserInfo.GlobalError += "Process:Journal Update(" + SessionInfo.UserInfo.CurrentRef + ") - Fail : " + ex.Message;
                }
                LogHelper.WriteLog(typeof(Ribbon2), ex.Message + "Journal update error");
                tupf.Dispose();
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                Clipboard.SetText("\r\n");
            }
            finally
            {
                #region 5. Data Form 's data binding
                tupf.Focus();
                tupf.bddata();
                tupf.panel3.Visible = false;
                tupf.panel1.Visible = true;
                tupf.panel2.Visible = true;
                tupf.ControlBox = true;
                StripMenuItemSetDisable();
                #endregion
                #region 6 . Remove Temporary data and show cost time
                DateTime stoptime = DateTime.Now;
                string costtime = Finance_Tools.DateDiff(starttime, stoptime);
                Globals.ThisAddIn.Application.StatusBar = "latest journal update process costs " + costtime;
                GC.Collect();
                #endregion
                OutputContainer.isTransUpdFlag = false;
                #region 7 . if 'show journal before posting' uncheck , post to SunSystem and Save history
                try
                {
                    if (sender == null)
                        tupf.btnPost_Click(null, null);
                    else
                        tupf.btnPost_Click(sender, null);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                #endregion
            }
        }
        /// <summary>
        /// 'Save' process event (without post to sun)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnATC_Posting_Click(object sender, RibbonControlEventArgs e)
        {
            if (ft.IsGUID(Path.GetFileNameWithoutExtension(SessionInfo.UserInfo.CachePath)) && !string.IsNullOrEmpty(ft.ProcessJournalNumber()))//(SessionInfo.UserInfo.CachePath != SessionInfo.UserInfo.FilePath)
            {
                MessageBox.Show("Can't be changed! This document has been Posted! ", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LogHelper.WriteLog(typeof(Ribbon2), "Can't be changed! This document has been Posted! ");
                return;
            }
            #region 2. Initialize Data Form
            xpf = new XMLPostFrm();
            #endregion
            try
            {
                #region 3. Save history
                SessionInfo.UserInfo.SunJournalNumber = "";
                outputPane.SetSession();
                xpf.SaveHistory();
                #endregion
            }
            catch (Exception ex)
            {
                SessionInfo.UserInfo.GlobalError += "Process:Save(" + SessionInfo.UserInfo.CurrentSaveRef + ") - Fail!" + ex.Message;
                LogHelper.WriteLog(typeof(Ribbon2), ex.Message + "Save error");
                throw ex;
            }
        }
        /// <summary>
        /// 'Post Data' process event 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnATC_Post_Click(object sender, RibbonControlEventArgs e)
        {
            if ((ft.IsGUID(Path.GetFileNameWithoutExtension(SessionInfo.UserInfo.CachePath))) && !string.IsNullOrEmpty(ft.ProcessJournalNumber()))
            {
                MessageBox.Show("Can't be changed! This document has been Posted! ", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LogHelper.WriteLog(typeof(Ribbon2), "Can't be changed! This document has been Posted! ");
                return;
            }
            #region 2. Show notify
            MyDelegate myDelegate = new MyDelegate(BindNotify);
            myDelegate.Invoke();
            #endregion
            #region 3. Initialize Data Form
            DateTime starttime = DateTime.Now;
            if (xpf != null)
                xpf.Dispose();
            xpf = new XMLPostFrm();
            xpf.panel3.Visible = true;
            xpf.panel1.Visible = false;
            xpf.panel2.Visible = false;
            xpf.panel3.Dock = DockStyle.Fill;
            xpf.ControlBox = false;
            #endregion
            try
            {
                #region 4. Data save OUTPUT PANE DGV and extraction data
                Globals.ThisAddIn.Application.DisplayAlerts = false;
                wsRrigin = (Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                var lastColumn = wsRrigin.Cells.Find("*", wsRrigin.Cells[1, 1], Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, Microsoft.Office.Interop.Excel.XlLookAt.xlPart, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, Type.Missing, Type.Missing);
                LastColumnName = Finance_Tools.RemoveNumber(lastColumn.Address).Replace("$", "");
                Finance_Tools.MaxColumnCount = lastColumn.Column;
                var lastrow = wsRrigin.Cells.Find("*", wsRrigin.Cells[1, 1], Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, Microsoft.Office.Interop.Excel.XlLookAt.xlPart, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, Type.Missing, Type.Missing);
                LastRowNumber = Finance_Tools.RemoveNotNumber(lastrow.Address);
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                if (sender != null)
                    xpf.Show();
                XMLPostFrm.richTextBox1.Text += "Error List :\r\n";
                outputPane.Save(null, wsRrigin);
                outputPane.SaveCons(null, wsRrigin);
                outputPane.SetSession();
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                Globals.ThisAddIn.Application.DisplayAlerts = true;
                if (XMLPostFrm.richTextBox1.Text.Length > 21)
                {
                    LogHelper.WriteLog(typeof(Ribbon2), XMLPostFrm.richTextBox1.Text + " - Post Journal Processing error , Template:" + SessionInfo.UserInfo.FileName);
                }
                #endregion
            }
            catch (Exception ex)
            {
                if (ex.ToString().Contains("Clipboard"))
                {
                    MessageBox.Show("Clipboard not ready, please try again.", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    if (Ribbon2.xpf.Visible == true)
                        MessageBox.Show(ex.ToString(), "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else
                        SessionInfo.UserInfo.GlobalError += "Process: Journal Post - Fail : " + ex.Message;
                }
                LogHelper.WriteLog(typeof(Ribbon2), ex.Message + "Post error");
                xpf.Dispose();
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                Clipboard.SetText("\r\n");
            }
            finally
            {
                #region 5. Data Form 's data binding
                xpf.Focus();
                xpf.bddata();
                xpf.panel3.Visible = false;
                xpf.panel1.Visible = true;
                xpf.panel2.Visible = true;
                xpf.ControlBox = true;
                StripMenuItemSetDisable();
                #endregion
                #region 6 . Remove Temporary data and show cost time
                DateTime stoptime = DateTime.Now;
                string costtime = Finance_Tools.DateDiff(starttime, stoptime);
                Globals.ThisAddIn.Application.StatusBar = "latest posting process costs " + costtime;
                GC.Collect();
                #endregion
                #region 7 . if 'show journal before posting' uncheck , post to SunSystem and Save history
                try
                {
                    if (sender == null)
                        xpf.btnPost_Click(null, null);
                    else
                        xpf.btnPost_Click(sender, null);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                #endregion
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCreateXMLorTextFile_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                CreateXMLTextProfile lf = new CreateXMLTextProfile();
                lf.ShowDialog();
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPreferences_Preferences_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Preferences lf = new Preferences();
                lf.ShowDialog();
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnUpgrade_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Upgrade u = new Upgrade();
                u.ShowDialog();
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSecurity_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Security s = new Security();
                s.ShowDialog();
            }
            catch { }
        }
        /// <summary>
        /// 'Output' CheckBox Click event (Hide or show OutPut Container) 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnATVP_Output_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Microsoft.Office.Tools.CustomTaskPane filtered2 = Globals.ThisAddIn.CustomTaskPanes.First(s => s.Title.Contains("Output settings"));
                if (this.btnATVP_Output.Checked == true)
                    filtered2.Visible = true;
                else
                    filtered2.Visible = false;
            }
            catch { }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnATVP_Help_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Microsoft.Office.Tools.CustomTaskPane filtered3 = Globals.ThisAddIn.CustomTaskPanes.First(s => s.Title.Contains("Help Container"));
                if (this.btnATVP_Help.Checked == true)
                    filtered3.Visible = true;
                else
                    filtered3.Visible = false;
            }
            catch { }
        }
        /// <summary>
        /// 'new action' Button Click event 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnNewButton_Click(object sender, RibbonControlEventArgs e)
        {
            SessionInfo.UserInfo.GlobalError = "";
            RibbonButton rb = (RibbonButton)sender;
            string rbid = rb.Tag.ToString();
            string[] sArray = Regex.Split(rbid, ",");
            string templateid = sArray[0];
            string buttonName = sArray[1];
            DataTable processMacro = ft.GetProcessMacroFromDB(templateid, buttonName);
            bool stop = (bool)processMacro.Rows[0]["StopOnError"];
            bool showmsg = (bool)processMacro.Rows[0]["ShowMsg"];
            for (int i = 0; i < processMacro.Rows.Count; i++)
            {
                try
                {
                    int type = (int)processMacro.Rows[i]["Type"];
                    string reference = processMacro.Rows[i]["Reference"].ToString().Replace(" ", "");
                    string processID = processMacro.Rows[i]["ProcessID"].ToString();
                    string macro = string.Empty;
                    try
                    {
                        macro = processMacro.Rows[i]["MacroName"].ToString();
                    }
                    catch { }
                    string[] sArray2 = Regex.Split(processID, ",");
                    try
                    {
                        if (sArray2[1].Contains(".") && sArray2[1].Substring(sArray2[1].IndexOf(".") + 1, sArray2[1].Length - sArray2[1].IndexOf(".") - 1) != "")//this means it's xml otherwise it's text
                        {
                            SessionInfo.UserInfo.ComName = sArray2[1].Substring(0, sArray2[1].IndexOf("."));
                            SessionInfo.UserInfo.MethodName = sArray2[1].Substring(sArray2[1].IndexOf(".") + 1, sArray2[1].Length - sArray2[1].IndexOf(".") - 1);
                            outputPane.cbXMLOrText.SelectedIndex = 0;
                            outputPane.cbXMLOrText_SelectedIndexChanged(null, null);
                            outputPane.cbItems.SelectedItem = SessionInfo.UserInfo.ComName + "," + SessionInfo.UserInfo.MethodName;
                        }
                        else
                        {
                            SessionInfo.UserInfo.Textfilename = sArray2[1].Substring(0, sArray2[1].IndexOf("."));
                            outputPane.cbXMLOrText.SelectedIndex = 1;
                            outputPane.cbXMLOrText_SelectedIndexChanged(null, null);
                            outputPane.cbItems.SelectedItem = SessionInfo.UserInfo.Textfilename;
                        }
                    }
                    catch { }
                    if (processMacro.Rows.Count == 1)
                    {
                        #region if processMacroId.Count==1 means the button has only process exist, so the second parameter is real sender to make Data Form show.
                        ExecProcess(type, macro, sender, reference);
                        #endregion
                    }
                    else if (processMacro.Rows.Count > 1)
                    {
                        #region if processMacroId.Count>1 means the button have process and Macro exist, so the second parameter is null to make Data Form not show.
                        ExecProcess(type, macro, null, reference);
                        #endregion
                    }
                }
                catch
                {
                    if (stop)
                    {
                        GenerateError(i, processMacro);
                        break;
                    }
                }
                finally
                {
                    SessionInfo.UserInfo.ComName = "";
                    SessionInfo.UserInfo.MethodName = "";
                    SessionInfo.UserInfo.Textfilename = "";
                    SessionInfo.UserInfo.CurrentRef = "";
                    SessionInfo.UserInfo.CurrentSaveRef = "";
                }
            }
            if (showmsg && !string.IsNullOrEmpty(SessionInfo.UserInfo.GlobalError.Replace("\r\n", "")) && (SessionInfo.UserInfo.GlobalError.Replace("\r\n", "") != "Stop on Error,Action ended.The following processes did not run:"))
            {
                PostErrorFrm pef = new PostErrorFrm(SessionInfo.UserInfo.GlobalError);
                pef.ShowDialog();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="j"></param>
        /// <param name="dt"></param>
        private void GenerateError(int j, DataTable dt)
        {
            SessionInfo.UserInfo.GlobalError += "\r\n\r\n";
            SessionInfo.UserInfo.GlobalError += "Stop on Error,Action ended.The following processes did not run:\r\n";
            for (int i = j + 1; i < dt.Rows.Count; i++)
            {
                int type = (int)dt.Rows[i]["Type"];
                string reference = dt.Rows[i]["Reference"].ToString().Replace(" ", "");
                string processID = dt.Rows[i]["ProcessID"].ToString();
                string macro = string.Empty;
                try
                {
                    macro = dt.Rows[i]["MacroName"].ToString();
                }
                catch { }
                if (type == 1)
                {
                    SessionInfo.UserInfo.GlobalError += "Process:Journal Post(" + reference + ")";
                    SessionInfo.UserInfo.GlobalError += "\r\n";
                }
                else if (type == 2)
                {
                    SessionInfo.UserInfo.GlobalError += "Process:Journal Update(" + reference + ")";
                    SessionInfo.UserInfo.GlobalError += "\r\n";
                }
                else if (type == 4)
                {
                    SessionInfo.UserInfo.GlobalError += "Process:Save(" + reference + ")";
                    SessionInfo.UserInfo.GlobalError += "\r\n";
                }
                else if (type == 7)
                {
                    SessionInfo.UserInfo.GlobalError += "Process:Save PDF(" + reference + ")";
                    SessionInfo.UserInfo.GlobalError += "\r\n";
                }
                else if (type == 5)
                {
                    SessionInfo.UserInfo.GlobalError += "Macro:" + macro;
                    SessionInfo.UserInfo.GlobalError += "\r\n";
                }
                else if (type == 6)
                {
                    SessionInfo.UserInfo.GlobalError += "Process:Reopen template";
                    SessionInfo.UserInfo.GlobalError += "\r\n";
                }
                else if (type == 3)
                {
                    try
                    {
                        string[] sArray2 = Regex.Split(processID, ",");
                        if (sArray2[1].Contains(".") && sArray2[1].Substring(sArray2[1].IndexOf(".") + 1, sArray2[1].Length - sArray2[1].IndexOf(".") - 1) != "")//this means it's xml otherwise it's text
                        {
                            SessionInfo.UserInfo.GlobalError += "Process:" + sArray2[1].Substring(0, sArray2[1].IndexOf(".")) + " ";
                            SessionInfo.UserInfo.GlobalError += sArray2[1].Substring(sArray2[1].IndexOf(".") + 1, sArray2[1].Length - sArray2[1].IndexOf(".") - 1);
                            SessionInfo.UserInfo.GlobalError += "(" + reference + ")";
                            SessionInfo.UserInfo.GlobalError += "\r\n";
                        }
                        else
                        {
                            SessionInfo.UserInfo.GlobalError += "Process:" + sArray2[1].Substring(0, sArray2[1].IndexOf("."));
                            SessionInfo.UserInfo.GlobalError += "(" + reference + ")";
                            SessionInfo.UserInfo.GlobalError += "\r\n";
                        }
                    }
                    catch { }
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="type"></param>
        /// <param name="macro"></param>
        /// <param name="sender"></param>
        /// <param name="reference"></param>
        private void ExecProcess(int type, string macro, object sender, string reference)
        {
            SessionInfo.UserInfo.CurrentRef = reference;
            try
            {
                if (type == 1)
                {
                    btnATC_Post_Click(null, null);
                    SessionInfo.UserInfo.GlobalError += "\r\n";
                }
                else if (type == 2)
                {
                    btnATC_TransUpd_Click(sender, null);
                    SessionInfo.UserInfo.GlobalError += "\r\n";
                }
                else if (type == 4)
                {
                    SessionInfo.UserInfo.CurrentSaveRef = reference;
                    btnATC_Posting_Click(sender, null);
                    SessionInfo.UserInfo.GlobalError += "\r\n";
                }
                else if (type == 7)
                {
                    //SessionInfo.UserInfo.CurrentSaveRef = reference;
                    SavePDF(reference);
                    SessionInfo.UserInfo.GlobalError += "\r\n";
                }
                else if (type == 5)
                {
                    ExecMacro(macro);
                    SessionInfo.UserInfo.GlobalError += "\r\n";
                }
                else if (type == 6)
                {
                    btnATC_ReopenTemplate_Click(sender, null);
                    SessionInfo.UserInfo.GlobalError += "\r\n";
                }
                else if (type == 3)
                {
                    btnCtrl_CreateTextFile_Click(sender, null);
                    SessionInfo.UserInfo.GlobalError += "\r\n";
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="reference"></param>
        private void SavePDF(string reference)
        {
            DataTable dt = ft.GetReportCriteriaByRef(SessionInfo.UserInfo.File_ftid, reference);
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("Save PDF Reference " + reference + " error!"); return;
            }
            string path = ft.GetValueOfAddress(dt.Rows[0]["PDFFolder"].ToString()) + "\\" + ft.GetValueOfAddress(dt.Rows[0]["PDFName"].ToString());
            path = path.Replace("\\\\", "\\") + ".pdf";
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsTemplateTransactions_UpdatePDF", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                cmd.Parameters.Add(new SqlParameter("@PDFData", ft.GetData(path)));
                cmd.Parameters.Add(new SqlParameter("@maxNum", SessionInfo.UserInfo.InvNumber));
                cmd.ExecuteNonQuery();
                SessionInfo.UserInfo.GlobalError += "Process: Save PDF(" + reference + ") - Success !";
            }
            catch (Exception ex)
            {
                SessionInfo.UserInfo.GlobalError += "Process: Save PDF(" + reference + ") - Fail !" + ex.Message;
                throw ex;
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
        /// <param name="elem"></param>
        /// <returns></returns>
        private bool EquaWithName(KeyValuePair<string, string> elem)
        {
            if (elem.Value == idv)
                return true;
            return false;
        }
        /// <summary>
        /// 
        /// </summary>
        private string idv = string.Empty;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnATC_ReopenTemplate_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Globals.ThisAddIn.Application.DisplayAlerts = false;
                var xlapp = Globals.ThisAddIn.Application;
                string mypath = xlapp.ActiveWorkbook.FullName;
                xlapp.ActiveWorkbook.Close();
                if (mypath.Contains("RSDataCache") || ft.IsGUID(Path.GetFileNameWithoutExtension(mypath)))
                {
                    string IDvalue = SessionInfo.UserInfo.Dictionary.dict[mypath];
                    string[] sArray = Regex.Split(IDvalue, ",");
                    IDvalue = sArray[0];
                    idv = IDvalue;
                    Predicate<KeyValuePair<string, string>> pred = EquaWithName;
                    if (TemplateAndPath.Exists(pred))
                    {
                        string p = ft.getFilePath(IDvalue);
                        try
                        {
                            xlapp.Workbooks.Open(p);
                        }
                        catch { xlapp.Workbooks.Open(p.Replace(".xlsm", ".xlsx")); }
                        SessionInfo.UserInfo.GlobalError += "Process: Reopen template - Success !";
                    }
                    else
                    {
                        SessionInfo.UserInfo.GlobalError += "Process: Reopen template - Fail ! Template's access does not exist in your permissions, Please contact the admin. ";//MessageBox.Show("Template's access does not exist in your permissions, Please contact the admin.");
                    }
                }
                else
                {
                    xlapp.Workbooks.Open(mypath);
                    SessionInfo.UserInfo.GlobalError += "Process: Reopen template - Success !";
                }
                Globals.ThisAddIn.Application.DisplayAlerts = true;
            }
            catch (Exception ex)
            {
                SessionInfo.UserInfo.GlobalError += "Process: Reopen template - Fail !" + ex.Message;
                throw ex;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="macro"></param>
        private void ExecMacro(string macro)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Workbook Wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                var xlapp = Globals.ThisAddIn.Application;
                string s = Wb.FullName;
                xlapp.Run("'" + Wb.FullName + "'!" + macro);
                SessionInfo.UserInfo.GlobalError += "Macro: " + macro + " - Success!";
            }
            catch (Exception ex)
            {
                SessionInfo.UserInfo.GlobalError += "Macro:  " + macro + " - Fail!" + ex.Message;
                LogHelper.WriteLog(typeof(Ribbon2), ex.Message + "Exec Macro error");
                throw ex;
            }
        }
        /// <summary>
        /// 'Create new action' Button Click event 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAd_NewButton_Click(object sender, RibbonControlEventArgs e)
        {
            frmAd_NewButton frm = new frmAd_NewButton();
            frm.ShowDialog();
        }
        /// <summary>
        /// 'Add Document View' Button Click event (Show frmAd_Doc_View Form) 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAd_DocView_Click(object sender, RibbonControlEventArgs e)
        {
            frmAd_Doc_View frm = new frmAd_Doc_View();
            frm.ShowDialog();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAd_AttachPDF_Click(object sender, RibbonControlEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = false;
            fileDialog.Title = "Please choose a PDF file";
            fileDialog.Filter = "PDF Files|*.pdf|All Files|*.*";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                Microsoft.Office.Interop.Excel.Workbook Wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                string path = Wb.FullName;
                if (path.Contains("RSDataCache") || ft.IsGUID(Path.GetFileNameWithoutExtension(path)))
                {
                    SqlConnection conn = null;
                    try
                    {
                        conn = new
                            SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                        conn.Open();
                        SqlCommand cmd = new SqlCommand("rsTemplateTransactions_UpdatePDF", conn);
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                        cmd.Parameters.Add(new SqlParameter("@PDFData", ft.GetData(fileDialog.FileName)));
                        cmd.Parameters.Add(new SqlParameter("@maxNum", SessionInfo.UserInfo.InvNumber));
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Success!", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    finally
                    {
                        if (conn != null)
                            conn.Close();
                    }
                }
                else
                {
                    MessageBox.Show("Please open a transaction first !", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnJournal_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var xlapp = Globals.ThisAddIn.Application;//take the value from the activecell in excel
                string cellvalue = xlapp.ActiveCell.Value.ToString();
                int i = 0;
                bool isNum = int.TryParse(cellvalue, out i);
                if (isNum)
                {
                    //open the "vwJournal.xlsm" file. This name will not change so hard code in this setting
                    var path = Finance_Tools.RootPath;
                    string filename = path + "\\vwJournal.xlsm";
                    var mybk = xlapp.Workbooks.Open(filename);
                    var sheet = (Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                    //find the cell named myRef
                    //the cell "myRef" will always be called this in the vwJournal file
                    var range = sheet.get_Range("myRef", Type.Missing);
                    range.Value2 = cellvalue;
                    //run the macro "runProc", which will not change
                    xlapp.Run("'" + filename + "'!RSystems.runProc");
                    //return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LogHelper.WriteLog(typeof(Ribbon2), ex.Message + "Journal error");
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPDF_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var xlapp = Globals.ThisAddIn.Application;
                string cellvalue = xlapp.ActiveCell.Value.ToString();
                if (!ft.ReadPDFinDB(cellvalue: cellvalue))
                    if (!ft.ReadPDFinViewDocumentSetting(cellvalue))
                        if (SessionInfo.UserInfo.Dictionary.dict.Count != 0 && SessionInfo.UserInfo.Dictionary.dict.ContainsKey(SessionInfo.UserInfo.CachePath))
                            if (MessageBox.Show("The PDF you input does not exist! Open current transaction's PDF?", "Message - RSystems FinanceTools", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                ft.ReadPDFOfCurrentTransaction();
                            else
                                MessageBox.Show("The PDF you input does not exist!", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                ft.ReadPDFOfCurrentTransaction();
            }
        }
        /// <summary>
        /// 'View Document' Button Click event (open files and run the macro)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCtrl_VD_Click(object sender, RibbonControlEventArgs e)
        {
            var xlapp = Globals.ThisAddIn.Application;
            string cellvalue = xlapp.ActiveCell.Value.ToString();
            SessionInfo.UserInfo.File_ftid = ft.ProcessFilePath(cellvalue);
            SessionInfo.UserInfo.InvNumber = ft.ProcessInvNumber(cellvalue);
            SessionInfo.UserInfo.FilePath = ft.getFilePath(SessionInfo.UserInfo.File_ftid);
            try
            {
                var data = ft.ProcessData(cellvalue);
                if (data.Length == 0) throw new Exception();
                var fileType = ft.ProcessFileType(cellvalue);
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
                    xlapp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType);//xlapp.Run("'" + AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType + "'!RSystems.runProc");
                    return;
                }
            }
            catch
            {
                Array arr = ft.vdPrefix();
                foreach (string prefix in arr)
                {
                    if (prefix.Trim().ToLower() == cellvalue.Substring(0, prefix.Trim().Length).Trim().ToLower())
                    {
                        string filename = ft.vdFilepath(prefix).Trim();
                        bool file = ft.vdUseFile(prefix);
                        if (file)
                        {
                            filename = filename + "\\" + cellvalue;
                        }
                        else
                        {
                            filename = filename + "\\" + ft.vdFilename(prefix).Trim();
                        }
                        string ftype = ft.vdFiletype(prefix).Trim();
                        if (ftype == "pdf")
                        {
                            Process.Start(filename.Trim() + "." + ftype);
                            return;
                        }
                        else if (ftype == "html")
                        {
                            Process.Start(filename.Trim().Replace("\\", ""));
                            return;
                        }
                        else
                        {
                            try
                            {
                                filename = filename.Trim() + "." + ftype;
                                xlapp.Workbooks.Open(filename);
                                string macro = ft.vdMacro01(prefix);
                                if (!string.IsNullOrEmpty(macro.Replace(".", "")))
                                {
                                    xlapp.Run("'" + filename + "'!" + macro);
                                    return;
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                                continue;
                            }
                        }
                    }
                }
            }
        }
        /// <summary>
        /// 'Transaction search' Button Click event (open files and run the macro)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCtrl_VR_Click(object sender, RibbonControlEventArgs e)
        {
            ViewReport vr = new ViewReport();
            vr.ShowDialog();
        }
        /// <summary>
        /// 'View XML' Button Click event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCtrl_VX_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Workbook Wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            string path = Wb.FullName;
            if (path.Contains("RSDataCache") || ft.IsGUID(Path.GetFileNameWithoutExtension(path)))
            {
                var data = ft.ProcessXML(SessionInfo.UserInfo.File_ftid, SessionInfo.UserInfo.InvNumber);
                if (data.Length == 0)
                {
                    MessageBox.Show("No Data!", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                var fileType = ".txt";
                string tmp = Guid.NewGuid().ToString();
                var file = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType, FileMode.OpenOrCreate, FileAccess.Write, FileShare.ReadWrite);
                StreamWriter sw = new StreamWriter(file);
                sw.Write(data);
                sw.Close();
                Process.Start(AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType);
                return;
            }
            else
            {
                MessageBox.Show("Please open a transaction first !", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        /// <summary>
        /// 'Create text file' Button Click event 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCtrl_CreateTextFile_Click(object sender, RibbonControlEventArgs e)
        {
            #region 1. Show notify
            MyDelegate myDelegate = new MyDelegate(BindNotifyCreateTextFile);
            myDelegate.Invoke();
            #endregion
            #region 2. Initialize Data Form
            DateTime starttime = DateTime.Now;
            if (ctff != null)
                ctff.Dispose();
            ctff = new CreateTextFileForm();
            #endregion
            try
            {
                Globals.ThisAddIn.Application.DisplayAlerts = false;
                wsRrigin = (Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                var lastColumn = wsRrigin.Cells.Find("*", wsRrigin.Cells[1, 1], Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, Microsoft.Office.Interop.Excel.XlLookAt.xlPart, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, Type.Missing, Type.Missing);
                LastColumnName = Finance_Tools.RemoveNumber(lastColumn.Address).Replace("$", "");
                Finance_Tools.MaxColumnCount = lastColumn.Column;
                var lastrow = wsRrigin.Cells.Find("*", wsRrigin.Cells[1, 1], Microsoft.Office.Interop.Excel.XlFindLookIn.xlFormulas, Microsoft.Office.Interop.Excel.XlLookAt.xlPart, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious, false, Type.Missing, Type.Missing);
                LastRowNumber = Finance_Tools.RemoveNotNumber(lastrow.Address);
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                CreateTextFileForm.richTextBox1.Text += "Error List :\r\n";
                if (string.IsNullOrEmpty(SessionInfo.UserInfo.Textfilename))
                {
                    if (sender != null)
                        Ribbon2.ctff.Show();
                    outputPane.SaveXML(null, Ribbon2.wsRrigin);
                }
                else
                {
                    outputPane.SaveCTF(null, Ribbon2.wsRrigin);
                }
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                Globals.ThisAddIn.Application.DisplayAlerts = true;
                if (CreateTextFileForm.richTextBox1.Text.Length > 21)
                {
                    LogHelper.WriteLog(typeof(Ribbon2), CreateTextFileForm.richTextBox1.Text + " - Create Text File Processing error , Template:" + SessionInfo.UserInfo.FileName);
                }
                if (!string.IsNullOrEmpty(SessionInfo.UserInfo.Textfilename))
                {
                    if (outputPane.finallistCTF.Count != 0)
                    {
                        string fileName = "";
                        string filepath = "";
                        bool includeHeaderRow = false;
                        for (int i = 0; i < outputPane.dgvCreateTextFile.Rows.Count; i++)
                        {
                            if (SessionInfo.UserInfo.CurrentRef == (outputPane.dgvCreateTextFile.Rows[i].Cells[0].Value == null ? "" : outputPane.dgvCreateTextFile.Rows[i].Cells[0].Value.ToString().Replace(" ", "")))
                            {
                                filepath = outputPane.dgvCreateTextFile.Rows[i].Cells[3].Value == null ? "" : outputPane.dgvCreateTextFile.Rows[i].Cells[3].Value.ToString().Replace(" ", "");
                                fileName = outputPane.dgvCreateTextFile.Rows[i].Cells[4].Value == null ? "" : outputPane.dgvCreateTextFile.Rows[i].Cells[4].Value.ToString().Replace(" ", "");
                                includeHeaderRow = outputPane.dgvCreateTextFile.Rows[i].Cells[5].Value.ToString().Replace(" ", "") == "True" ? true : false;
                            }
                        }
                        outputPane.GenTextFile(null, fileName, filepath, includeHeaderRow);
                    }
                    else
                        SessionInfo.UserInfo.GlobalError += "Process:" + SessionInfo.UserInfo.Textfilename + "(" + SessionInfo.UserInfo.CurrentRef + ") - Warn : No data!";
                    return;
                }
            }
            catch (Exception ex)
            {
                if (ex.ToString().Contains("Clipboard"))
                {
                    MessageBox.Show("Clipboard not ready, please try again.", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    if (Ribbon2.ctff.Visible == true)
                        MessageBox.Show(ex.ToString(), "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else
                        SessionInfo.UserInfo.GlobalError += "Process:" + (SessionInfo.UserInfo.ComName == "" ? SessionInfo.UserInfo.Textfilename : SessionInfo.UserInfo.ComName + " " + SessionInfo.UserInfo.MethodName) + "(" + SessionInfo.UserInfo.CurrentRef + ") - Fail : " + ex.Message;
                }
                LogHelper.WriteLog(typeof(Ribbon2), ex.Message + "Create Text File error");
                ctff.Dispose();
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                Clipboard.SetText("\r\n");
            }
            finally
            {
                StripMenuItemSetDisable();
                Ribbon2.ctff.Focus();
                Ribbon2.ctff.bddata();
                Ribbon2.ctff.panel3.Visible = false;
                Ribbon2.ctff.panel1.Visible = true;
                Ribbon2.ctff.panel2.Visible = true;
                Ribbon2.ctff.ControlBox = true;
                DateTime stoptime = DateTime.Now;
                string costtime = Finance_Tools.DateDiff(starttime, stoptime);
                Globals.ThisAddIn.Application.StatusBar = "latest create xml/text file process costs " + costtime;
                GC.Collect();
                if (string.IsNullOrEmpty(SessionInfo.UserInfo.Textfilename))
                    try
                    {
                        if (sender == null)
                            ctff.btnPost_Click(null, null);
                        else
                            ctff.btnPost_Click(sender, null);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        private bool Config()
        {
            if (ConfigurationManager.AppSettings["isSqlInitialize"] == "false")
            {
                //FileStream fs = new FileStream("d:\\a.txt", FileMode.Open); 
                //StreamReader m_streamReader = new StreamReader(fs); 
                //m_streamReader.BaseStream.Seek(0, SeekOrigin.Begin);        

                //string strLine = m_streamReader.ReadLine();            
                //string[] split = strLine.Split(new char[] { '=' });            
                //string a = split[0];           
                //string n = a.Substring(0);                       
                //while(strLine!=null)            
                //{                
                //    if (a == m)                
                //    {                    
                //        arry += strLine + "\n";                    
                //        strLine = m_streamReader.ReadLine();                                                         
                //    }                             
                //}            
                //Console.Write(arry);                      
                //Console.ReadLine();

                FileStream aFile = new FileStream("C:\\ProgramData\\RSDataV2\\RSDataConfig\\Server.txt", FileMode.Open);
                StreamReader sr = new StreamReader(aFile);
                string strLine = sr.ReadLine();
                string ServerName = DEncrypt.Decrypt(strLine);
                string UserName = DEncrypt.Decrypt(sr.ReadLine());
                string Password = DEncrypt.Decrypt(sr.ReadLine());
                sr.Close();

                ////add windows authentication to Sql
                //string UserSql = @"IF  EXISTS (SELECT * FROM sys.server_principals WHERE name = N'" + WindowsIdentity.GetCurrent().Name + "') \r\n DROP LOGIN [" + WindowsIdentity.GetCurrent().Name + "]  \r\n   CREATE LOGIN [" + WindowsIdentity.GetCurrent().Name + "] FROM WINDOWS WITH DEFAULT_DATABASE=[RSData]  \r\n EXEC SP_ADDROLEMEMBER 'db_owner',[" + WindowsIdentity.GetCurrent().Name + "]  \r\n ";

                //string connString2 = string.Format("Data Source={0};Initial Catalog=RSData;User ID={1};Password={2}", ServerName, UserName, Password);

                //SqlConnection sqlConn2 = new SqlConnection(connString2);
                //SqlCommand myCommand = new SqlCommand(UserSql, sqlConn2);
                //try
                //{
                //    sqlConn2.Open();
                //    myCommand.ExecuteNonQuery();
                //}
                //catch (System.Exception ex)
                //{
                //    MessageBox.Show(ex.ToString(), "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //}
                //finally
                //{
                //    if (sqlConn2.State == ConnectionState.Open)
                //    {
                //        sqlConn2.Close();
                //    }
                //}

                //add windows authentication to DB
                string connString = string.Format("Data Source={0};Initial Catalog=RSDataV2;Integrated Security=True;MultipleActiveResultSets=True", ServerName);
                string str = ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.Replace("(local)", ServerName);
                string str2 = ConfigurationManager.ConnectionStrings["RSFinanceToolsEntities"].ConnectionString.Replace("127.0.0.1", ServerName);
                Finance_Tools.ConnectionStringsSave("conRsTool", str);
                Finance_Tools.ConnectionStringsSave("RSFinanceToolsEntities", str2);
                var wUserID = string.Empty;
                try
                {
                    string userName = WindowsIdentity.GetCurrent().Name;
                    wUserID = (from FT_user in db.rsUsers
                               where FT_user.WindowsUserID == userName
                               select FT_user.ft_id).First().ToString();
                }
                catch
                {
                    SqlConnection sqlConn = new SqlConnection(connString);//FT_Users_DelByWindowsUserID
                    SqlDataReader rdr = null;
                    try
                    {
                        sqlConn.Open();
                        SqlCommand cmd = new SqlCommand("rsUsers_Ins", sqlConn);
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add(new SqlParameter("@WindowsUserID", WindowsIdentity.GetCurrent().Name));
                        rdr = cmd.ExecuteReader();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("The test database connection failed! Please check the server name is correct!", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LogHelper.WriteLog(typeof(Ribbon2), ex.Message + "Sql Initialize Config error");
                        return false;
                    }
                    finally
                    {
                        if (sqlConn != null)
                        {
                            sqlConn.Close();
                        }
                        if (rdr != null)
                        {
                            rdr.Close();
                        }
                    }
                }
                Finance_Tools.AppSettingSave("isSqlInitialize", "true");
                Finance_Tools.ConfigFileStructureFile();
                return true;
            }
            else
            {
                return true;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Wb"></param>
        private void Application_WorkbookActivate(Microsoft.Office.Interop.Excel.Workbook Wb)
        {
            string path = Wb.FullName;
            //if (path != SessionInfo.UserInfo.FilePath)
            SyncPanes(Path.GetFileNameWithoutExtension(path), path);

            if (string.IsNullOrEmpty(SessionInfo.UserInfo.FilePath))
            {
                SessionInfo.UserInfo.FilePath = path;
                SessionInfo.UserInfo.CachePath = path;
                SessionInfo.UserInfo.FileName = Path.GetFileNameWithoutExtension(path);
                if (!ft.IsGUID(Path.GetFileNameWithoutExtension(path)))
                    SessionInfo.UserInfo.File_ftid = ft.GetTemplateIDByName(Path.GetFileNameWithoutExtension(path)).ToString();
            }
            else if (!path.Contains("RSDataCache") && !ft.IsGUID(Path.GetFileNameWithoutExtension(path)))
            {
                SessionInfo.UserInfo.FilePath = path;
                SessionInfo.UserInfo.CachePath = path;
                SessionInfo.UserInfo.FileName = Path.GetFileNameWithoutExtension(path);
                if (!ft.IsGUID(Path.GetFileNameWithoutExtension(path)))
                    SessionInfo.UserInfo.File_ftid = ft.GetTemplateIDByName(Path.GetFileNameWithoutExtension(path)).ToString();
                ThisWorkbook_SheetSelectionChange(null, Globals.ThisAddIn.Application.ActiveCell);
            }
            else
            {
                SessionInfo.UserInfo.CachePath = path;
                if (!SessionInfo.UserInfo.Dictionary.dict.ContainsKey(SessionInfo.UserInfo.CachePath))
                    SessionInfo.UserInfo.Dictionary.dict.Add(SessionInfo.UserInfo.CachePath, SessionInfo.UserInfo.File_ftid + "," + SessionInfo.UserInfo.InvNumber);
                else
                {
                    SessionInfo.UserInfo.File_ftid = SessionInfo.UserInfo.Dictionary.dict[SessionInfo.UserInfo.CachePath];
                    string[] sArray = Regex.Split(SessionInfo.UserInfo.File_ftid, ",");
                    SessionInfo.UserInfo.File_ftid = sArray[0];
                    SessionInfo.UserInfo.InvNumber = int.Parse(sArray[1]);
                    SessionInfo.UserInfo.FilePath = ft.getFilePath(SessionInfo.UserInfo.File_ftid);
                }
                SessionInfo.UserInfo.FileName = Path.GetFileNameWithoutExtension(SessionInfo.UserInfo.FilePath);
                readfromDB(invnumber: SessionInfo.UserInfo.InvNumber, id: SessionInfo.UserInfo.File_ftid);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Sh"></param>
        /// <param name="Target"></param>
        private void ThisWorkbook_SheetSelectionChange(object Sh, Microsoft.Office.Interop.Excel.Range Target)
        {
            //int row = Target.Row;
            //int column = Target.Column;
            //string ColumnLetter = Target.Address.Replace("$", "").Replace(row.ToString(), "");
            //var xlapp = Globals.ThisAddIn.Application;
            //string cellvalue = xlapp.ActiveCell.Value.ToString();
            //var columnInDB = ft.PDFColumn(SessionInfo.UserInfo.File_ftid);
            //var relatedPDFPath = ft.GetPDFViaTemplatePath(SessionInfo.UserInfo.File_ftid);
            //bool? isFromDB = ft.GetIsFromDBViaTemplatePath(SessionInfo.UserInfo.File_ftid);
            //if ((column.ToString().Equals(columnInDB)) || (ColumnLetter.ToUpper().Equals(columnInDB.ToUpper())))
            //{
            //    if ((bool)isFromDB)
            //    {
            //        if (!readfromDB(cellvalue: cellvalue))
            //        {
            //            #region Initialize PDFViewer
            //            string folder = Path.GetFullPath(relatedPDFPath);
            //            string pdfpath = folder + "\\" + cellvalue + ".pdf";
            //            if (taskPane.pdfViewer1.FileName != pdfpath)
            //            {
            //                SetPDFViewer(pdfpath);
            //                SessionInfo.UserInfo.Containerpath = pdfpath;
            //            }
            //            #endregion
            //        }
            //    }
            //    else
            //    {
            //        #region Initialize PDFViewer
            //        string folder = Path.GetFullPath(relatedPDFPath);
            //        string pdfpath = folder + "\\" + cellvalue + ".pdf";
            //        if (taskPane.pdfViewer1.FileName != pdfpath)
            //        {
            //            SetPDFViewer(pdfpath);
            //            SessionInfo.UserInfo.Containerpath = pdfpath;
            //        }
            //        #endregion
            //    }
            //}
            if (this.btnATVP_View.Checked == false) return;
            int row = Target.Row;
            int column = Target.Column;
            string ColumnLetter = Target.Address.Replace("$", "").Replace(row.ToString(), "");
            var xlapp = Globals.ThisAddIn.Application;
            string cellvalue = xlapp.ActiveCell.Value.ToString();
            //var columnInDB = ft.PDFColumn(SessionInfo.UserInfo.File_ftid);
            var relatedPDFPath = ft.GetPDFViaTemplatePath(SessionInfo.UserInfo.File_ftid);
            //bool? isFromDB = ft.GetIsFromDBViaTemplatePath(SessionInfo.UserInfo.File_ftid);
            bool IsWebBrowser = false; bool IsDataBaseQuery = false; string ConnectionString = string.Empty; string UserID = string.Empty; string Password = string.Empty; string SQLString = string.Empty;
            GetPDFDetails(ref IsWebBrowser, ref IsDataBaseQuery, ref ConnectionString, ref UserID, ref Password, ref SQLString);
            //if ((column.ToString().Equals(columnInDB)) || (ColumnLetter.ToUpper().Equals(columnInDB.ToUpper())))
            //{
            if (!readfromDB(cellvalue: cellvalue))
            {
                if (IsDataBaseQuery)
                {
                    OpenDataBaseQuery(IsWebBrowser, ConnectionString, UserID, Password, SQLString, cellvalue);
                }
                else
                {
                    if (IsWebBrowser)
                    {
                        taskPane.Panel1.Visible = false;
                        taskPane.Panel2.Visible = false;
                        taskPane.pdfViewer1.Visible = false;
                        taskPane.panel3.Visible = true;
                        string Url = SessionInfo.UserInfo.Containerpath.Replace("\\\\", "\\");
                        if (Url.EndsWith("//") || Url.EndsWith("/"))
                        {
                            Url += cellvalue + ".html";
                        }
                        else
                        {
                            Url += "/" + cellvalue + ".html";
                        }
                        taskPane.webBrowser1.Navigate(Url);
                    }
                    else
                    {
                        //if ((bool)isFromDB)
                        //{
                        //if (!readfromDB(cellvalue: cellvalue))
                        //{
                        #region Initialize PDFViewer
                        string folder = Path.GetFullPath(relatedPDFPath);
                        string pdfpath = folder + "\\" + cellvalue + ".pdf";
                        if (taskPane.pdfViewer1.FileName != pdfpath)
                        {
                            SetPDFViewer(pdfpath);
                            SessionInfo.UserInfo.Containerpath = pdfpath;
                        }
                        #endregion
                        //}
                        //}
                        //else
                        //{
                        //    #region Initialize PDFViewer
                        //    string folder = Path.GetFullPath(relatedPDFPath);
                        //    string pdfpath = folder + "\\" + cellvalue + ".pdf";
                        //    if (taskPane.pdfViewer1.FileName != pdfpath)
                        //    {
                        //        SetPDFViewer(pdfpath);
                        //        SessionInfo.UserInfo.Containerpath = pdfpath;
                        //    }
                        //    #endregion
                        //}
                    }
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="IsWebBrowser"></param>
        /// <param name="IsDataBaseQuery"></param>
        /// <param name="ConnectionString"></param>
        /// <param name="UserID"></param>
        /// <param name="Password"></param>
        /// <param name="SQLString"></param>
        private void GetPDFDetails(ref bool IsWebBrowser, ref bool IsDataBaseQuery, ref string ConnectionString, ref string UserID, ref string Password, ref string SQLString)
        {
            IsWebBrowser = (bool)(from FT_sett in db.rsTemplateContainers
                                  where FT_sett.TemplateID == SessionInfo.UserInfo.File_ftid
                                  select FT_sett.IsWebBrowser).First();
            IsDataBaseQuery = (bool)(from FT_sett in db.rsTemplateContainers
                                     where FT_sett.TemplateID == SessionInfo.UserInfo.File_ftid
                                     select FT_sett.IsDataBaseQuery).First();
            ConnectionString = (from FT_sett in db.rsTemplateContainers
                                where FT_sett.TemplateID == SessionInfo.UserInfo.File_ftid
                                select FT_sett.ConnectionString).First();
            UserID = (from FT_sett in db.rsTemplateContainers
                      where FT_sett.TemplateID == SessionInfo.UserInfo.File_ftid
                      select FT_sett.UserID).First();
            Password = (from FT_sett in db.rsTemplateContainers
                        where FT_sett.TemplateID == SessionInfo.UserInfo.File_ftid
                        select FT_sett.Password).First();
            SQLString = (from FT_sett in db.rsTemplateContainers
                         where FT_sett.TemplateID == SessionInfo.UserInfo.File_ftid
                         select FT_sett.SQLString).First();
            SessionInfo.UserInfo.Containerpath = (from FT_sett in db.rsTemplateContainers
                                                  where FT_sett.TemplateID == SessionInfo.UserInfo.File_ftid
                                                  select FT_sett.ft_relatefilepath).First();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellvalue"></param>
        /// <param name="id"></param>
        /// <param name="invnumber"></param>
        /// <returns></returns>
        private bool readfromDB(string cellvalue = "", string id = "", int? invnumber = 0)
        {
            try
            {
                var data = ft.ProcessPDFData(cellvalue: cellvalue, ID: id, invnumber: invnumber);
                var fileType = ".pdf";
                string tmp = Guid.NewGuid().ToString();
                var file = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType, FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
                var bw = new BinaryWriter(file);
                bw.Write(data);
                bw.Close();
                file.Close();
                if (taskPane.pdfViewer1.FileName != (AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType))
                {
                    SetPDFViewer(AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType);
                    SessionInfo.UserInfo.Containerpath = AppDomain.CurrentDomain.BaseDirectory + "\\RSDataCache\\" + tmp + fileType;
                }
                return true;
            }
            catch { return false; }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        private void DoLogin(object sender)
        {
            string userName = WindowsIdentity.GetCurrent().Name;
            try
            {
                var wUserID = (from FT_user in db.rsUsers
                               where FT_user.WindowsUserID == userName
                               select FT_user.ft_id).First();
                var SunIP = (from FT_user in db.rsUsers
                             where FT_user.WindowsUserID == userName
                             select FT_user.SUNUserIP).First();
                var SunUID = (from FT_user in db.rsUsers
                              where FT_user.WindowsUserID == userName
                              select FT_user.SUNUserID).First();
                var SunUpass = (from FT_user in db.rsUsers
                                where FT_user.WindowsUserID == userName
                                select FT_user.SUNUserPass).First();
                var tabLabel = (from FT_user in db.rsUsers
                                where FT_user.WindowsUserID == userName
                                select FT_user.AddInTabName).First();
                if (!string.IsNullOrEmpty(tabLabel))
                    this.tabFinance_ToolsV2.Label = tabLabel;
                int? logintype;
                try
                {
                    logintype = (from FT_user in db.rsUsers
                                 where FT_user.WindowsUserID == userName
                                 select FT_user.LoginType).First();
                }
                catch
                {
                    logintype = 0;
                }
                if (SessionInfo.UserInfo == null)
                {
                    SessionInfo.UserInfo = new UserInfo();
                    SessionInfo.UserInfo.Dictionary = new SessionFileDictionary();
                    SessionInfo.UserInfo.Dictionary.dict = new Dictionary<string, string>();
                }
                SessionInfo.UserInfo.ID = wUserID.ToString();
                ft.Update(wUserID.ToString());
                SessionInfo.UserInfo.SunUserIP = SunIP;
                if (string.IsNullOrEmpty(SunUID))
                    SessionInfo.UserInfo.SunUserID = "";
                else
                    SessionInfo.UserInfo.SunUserID = DEncrypt.Decrypt(SunUID);
                if (string.IsNullOrEmpty(SunUpass))
                    SessionInfo.UserInfo.SunUserPass = "";
                else
                    SessionInfo.UserInfo.SunUserPass = DEncrypt.Decrypt(SunUpass);
                SessionInfo.UserInfo.LoginType = logintype;
                addfolders();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        /// <summary>
        /// Folder initialization (beginning when the exe launch)(Read xml file structure into Ribbon)
        /// </summary>
        internal void addfolders()
        {
            try
            {
                var path = Finance_Tools.RootPath;
                if (path != null)
                {
                    if (Finance_Tools.totalRunTimes == 0 || !Finance_Tools.CompareFileCountWithConfigCount())
                        Finance_Tools.UpdateFolderFileXMLStructure();
                    XmlDocument xdoc = new XmlDocument();
                    xdoc.Load(Finance_Tools.FolderFileXMLPath + Finance_Tools.simpleID + ".xml");
                    XmlNodeList nodeList = xdoc.SelectSingleNode("Nodes").ChildNodes;
                    foreach (XmlNode xn in nodeList)
                    {
                        try
                        {
                            XmlElement xe = xn as XmlElement;
                            string Name = xe.GetAttribute("Name");
                            RibbonMenu rm = new RibbonMenu();
                            rm.Label = Name;
                            rm.Tag = Name;
                            rm.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
                            rm.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
                            rm.OfficeImageId = "FileOpen";
                            this.grpDocuments.Items.Add(rm);
                            XmlNodeList nodeListSub = xn.ChildNodes;
                            foreach (XmlNode xnSub in nodeListSub)
                            {
                                XmlElement xeSub = xnSub as XmlElement;
                                string NameSub = xeSub.GetAttribute("Name");
                                string Description = xeSub.GetAttribute("Description");
                                string id = xeSub.GetAttribute("ID");
                                RibbonToggleButton rt = new RibbonToggleButton();
                                rt.Label = NameSub;
                                rt.OfficeImageId = "SheetInsert";
                                rt.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
                                rt.Description = Description;
                                rt.Tag = id;
                                rt.Click += new System.EventHandler<RibbonControlEventArgs>(rt_click);
                                rm.Items.Add(rt);
                                TemplateAndPath.Add(new KeyValuePair<string, string>(NameSub, id));
                            }
                        }
                        catch
                        {
                            throw new Exception();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("2" + ex.Message);
                LogHelper.WriteLog(typeof(Ribbon2), ex.Message + "Folder RibbonMenu initialization error");
            }
        }
    }
}