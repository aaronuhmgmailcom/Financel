using System.IO;
using Microsoft.Office.Tools.Ribbon;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Windows.Forms;
using ExcelAddIn4.Common;
using System.Data;
using System.Linq;
namespace ExcelAddIn4
{
    partial class Ribbon2 : Microsoft.Office.Tools.Ribbon.OfficeRibbon
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
        /// <summary>
        /// 
        /// </summary>
        public static int ButtonViewCount
        {
            get;
            set;
        }
        public Ribbon2()
        {
            try
            {
                if (Config())
                {
                    InitializeComponent();
                    DoLogin(this);
                }
                if (SessionInfo.UserInfo.LoginType == null || SessionInfo.UserInfo.LoginType == 0)//0.admin,1.not admin
                    this.grpAdmin.Visible = true;
                else
                    this.grpAdmin.Visible = false;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                LogHelper.WriteLog(typeof(Ribbon2), ex.Message);
            }
        }
        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }
        #region Component Designer generated code
        private bool buttonExist(string tag)
        {
            for (int i = 0; i < this.grpATC.Items.Count; i++)
            {
                if (this.grpATC.Items[i].Tag.ToString() == tag)
                    return true;
            }
            return false;
        }
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabFinance_ToolsV2 = new RibbonTab();
            this.ftControls = new RibbonGroup();
            this.btnView_Doc = new RibbonButton();
            this.btnView_Journal = new RibbonButton();
            this.btnView_PDF = new RibbonButton();
            this.btnView_ViewXML = new RibbonButton();
            this.btnView_Rep = new RibbonButton();
            this.btnTemplate_Save = new RibbonButton();
            this.btnAd_DocView = new RibbonButton();
            this.btnAd_AttachPDF = new RibbonButton();
            this.btnAd_NewButton = new RibbonButton();
            this.btnAd_DataFieldsSetting = new RibbonButton(); ;
            this.btnPreferences = new RibbonButton();
            btnCreateXMLorTextFile = new RibbonButton();
            this.btnUpgrade = new RibbonButton();
            this.btnSecurity = new RibbonButton();
            this.btnATVP_Output = new RibbonCheckBox();
            this.btnATVP_View = new RibbonCheckBox();
            this.btnATVP_Help = new RibbonCheckBox();
            this.grpDocuments = new RibbonGroup();
            this.grpAdmin = new RibbonGroup();
            this.grpATVP = new RibbonGroup();
            this.grpATC = new RibbonGroup();
            this.grpBlank = new RibbonGroup();
            this.suspendToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.contextMenuStrip2 = new System.Windows.Forms.ContextMenuStrip();
            this.tabFinance_ToolsV2.SuspendLayout();
            this.contextMenuStrip2.SuspendLayout();
            this.ftControls.SuspendLayout();
            // 
            // tabFinance_Tools
            // 
            this.tabFinance_ToolsV2.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabFinance_ToolsV2.Groups.Add(grpBlank);
            this.tabFinance_ToolsV2.Groups.Add(this.ftControls);
            this.tabFinance_ToolsV2.Groups.Add(this.grpATC);
            this.tabFinance_ToolsV2.Groups.Add(this.grpDocuments);
            this.tabFinance_ToolsV2.Groups.Add(this.grpAdmin);
            this.tabFinance_ToolsV2.Groups.Add(this.grpATVP);
            this.tabFinance_ToolsV2.Label = "Finance Tools V2";
            this.tabFinance_ToolsV2.Name = "tabFinance_Toolsv2";
            this.grpDocuments.Label = "Templates";
            // 
            // ftControls
            // 
            this.ftControls.Items.Add(this.btnView_Doc);
            this.ftControls.Items.Add(this.btnView_Journal);
            this.ftControls.Items.Add(this.btnView_PDF);
            this.ftControls.Items.Add(this.btnView_Rep);
            this.ftControls.Items.Add(this.btnView_ViewXML);
            this.ftControls.Label = "Global Controls";
            this.ftControls.Name = "ftControls";
            // 
            // btnView_Doc
            // 
            this.btnView_Doc.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnView_Doc.Description = "View Document";
            this.btnView_Doc.Label = "View Document";
            this.btnView_Doc.Name = "btnView_Doc";
            this.btnView_Doc.OfficeImageId = "FindDialogExcel";
            this.btnView_Doc.ShowImage = true;
            this.btnView_Doc.SuperTip = "Click here to drill down to document";
            this.btnView_Doc.Click += new System.EventHandler<RibbonControlEventArgs>(this.btnCtrl_VD_Click);
            // 
            // btnView_Journal
            // 
            this.btnView_Journal.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnView_Journal.Description = "View Journal";
            this.btnView_Journal.Label = "View Journal";
            this.btnView_Journal.Name = "btnView_Journal";
            this.btnView_Journal.OfficeImageId = "FunctionsFinancialInsertGallery";
            this.btnView_Journal.ShowImage = true;
            this.btnView_Journal.SuperTip = "Click here to drill down to journal";
            this.btnView_Journal.Click += new System.EventHandler<RibbonControlEventArgs>(this.btnJournal_Click);
            // 
            // btnView_PDF
            // 
            this.btnView_PDF.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnView_PDF.Description = "View PDF";
            this.btnView_PDF.Label = "View PDF";
            this.btnView_PDF.Name = "btnView_PDF";
            this.btnView_PDF.Image = ExcelAddIn4.Properties.Resources.resizeApi;
            this.btnView_PDF.ShowImage = true;
            this.btnView_PDF.SuperTip = "Click here to drill down to PDF";
            this.btnView_PDF.Click += new System.EventHandler<RibbonControlEventArgs>(this.btnPDF_Click);
            // 
            // btnView_Rep
            // 
            this.btnView_Rep.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnView_Rep.Description = "Transaction Search";
            this.btnView_Rep.Label = "Transaction Search";
            this.btnView_Rep.Name = "btnView_Rep";
            this.btnView_Rep.OfficeImageId = "AdpDiagramTableModesMenu";
            this.btnView_Rep.ShowImage = true;
            this.btnView_Rep.SuperTip = "Click here to start transaction search!";
            this.btnView_Rep.Click += new System.EventHandler<RibbonControlEventArgs>(this.btnCtrl_VR_Click);
            // 
            // btnView_ViewXML
            // 
            this.btnView_ViewXML.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnView_ViewXML.Description = "View XML";
            this.btnView_ViewXML.Label = "View XML";
            this.btnView_ViewXML.Name = "btnView_ViewXML";
            this.btnView_ViewXML.OfficeImageId = "SourceControlShowDifferences";
            this.btnView_ViewXML.ShowImage = true;
            this.btnView_ViewXML.SuperTip = "Click here to start view xml!";
            this.btnView_ViewXML.Click += new System.EventHandler<RibbonControlEventArgs>(this.btnCtrl_VX_Click);
            // 
            // btnTemplate_Save
            // 
            this.btnTemplate_Save.Label = "Save/Amend Template";
            this.btnTemplate_Save.Name = "btnTemplate_Save";
            this.btnTemplate_Save.OfficeImageId = "FileSaveACopy";
            this.btnTemplate_Save.ShowImage = true;
            this.btnTemplate_Save.Click += new System.EventHandler<RibbonControlEventArgs>(this.btnTemplate_Save_Click);
            BasePage.VerifyButton("Save/Amend Template - Global", "3", this.btnTemplate_Save);
            // 
            // grpAdmin
            // 
            this.grpAdmin.Items.Add(this.btnAd_DocView);
            this.grpAdmin.Items.Add(this.btnAd_AttachPDF);
            this.grpAdmin.Items.Add(this.btnAd_NewButton);
            this.grpAdmin.Items.Add(this.btnAd_DataFieldsSetting);
            this.grpAdmin.Items.Add(this.btnCreateXMLorTextFile);
            this.grpAdmin.Items.Add(this.btnTemplate_Save);
            this.grpAdmin.Items.Add(this.btnPreferences);
            this.grpAdmin.Items.Add(this.btnUpgrade);
            this.grpAdmin.Items.Add(this.btnSecurity);
            this.grpAdmin.Label = "Administration";
            this.grpAdmin.Name = "grpAdmin";
            // 
            // grpATVP
            // 
            this.grpATVP.Items.Add(this.btnATVP_View);
            this.grpATVP.Items.Add(this.btnATVP_Output);
            this.grpATVP.Items.Add(this.btnATVP_Help);
            this.grpATVP.Label = "Visible Panes";
            this.grpATVP.Name = "grpATVP";
            // 
            // grpATC
            // 
            string userName = WindowsIdentity.GetCurrent().Name;
            var wUserID = (from FT_user in db.rsUsers
                           where FT_user.WindowsUserID == userName
                           select FT_user.ft_id).First();
            System.Collections.Generic.List<string> list = ft.GetUserGroups(wUserID.ToString());
            ButtonViewCount = 0;
            for (int i = 0; i < list.Count; i++)
            {
                string groupid = list[i];
                bool groupdisable = (bool)ft.GetGroupDisableByID(int.Parse(groupid));
                if (groupdisable)
                    continue;
                DataTable dt = ft.GetGroupButtonsView(groupid);

                var query = from t in dt.AsEnumerable()
                            group t by new { t1 = t.Field<string>("ButtonGroup"), t2 = t.Field<string>("ButtonName") } into m
                            select new
                            {
                                ButtonName = m.First().Field<string>("ButtonName"),
                                ButtonText = m.First().Field<string>("ButtonText"),
                                ButtonIcon = m.First().Field<string>("ButtonIcon"),
                                ButtonGroup = m.First().Field<string>("ButtonGroup"),
                                ButtonSize = m.First().Field<string>("ButtonSize"),
                                ButtonOrder = m.First().Field<int>("ButtonOrder"),
                                GroupOrder = m.First().Field<int>("GroupOrder"),
                                TAG = m.First().Field<string>("TemplateID") + "," + m.First().Field<string>("ButtonName")
                            };
                string currentGroup = string.Empty;
                foreach (var employee in query)
                {//Microsoft.Office.Tools.Ribbon.RibbonSeparator rs = new Microsoft.Office.Tools.Ribbon.RibbonSeparator();
                    if (buttonExist(employee.TAG)) continue;
                    Microsoft.Office.Tools.Ribbon.RibbonButton newbutton = new RibbonButton();
                    newbutton.Label = employee.ButtonName;//newbutton.Name = "btnATC_NewButton" + ButtonViewCount.ToString();
                    newbutton.OfficeImageId = employee.ButtonIcon;
                    newbutton.Tag = employee.TAG;
                    newbutton.SuperTip = employee.ButtonText;
                    newbutton.ShowImage = true;
                    if (employee.ButtonSize == "32")
                        newbutton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
                    else
                        newbutton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
                    newbutton.Click += new System.EventHandler<RibbonControlEventArgs>(this.btnNewButton_Click);
                    newbutton.Visible = false;
                    this.grpATC.Items.Add(newbutton);
                    ButtonViewCount++;
                }
            }
            this.grpATC.Label = "Template Controls";
            this.grpATC.Name = "grpATC";
            // 
            // btnAd_DocView
            // 
            this.btnAd_DocView.Label = "Add Document View";
            this.btnAd_DocView.Name = "btnAd_DocView";
            this.btnAd_DocView.OfficeImageId = "ObjectAddText";
            this.btnAd_DocView.ShowImage = true;
            this.btnAd_DocView.Click += new System.EventHandler<RibbonControlEventArgs>(this.btnAd_DocView_Click);
            BasePage.VerifyButton("Add Document View - Global", "3", this.btnAd_DocView);
            // 
            // btnAd_AttachPDF
            // 
            this.btnAd_AttachPDF.Label = "Attach PDF";
            this.btnAd_AttachPDF.Name = "btnAd_AttachPDF";
            this.btnAd_AttachPDF.OfficeImageId = "FileSaveAsPdfOrXps";
            this.btnAd_AttachPDF.ShowImage = true;
            this.btnAd_AttachPDF.Click += new System.EventHandler<RibbonControlEventArgs>(this.btnAd_AttachPDF_Click);
            BasePage.VerifyButton("Attach PDF - Global", "3", this.btnAd_AttachPDF);
            // 
            // btnAd_NewButton
            // 
            this.btnAd_NewButton.Label = "Create New Action";
            this.btnAd_NewButton.Name = "btnAd_NewButton";
            this.btnAd_NewButton.OfficeImageId = "PivotSubtotal";
            this.btnAd_NewButton.ShowImage = true;
            this.btnAd_NewButton.Click += new System.EventHandler<RibbonControlEventArgs>(this.btnAd_NewButton_Click);
            BasePage.VerifyButton("Create New Action - Global", "3", this.btnAd_NewButton);
            // 
            // btnATVP_View
            // 
            this.btnATVP_View.Label = "View";
            this.btnATVP_View.Name = "btnATVP_View";
            this.btnATVP_View.Click += new System.EventHandler<RibbonControlEventArgs>(this.btnATVP_View_Click);
            // 
            // btnPreferences
            // 
            this.btnPreferences.Label = "Preferences...";
            this.btnPreferences.Name = "btnPreferences";
            this.btnPreferences.OfficeImageId = "DefinePrintStyles";
            this.btnPreferences.ShowImage = true;
            this.btnPreferences.Click += new System.EventHandler<RibbonControlEventArgs>(this.btnPreferences_Preferences_Click);
            // 
            // btnCreateXMLorTextFile
            // 
            this.btnCreateXMLorTextFile.Label = "Create XML/Text Profile";
            this.btnCreateXMLorTextFile.Name = "btnCreateXMLorTextFile";
            this.btnCreateXMLorTextFile.OfficeImageId = "XmlExport";
            this.btnCreateXMLorTextFile.ShowImage = true;
            this.btnCreateXMLorTextFile.Click += new System.EventHandler<RibbonControlEventArgs>(this.btnCreateXMLorTextFile_Click);
            BasePage.VerifyButton("Create XML or Text File Profile - Global", "3", this.btnCreateXMLorTextFile);
            // 
            // btnUpgrade
            // 
            this.btnUpgrade.Label = "Synchronization";
            this.btnUpgrade.Name = "Synchronization";
            this.btnUpgrade.OfficeImageId = "TableExportTableToSharePointList";
            this.btnUpgrade.ShowImage = true;
            this.btnUpgrade.Click += new System.EventHandler<RibbonControlEventArgs>(this.btnUpgrade_Click);
            BasePage.VerifyButton("Upgrade - Global", "3", this.btnUpgrade);
            // 
            // btnSecurity
            // 
            this.btnSecurity.Label = "Security";
            this.btnSecurity.Name = "Users";
            this.btnSecurity.OfficeImageId = "FileDocumentEncrypt";
            this.btnSecurity.ShowImage = true;
            this.btnSecurity.Click += new System.EventHandler<RibbonControlEventArgs>(this.btnSecurity_Click);
            BasePage.VerifyButton("Security - Global", "3", this.btnSecurity);

            // 
            // btnATVP_Output
            // 
            this.btnATVP_Output.Label = "Settings";
            this.btnATVP_Output.Name = "btnATVP_Output";
            this.btnATVP_Output.Click += new System.EventHandler<RibbonControlEventArgs>(this.btnATVP_Output_Click);
            // 
            // btnATVP_Help
            // 
            this.btnATVP_Help.Label = "Help";
            this.btnATVP_Help.Name = "btnATVP_Help";
            this.btnATVP_Help.Click += new System.EventHandler<RibbonControlEventArgs>(this.btnATVP_Help_Click);
            // 
            // btnAd_DataFieldsSetting
            // 
            this.btnAd_DataFieldsSetting.Label = "Amend Code Descriptions";
            this.btnAd_DataFieldsSetting.Name = "btnAd_DataFieldsSetting";
            this.btnAd_DataFieldsSetting.OfficeImageId = "DatabaseObjectDependencies";
            this.btnAd_DataFieldsSetting.ShowImage = true;
            this.btnAd_DataFieldsSetting.Click += new System.EventHandler<RibbonControlEventArgs>(this.btnAd_DataFieldsSetting_Click);
            BasePage.VerifyButton("Amend Code Descriptions - Global", "3", this.btnAd_DataFieldsSetting);
            // 
            // suspendToolStripMenuItem
            // 
            this.suspendToolStripMenuItem.Name = "suspendToolStripMenuItem";
            this.suspendToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.suspendToolStripMenuItem.Text = "Abort";
            this.suspendToolStripMenuItem.Enabled = false;
            this.suspendToolStripMenuItem.Click += new System.EventHandler(this.suspendToolStripMenuItem_Click);
            // 
            // contextMenuStrip2
            // 
            this.contextMenuStrip2.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.suspendToolStripMenuItem});
            this.contextMenuStrip2.Name = "contextMenuStrip2";
            this.contextMenuStrip2.Size = new System.Drawing.Size(153, 136);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabFinance_ToolsV2);
            this.Load += new System.EventHandler<RibbonUIEventArgs>(this.Ribbon2_Load);
            this.tabFinance_ToolsV2.ResumeLayout(false);
            this.tabFinance_ToolsV2.PerformLayout();
            this.contextMenuStrip2.ResumeLayout(false);
            this.ftControls.ResumeLayout(false);
            this.ftControls.PerformLayout();
        }
        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabFinance_ToolsV2;
        internal RibbonGroup ftControls;
        internal RibbonButton btnView_Doc;
        internal RibbonButton btnView_Journal;
        internal RibbonButton btnView_PDF;
        internal RibbonButton btnView_Rep;
        internal RibbonButton btnView_ViewXML;
        internal RibbonButton btnTemplate_Save;
        internal RibbonButton btnAd_DocView;
        internal RibbonButton btnAd_AttachPDF;
        internal RibbonButton btnAd_NewButton;
        internal RibbonButton btnAd_DataFieldsSetting;
        internal RibbonButton btnPreferences;
        internal RibbonButton btnCreateXMLorTextFile;
        internal RibbonButton btnUpgrade;
        internal RibbonButton btnSecurity;
        internal RibbonCheckBox btnATVP_Output;
        public RibbonCheckBox btnATVP_View;
        internal RibbonCheckBox btnATVP_Help;
        internal RibbonGroup grpDocuments;
        internal RibbonGroup grpAdmin;
        internal RibbonGroup grpATVP;
        internal RibbonGroup grpATC;
        internal RibbonGroup grpBlank;
        internal System.Windows.Forms.ToolStripMenuItem suspendToolStripMenuItem;
        internal System.Windows.Forms.ContextMenuStrip contextMenuStrip2;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon2 ribbFinance_Tools
        {
            get { return this.GetRibbon<Ribbon2>(); }
        }
    }
}
