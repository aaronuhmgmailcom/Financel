namespace ExcelAddIn2
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
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

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.ExcelAddIn2 = this.Factory.CreateRibbonTab();
            this.gpTaskPane = this.Factory.CreateRibbonGroup();
            this.btnOpen = this.Factory.CreateRibbonButton();
            this.btnClose = this.Factory.CreateRibbonButton();
            this.ExcelAddIn2.SuspendLayout();
            this.gpTaskPane.SuspendLayout();
            // 
            // ExcelAddIn2
            // 
            this.ExcelAddIn2.Groups.Add(this.gpTaskPane);
            this.ExcelAddIn2.Label = "ExcelAddIn2";
            this.ExcelAddIn2.Name = "ExcelAddIn2";
            // 
            // gpTaskPane
            // 
            this.gpTaskPane.Items.Add(this.btnOpen);
            this.gpTaskPane.Items.Add(this.btnClose);
            this.gpTaskPane.Label = "Task Pane";
            this.gpTaskPane.Name = "gpTaskPane";
            // 
            // btnOpen
            // 
            this.btnOpen.Image = ((System.Drawing.Image)(resources.GetObject("btnOpen.Image")));
            this.btnOpen.Label = "Open Task Pane";
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.ShowImage = true;
            this.btnOpen.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOpen_Click);
            // 
            // btnClose
            // 
            this.btnClose.Image = global::ExcelAddIn2.Properties.Resources.close;
            this.btnClose.Label = "Close Task Pane";
            this.btnClose.Name = "btnClose";
            this.btnClose.ShowImage = true;
            this.btnClose.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnClose_Click);
            // 
            // MyRibbon
            // 
            this.Name = "MyRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.ExcelAddIn2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MyRibbon_Load);
            this.ExcelAddIn2.ResumeLayout(false);
            this.ExcelAddIn2.PerformLayout();
            this.gpTaskPane.ResumeLayout(false);
            this.gpTaskPane.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab ExcelAddIn2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup gpTaskPane;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpen;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClose;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 MyRibbon
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
