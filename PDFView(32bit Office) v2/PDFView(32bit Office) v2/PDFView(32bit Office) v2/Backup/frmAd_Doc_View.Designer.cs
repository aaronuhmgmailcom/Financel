namespace ExcelAddIn4
{
    partial class frmAd_Doc_View
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmAd_Doc_View));
            this.fbdADV = new System.Windows.Forms.FolderBrowserDialog();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.ckb_isAWebFile = new System.Windows.Forms.CheckBox();
            this.txtADV_ModuleName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.tbADV_Filepath = new System.Windows.Forms.RichTextBox();
            this.cbADV_FType = new System.Windows.Forms.ComboBox();
            this.btnADV_Create = new System.Windows.Forms.Button();
            this.tbADV_Macro01 = new System.Windows.Forms.TextBox();
            this.lbADV_Macro01 = new System.Windows.Forms.Label();
            this.lbADV_Macro = new System.Windows.Forms.Label();
            this.lbADV_FType = new System.Windows.Forms.Label();
            this.tbADV_File = new System.Windows.Forms.TextBox();
            this.lbADV_FName = new System.Windows.Forms.Label();
            this.chkADV_UseRef = new System.Windows.Forms.CheckBox();
            this.btnADV_Folder = new System.Windows.Forms.Button();
            this.lbADV_Folder = new System.Windows.Forms.Label();
            this.tbADV_Prefix = new System.Windows.Forms.TextBox();
            this.lbADV_Prefix = new System.Windows.Forms.Label();
            this.lbADV_Type = new System.Windows.Forms.Label();
            this.tbADV_Type = new System.Windows.Forms.TextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.errorProvider1 = new System.Windows.Forms.ErrorProvider(this.components);
            this.fbdAd_UpdateFolder = new System.Windows.Forms.FolderBrowserDialog();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.panel22 = new System.Windows.Forms.Panel();
            this.label24 = new System.Windows.Forms.Label();
            this.tbColumn2 = new System.Windows.Forms.TextBox();
            this.tbConnectStr = new System.Windows.Forms.TextBox();
            this.lbConnectStr = new System.Windows.Forms.Label();
            this.tbUserID = new System.Windows.Forms.TextBox();
            this.lbUserID = new System.Windows.Forms.Label();
            this.tbPassword = new System.Windows.Forms.TextBox();
            this.lbSQL = new System.Windows.Forms.Label();
            this.lbPassword = new System.Windows.Forms.Label();
            this.tbSQL = new System.Windows.Forms.TextBox();
            this.panel21 = new System.Windows.Forms.Panel();
            this.label4 = new System.Windows.Forms.Label();
            this.txtPath = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.chkViewFromDB = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtInvName = new System.Windows.Forms.TextBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.cbDbQ = new System.Windows.Forms.CheckBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.rbPdf = new System.Windows.Forms.RadioButton();
            this.rbWb = new System.Windows.Forms.RadioButton();
            this.btnSave = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.panel22.SuspendLayout();
            this.panel21.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(620, 598);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.ckb_isAWebFile);
            this.tabPage1.Controls.Add(this.txtADV_ModuleName);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.tbADV_Filepath);
            this.tabPage1.Controls.Add(this.cbADV_FType);
            this.tabPage1.Controls.Add(this.btnADV_Create);
            this.tabPage1.Controls.Add(this.tbADV_Macro01);
            this.tabPage1.Controls.Add(this.lbADV_Macro01);
            this.tabPage1.Controls.Add(this.lbADV_Macro);
            this.tabPage1.Controls.Add(this.lbADV_FType);
            this.tabPage1.Controls.Add(this.tbADV_File);
            this.tabPage1.Controls.Add(this.lbADV_FName);
            this.tabPage1.Controls.Add(this.chkADV_UseRef);
            this.tabPage1.Controls.Add(this.btnADV_Folder);
            this.tabPage1.Controls.Add(this.lbADV_Folder);
            this.tabPage1.Controls.Add(this.tbADV_Prefix);
            this.tabPage1.Controls.Add(this.lbADV_Prefix);
            this.tabPage1.Controls.Add(this.lbADV_Type);
            this.tabPage1.Controls.Add(this.tbADV_Type);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(612, 572);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Add Document View";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // ckb_isAWebFile
            // 
            this.ckb_isAWebFile.AutoSize = true;
            this.ckb_isAWebFile.Location = new System.Drawing.Point(105, 10);
            this.ckb_isAWebFile.Name = "ckb_isAWebFile";
            this.ckb_isAWebFile.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.ckb_isAWebFile.Size = new System.Drawing.Size(161, 17);
            this.ckb_isAWebFile.TabIndex = 43;
            this.ckb_isAWebFile.Text = "?Is it a Web Document View";
            this.ckb_isAWebFile.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.ckb_isAWebFile.UseVisualStyleBackColor = true;
            this.ckb_isAWebFile.CheckedChanged += new System.EventHandler(this.ckb_isAWebFile_CheckedChanged);
            // 
            // txtADV_ModuleName
            // 
            this.txtADV_ModuleName.Enabled = false;
            this.txtADV_ModuleName.Location = new System.Drawing.Point(202, 370);
            this.txtADV_ModuleName.MaxLength = 50;
            this.txtADV_ModuleName.Name = "txtADV_ModuleName";
            this.txtADV_ModuleName.Size = new System.Drawing.Size(299, 20);
            this.txtADV_ModuleName.TabIndex = 39;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(107, 373);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(79, 13);
            this.label1.TabIndex = 38;
            this.label1.Text = "Module, Name:";
            // 
            // tbADV_Filepath
            // 
            this.tbADV_Filepath.Enabled = false;
            this.tbADV_Filepath.Location = new System.Drawing.Point(110, 131);
            this.tbADV_Filepath.Name = "tbADV_Filepath";
            this.tbADV_Filepath.Size = new System.Drawing.Size(397, 47);
            this.tbADV_Filepath.TabIndex = 37;
            this.tbADV_Filepath.Text = "";
            // 
            // cbADV_FType
            // 
            this.cbADV_FType.FormattingEnabled = true;
            this.cbADV_FType.Location = new System.Drawing.Point(202, 296);
            this.cbADV_FType.Name = "cbADV_FType";
            this.cbADV_FType.Size = new System.Drawing.Size(69, 21);
            this.cbADV_FType.TabIndex = 36;
            this.cbADV_FType.SelectedIndexChanged += new System.EventHandler(this.cbADV_FType_SelectedIndexChanged);
            // 
            // btnADV_Create
            // 
            this.btnADV_Create.Location = new System.Drawing.Point(205, 435);
            this.btnADV_Create.Name = "btnADV_Create";
            this.btnADV_Create.Size = new System.Drawing.Size(221, 23);
            this.btnADV_Create.TabIndex = 42;
            this.btnADV_Create.Text = "Create Document View";
            this.btnADV_Create.UseVisualStyleBackColor = true;
            this.btnADV_Create.Click += new System.EventHandler(this.btnADV_Create_Click);
            // 
            // tbADV_Macro01
            // 
            this.tbADV_Macro01.Enabled = false;
            this.tbADV_Macro01.Location = new System.Drawing.Point(202, 396);
            this.tbADV_Macro01.MaxLength = 50;
            this.tbADV_Macro01.Name = "tbADV_Macro01";
            this.tbADV_Macro01.Size = new System.Drawing.Size(299, 20);
            this.tbADV_Macro01.TabIndex = 41;
            // 
            // lbADV_Macro01
            // 
            this.lbADV_Macro01.AutoSize = true;
            this.lbADV_Macro01.Location = new System.Drawing.Point(107, 399);
            this.lbADV_Macro01.Name = "lbADV_Macro01";
            this.lbADV_Macro01.Size = new System.Drawing.Size(80, 13);
            this.lbADV_Macro01.TabIndex = 40;
            this.lbADV_Macro01.Text = "Macro1, Name:";
            // 
            // lbADV_Macro
            // 
            this.lbADV_Macro.AutoSize = true;
            this.lbADV_Macro.Location = new System.Drawing.Point(107, 344);
            this.lbADV_Macro.Name = "lbADV_Macro";
            this.lbADV_Macro.Size = new System.Drawing.Size(134, 13);
            this.lbADV_Macro.TabIndex = 35;
            this.lbADV_Macro.Text = "Run Macros on File Open?";
            // 
            // lbADV_FType
            // 
            this.lbADV_FType.AutoSize = true;
            this.lbADV_FType.Location = new System.Drawing.Point(107, 300);
            this.lbADV_FType.Name = "lbADV_FType";
            this.lbADV_FType.Size = new System.Drawing.Size(49, 13);
            this.lbADV_FType.TabIndex = 34;
            this.lbADV_FType.Text = "File type:";
            // 
            // tbADV_File
            // 
            this.tbADV_File.Enabled = false;
            this.tbADV_File.Location = new System.Drawing.Point(202, 213);
            this.tbADV_File.MaxLength = 50;
            this.tbADV_File.Name = "tbADV_File";
            this.tbADV_File.Size = new System.Drawing.Size(299, 20);
            this.tbADV_File.TabIndex = 33;
            // 
            // lbADV_FName
            // 
            this.lbADV_FName.AutoSize = true;
            this.lbADV_FName.Location = new System.Drawing.Point(107, 217);
            this.lbADV_FName.Name = "lbADV_FName";
            this.lbADV_FName.Size = new System.Drawing.Size(52, 13);
            this.lbADV_FName.TabIndex = 32;
            this.lbADV_FName.Text = "Filename:";
            // 
            // chkADV_UseRef
            // 
            this.chkADV_UseRef.AutoSize = true;
            this.chkADV_UseRef.Checked = true;
            this.chkADV_UseRef.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkADV_UseRef.Location = new System.Drawing.Point(110, 187);
            this.chkADV_UseRef.Name = "chkADV_UseRef";
            this.chkADV_UseRef.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.chkADV_UseRef.Size = new System.Drawing.Size(213, 17);
            this.chkADV_UseRef.TabIndex = 31;
            this.chkADV_UseRef.Text = "Use Transaction Reference as filename";
            this.chkADV_UseRef.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            this.chkADV_UseRef.UseVisualStyleBackColor = true;
            this.chkADV_UseRef.Click += new System.EventHandler(this.chkADV_UseRef_CheckedChanged);
            // 
            // btnADV_Folder
            // 
            this.btnADV_Folder.Location = new System.Drawing.Point(110, 101);
            this.btnADV_Folder.Name = "btnADV_Folder";
            this.btnADV_Folder.Size = new System.Drawing.Size(54, 25);
            this.btnADV_Folder.TabIndex = 29;
            this.btnADV_Folder.Text = "Browse";
            this.btnADV_Folder.UseVisualStyleBackColor = true;
            this.btnADV_Folder.Click += new System.EventHandler(this.btnADV_Folder_Click);
            // 
            // lbADV_Folder
            // 
            this.lbADV_Folder.AutoSize = true;
            this.lbADV_Folder.Location = new System.Drawing.Point(107, 83);
            this.lbADV_Folder.Name = "lbADV_Folder";
            this.lbADV_Folder.Size = new System.Drawing.Size(244, 13);
            this.lbADV_Folder.TabIndex = 30;
            this.lbADV_Folder.Text = "Folder containing document or view document file:";
            // 
            // tbADV_Prefix
            // 
            this.tbADV_Prefix.Location = new System.Drawing.Point(202, 57);
            this.tbADV_Prefix.MaxLength = 10;
            this.tbADV_Prefix.Name = "tbADV_Prefix";
            this.tbADV_Prefix.Size = new System.Drawing.Size(69, 20);
            this.tbADV_Prefix.TabIndex = 26;
            // 
            // lbADV_Prefix
            // 
            this.lbADV_Prefix.AutoSize = true;
            this.lbADV_Prefix.Location = new System.Drawing.Point(107, 60);
            this.lbADV_Prefix.Name = "lbADV_Prefix";
            this.lbADV_Prefix.Size = new System.Drawing.Size(88, 13);
            this.lbADV_Prefix.TabIndex = 28;
            this.lbADV_Prefix.Text = "Document Prefix:";
            // 
            // lbADV_Type
            // 
            this.lbADV_Type.AutoSize = true;
            this.lbADV_Type.Location = new System.Drawing.Point(107, 35);
            this.lbADV_Type.Name = "lbADV_Type";
            this.lbADV_Type.Size = new System.Drawing.Size(82, 13);
            this.lbADV_Type.TabIndex = 27;
            this.lbADV_Type.Text = "Document type:";
            // 
            // tbADV_Type
            // 
            this.tbADV_Type.Location = new System.Drawing.Point(202, 32);
            this.tbADV_Type.MaxLength = 50;
            this.tbADV_Type.Name = "tbADV_Type";
            this.tbADV_Type.Size = new System.Drawing.Size(310, 20);
            this.tbADV_Type.TabIndex = 25;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.groupBox2);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(612, 572);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "PDF Viewer";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // errorProvider1
            // 
            this.errorProvider1.ContainerControl = this;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.panel22);
            this.groupBox2.Controls.Add(this.panel21);
            this.groupBox2.Controls.Add(this.groupBox5);
            this.groupBox2.Controls.Add(this.groupBox4);
            this.groupBox2.Controls.Add(this.btnSave);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox2.Location = new System.Drawing.Point(3, 3);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(606, 566);
            this.groupBox2.TabIndex = 2;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "View Path and Column to save by:";
            // 
            // panel22
            // 
            this.panel22.Controls.Add(this.label24);
            this.panel22.Controls.Add(this.tbColumn2);
            this.panel22.Controls.Add(this.tbConnectStr);
            this.panel22.Controls.Add(this.lbConnectStr);
            this.panel22.Controls.Add(this.tbUserID);
            this.panel22.Controls.Add(this.lbUserID);
            this.panel22.Controls.Add(this.tbPassword);
            this.panel22.Controls.Add(this.lbSQL);
            this.panel22.Controls.Add(this.lbPassword);
            this.panel22.Controls.Add(this.tbSQL);
            this.panel22.Location = new System.Drawing.Point(6, 217);
            this.panel22.Name = "panel22";
            this.panel22.Size = new System.Drawing.Size(595, 131);
            this.panel22.TabIndex = 38;
            this.panel22.Visible = false;
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Location = new System.Drawing.Point(61, 104);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(45, 13);
            this.label24.TabIndex = 36;
            this.label24.Text = "Column:";
            this.label24.Visible = false;
            // 
            // tbColumn2
            // 
            this.tbColumn2.Location = new System.Drawing.Point(130, 101);
            this.tbColumn2.Name = "tbColumn2";
            this.tbColumn2.Size = new System.Drawing.Size(32, 20);
            this.tbColumn2.TabIndex = 35;
            this.tbColumn2.Visible = false;
            // 
            // tbConnectStr
            // 
            this.tbConnectStr.ForeColor = System.Drawing.SystemColors.ActiveBorder;
            this.tbConnectStr.Location = new System.Drawing.Point(130, 4);
            this.tbConnectStr.Multiline = true;
            this.tbConnectStr.Name = "tbConnectStr";
            this.tbConnectStr.Size = new System.Drawing.Size(457, 41);
            this.tbConnectStr.TabIndex = 26;
            this.tbConnectStr.Text = "Data Source=myDatabaseName;Initial Catalog=myDBTables;User ID=[UserID];Password=[" +
    "Password];";
            // 
            // lbConnectStr
            // 
            this.lbConnectStr.AutoSize = true;
            this.lbConnectStr.Location = new System.Drawing.Point(11, 17);
            this.lbConnectStr.Name = "lbConnectStr";
            this.lbConnectStr.Size = new System.Drawing.Size(94, 13);
            this.lbConnectStr.TabIndex = 27;
            this.lbConnectStr.Text = "Connection String:";
            // 
            // tbUserID
            // 
            this.tbUserID.Location = new System.Drawing.Point(130, 50);
            this.tbUserID.Name = "tbUserID";
            this.tbUserID.Size = new System.Drawing.Size(141, 20);
            this.tbUserID.TabIndex = 28;
            // 
            // lbUserID
            // 
            this.lbUserID.AutoSize = true;
            this.lbUserID.Location = new System.Drawing.Point(61, 52);
            this.lbUserID.Name = "lbUserID";
            this.lbUserID.Size = new System.Drawing.Size(43, 13);
            this.lbUserID.TabIndex = 29;
            this.lbUserID.Text = "UserID:";
            // 
            // tbPassword
            // 
            this.tbPassword.Location = new System.Drawing.Point(360, 50);
            this.tbPassword.Name = "tbPassword";
            this.tbPassword.PasswordChar = '*';
            this.tbPassword.Size = new System.Drawing.Size(141, 20);
            this.tbPassword.TabIndex = 30;
            this.tbPassword.UseSystemPasswordChar = true;
            // 
            // lbSQL
            // 
            this.lbSQL.AutoSize = true;
            this.lbSQL.Location = new System.Drawing.Point(45, 78);
            this.lbSQL.Name = "lbSQL";
            this.lbSQL.Size = new System.Drawing.Size(61, 13);
            this.lbSQL.TabIndex = 33;
            this.lbSQL.Text = "SQL String:";
            // 
            // lbPassword
            // 
            this.lbPassword.AutoSize = true;
            this.lbPassword.Location = new System.Drawing.Point(280, 52);
            this.lbPassword.Name = "lbPassword";
            this.lbPassword.Size = new System.Drawing.Size(56, 13);
            this.lbPassword.TabIndex = 31;
            this.lbPassword.Text = "Password:";
            // 
            // tbSQL
            // 
            this.tbSQL.ForeColor = System.Drawing.SystemColors.ActiveBorder;
            this.tbSQL.Location = new System.Drawing.Point(130, 75);
            this.tbSQL.Name = "tbSQL";
            this.tbSQL.Size = new System.Drawing.Size(457, 20);
            this.tbSQL.TabIndex = 32;
            this.tbSQL.Text = "Select di.urlname from docInvoice di Where di.invoiceNum = ‘~’";
            // 
            // panel21
            // 
            this.panel21.Controls.Add(this.label4);
            this.panel21.Controls.Add(this.txtPath);
            this.panel21.Controls.Add(this.btnBrowse);
            this.panel21.Controls.Add(this.chkViewFromDB);
            this.panel21.Controls.Add(this.label3);
            this.panel21.Controls.Add(this.txtInvName);
            this.panel21.Location = new System.Drawing.Point(6, 99);
            this.panel21.Name = "panel21";
            this.panel21.Size = new System.Drawing.Size(595, 102);
            this.panel21.TabIndex = 37;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(16, 11);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(32, 13);
            this.label4.TabIndex = 20;
            this.label4.Text = "Path:";
            // 
            // txtPath
            // 
            this.txtPath.Location = new System.Drawing.Point(67, 8);
            this.txtPath.Name = "txtPath";
            this.txtPath.Size = new System.Drawing.Size(249, 20);
            this.txtPath.TabIndex = 17;
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(322, 6);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnBrowse.TabIndex = 16;
            this.btnBrowse.Text = "Browse...";
            this.btnBrowse.UseVisualStyleBackColor = true;
            // 
            // chkViewFromDB
            // 
            this.chkViewFromDB.AutoSize = true;
            this.chkViewFromDB.Location = new System.Drawing.Point(67, 77);
            this.chkViewFromDB.Name = "chkViewFromDB";
            this.chkViewFromDB.Size = new System.Drawing.Size(164, 17);
            this.chkViewFromDB.TabIndex = 22;
            this.chkViewFromDB.Text = "View PDF document from DB";
            this.chkViewFromDB.UseVisualStyleBackColor = true;
            this.chkViewFromDB.Visible = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(3, 43);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(45, 13);
            this.label3.TabIndex = 21;
            this.label3.Text = "Column:";
            this.label3.Visible = false;
            // 
            // txtInvName
            // 
            this.txtInvName.Location = new System.Drawing.Point(67, 40);
            this.txtInvName.Name = "txtInvName";
            this.txtInvName.Size = new System.Drawing.Size(32, 20);
            this.txtInvName.TabIndex = 19;
            this.txtInvName.Visible = false;
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.cbDbQ);
            this.groupBox5.Location = new System.Drawing.Point(319, 19);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(262, 54);
            this.groupBox5.TabIndex = 36;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Query Type";
            // 
            // cbDbQ
            // 
            this.cbDbQ.AutoSize = true;
            this.cbDbQ.Location = new System.Drawing.Point(16, 23);
            this.cbDbQ.Name = "cbDbQ";
            this.cbDbQ.Size = new System.Drawing.Size(103, 17);
            this.cbDbQ.TabIndex = 25;
            this.cbDbQ.Text = "Database Query";
            this.cbDbQ.UseVisualStyleBackColor = true;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.rbPdf);
            this.groupBox4.Controls.Add(this.rbWb);
            this.groupBox4.Location = new System.Drawing.Point(20, 19);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(262, 54);
            this.groupBox4.TabIndex = 35;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Viewer Type";
            // 
            // rbPdf
            // 
            this.rbPdf.AutoSize = true;
            this.rbPdf.Checked = true;
            this.rbPdf.Location = new System.Drawing.Point(16, 22);
            this.rbPdf.Name = "rbPdf";
            this.rbPdf.Size = new System.Drawing.Size(80, 17);
            this.rbPdf.TabIndex = 23;
            this.rbPdf.TabStop = true;
            this.rbPdf.Text = "PDF viewer";
            this.rbPdf.UseVisualStyleBackColor = true;
            // 
            // rbWb
            // 
            this.rbWb.AutoSize = true;
            this.rbWb.Location = new System.Drawing.Point(102, 22);
            this.rbWb.Name = "rbWb";
            this.rbWb.Size = new System.Drawing.Size(123, 17);
            this.rbWb.TabIndex = 24;
            this.rbWb.Text = "Web Browser viewer";
            this.rbWb.UseVisualStyleBackColor = true;
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(205, 435);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(221, 23);
            this.btnSave.TabIndex = 18;
            this.btnSave.Text = "Create Document View";
            this.btnSave.UseVisualStyleBackColor = true;
            // 
            // frmAd_Doc_View
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(620, 598);
            this.Controls.Add(this.tabControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmAd_Doc_View";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Add a new Document View  - RSystems FinanceTools v2";
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.panel22.ResumeLayout(false);
            this.panel22.PerformLayout();
            this.panel21.ResumeLayout(false);
            this.panel21.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.FolderBrowserDialog fbdADV;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.CheckBox ckb_isAWebFile;
        private System.Windows.Forms.TextBox txtADV_ModuleName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RichTextBox tbADV_Filepath;
        private System.Windows.Forms.ComboBox cbADV_FType;
        private System.Windows.Forms.Button btnADV_Create;
        private System.Windows.Forms.TextBox tbADV_Macro01;
        private System.Windows.Forms.Label lbADV_Macro01;
        private System.Windows.Forms.Label lbADV_Macro;
        private System.Windows.Forms.Label lbADV_FType;
        private System.Windows.Forms.TextBox tbADV_File;
        private System.Windows.Forms.Label lbADV_FName;
        private System.Windows.Forms.CheckBox chkADV_UseRef;
        private System.Windows.Forms.Button btnADV_Folder;
        private System.Windows.Forms.Label lbADV_Folder;
        private System.Windows.Forms.TextBox tbADV_Prefix;
        private System.Windows.Forms.Label lbADV_Prefix;
        private System.Windows.Forms.Label lbADV_Type;
        private System.Windows.Forms.TextBox tbADV_Type;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.ErrorProvider errorProvider1;
        private System.Windows.Forms.FolderBrowserDialog fbdAd_UpdateFolder;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Panel panel22;
        private System.Windows.Forms.Label label24;
        private System.Windows.Forms.TextBox tbColumn2;
        private System.Windows.Forms.TextBox tbConnectStr;
        private System.Windows.Forms.Label lbConnectStr;
        private System.Windows.Forms.TextBox tbUserID;
        private System.Windows.Forms.Label lbUserID;
        private System.Windows.Forms.TextBox tbPassword;
        private System.Windows.Forms.Label lbSQL;
        private System.Windows.Forms.Label lbPassword;
        public System.Windows.Forms.TextBox tbSQL;
        private System.Windows.Forms.Panel panel21;
        private System.Windows.Forms.Label label4;
        public System.Windows.Forms.TextBox txtPath;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.CheckBox chkViewFromDB;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtInvName;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.CheckBox cbDbQ;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.RadioButton rbPdf;
        private System.Windows.Forms.RadioButton rbWb;
        private System.Windows.Forms.Button btnSave;
    }
}