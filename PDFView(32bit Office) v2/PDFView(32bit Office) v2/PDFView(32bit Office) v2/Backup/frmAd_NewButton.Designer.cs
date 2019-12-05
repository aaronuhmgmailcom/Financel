namespace ExcelAddIn4
{
    partial class frmAd_NewButton
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmAd_NewButton));
            this.cbFunctionType = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.tbMacroName = new System.Windows.Forms.TextBox();
            this.lbADV_FName = new System.Windows.Forms.Label();
            this.btnCNB_Create = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.cbMsg = new System.Windows.Forms.CheckBox();
            this.lblRef = new System.Windows.Forms.Label();
            this.cbReference = new System.Windows.Forms.ComboBox();
            this.chkStop = new System.Windows.Forms.CheckBox();
            this.btnAddProcMacro = new System.Windows.Forms.Button();
            this.btnDown = new System.Windows.Forms.Button();
            this.btnUp = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.errorProvider1 = new System.Windows.Forms.ErrorProvider(this.components);
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.deleteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.panel2 = new System.Windows.Forms.Panel();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.btnIcon = new System.Windows.Forms.Button();
            this.btnGroup = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnDelete = new System.Windows.Forms.Button();
            this.tbButtonName = new System.Windows.Forms.ComboBox();
            this.cbTemplateName = new System.Windows.Forms.ComboBox();
            this.lbADV_FType = new System.Windows.Forms.Label();
            this.lbADV_Type = new System.Windows.Forms.Label();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.clbGroup = new System.Windows.Forms.CheckedListBox();
            this.panel1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // cbFunctionType
            // 
            this.cbFunctionType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbFunctionType.FormattingEnabled = true;
            this.cbFunctionType.Items.AddRange(new object[] {
            "Output Process",
            "Output Macro",
            "Save",
            "Save PDF"});
            this.cbFunctionType.Location = new System.Drawing.Point(129, 30);
            this.cbFunctionType.Name = "cbFunctionType";
            this.cbFunctionType.Size = new System.Drawing.Size(121, 21);
            this.cbFunctionType.TabIndex = 21;
            this.cbFunctionType.SelectedIndexChanged += new System.EventHandler(this.cbFunctionType_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(33, 33);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(69, 13);
            this.label1.TabIndex = 20;
            this.label1.Text = "Output Type:";
            // 
            // tbMacroName
            // 
            this.tbMacroName.Location = new System.Drawing.Point(129, 67);
            this.tbMacroName.MaxLength = 50;
            this.tbMacroName.Name = "tbMacroName";
            this.tbMacroName.Size = new System.Drawing.Size(194, 20);
            this.tbMacroName.TabIndex = 23;
            // 
            // lbADV_FName
            // 
            this.lbADV_FName.AutoSize = true;
            this.lbADV_FName.Location = new System.Drawing.Point(29, 70);
            this.lbADV_FName.Name = "lbADV_FName";
            this.lbADV_FName.Size = new System.Drawing.Size(73, 13);
            this.lbADV_FName.TabIndex = 22;
            this.lbADV_FName.Text = "Name(Profile):";
            // 
            // btnCNB_Create
            // 
            this.btnCNB_Create.Location = new System.Drawing.Point(148, 449);
            this.btnCNB_Create.Name = "btnCNB_Create";
            this.btnCNB_Create.Size = new System.Drawing.Size(221, 32);
            this.btnCNB_Create.TabIndex = 24;
            this.btnCNB_Create.Text = "Create New Action";
            this.btnCNB_Create.UseVisualStyleBackColor = true;
            this.btnCNB_Create.Click += new System.EventHandler(this.btnCNB_Create_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(458, 313);
            this.panel1.TabIndex = 25;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.cbMsg);
            this.groupBox1.Controls.Add(this.lblRef);
            this.groupBox1.Controls.Add(this.cbReference);
            this.groupBox1.Controls.Add(this.chkStop);
            this.groupBox1.Controls.Add(this.btnAddProcMacro);
            this.groupBox1.Controls.Add(this.btnDown);
            this.groupBox1.Controls.Add(this.btnUp);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.cbFunctionType);
            this.groupBox1.Controls.Add(this.tbMacroName);
            this.groupBox1.Controls.Add(this.listBox1);
            this.groupBox1.Controls.Add(this.lbADV_FName);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(458, 313);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Output Mode:";
            // 
            // cbMsg
            // 
            this.cbMsg.AutoSize = true;
            this.cbMsg.Location = new System.Drawing.Point(372, 32);
            this.cbMsg.Name = "cbMsg";
            this.cbMsg.Size = new System.Drawing.Size(82, 17);
            this.cbMsg.TabIndex = 29;
            this.cbMsg.Text = "Show Msg?";
            this.cbMsg.UseVisualStyleBackColor = true;
            // 
            // lblRef
            // 
            this.lblRef.AutoSize = true;
            this.lblRef.Location = new System.Drawing.Point(42, 106);
            this.lblRef.Name = "lblRef";
            this.lblRef.Size = new System.Drawing.Size(60, 13);
            this.lblRef.TabIndex = 27;
            this.lblRef.Text = "Reference:";
            // 
            // cbReference
            // 
            this.cbReference.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbReference.FormattingEnabled = true;
            this.cbReference.Location = new System.Drawing.Point(129, 103);
            this.cbReference.Name = "cbReference";
            this.cbReference.Size = new System.Drawing.Size(194, 21);
            this.cbReference.TabIndex = 28;
            // 
            // chkStop
            // 
            this.chkStop.AutoSize = true;
            this.chkStop.Location = new System.Drawing.Point(266, 32);
            this.chkStop.Name = "chkStop";
            this.chkStop.Size = new System.Drawing.Size(93, 17);
            this.chkStop.TabIndex = 26;
            this.chkStop.Text = "Stop on error?";
            this.chkStop.UseVisualStyleBackColor = true;
            // 
            // btnAddProcMacro
            // 
            this.btnAddProcMacro.Location = new System.Drawing.Point(129, 139);
            this.btnAddProcMacro.Name = "btnAddProcMacro";
            this.btnAddProcMacro.Size = new System.Drawing.Size(194, 23);
            this.btnAddProcMacro.TabIndex = 25;
            this.btnAddProcMacro.Text = "Add Output/Save";
            this.btnAddProcMacro.UseVisualStyleBackColor = true;
            this.btnAddProcMacro.Click += new System.EventHandler(this.btnAddProcMacro_Click);
            // 
            // btnDown
            // 
            this.btnDown.Location = new System.Drawing.Point(322, 262);
            this.btnDown.Name = "btnDown";
            this.btnDown.Size = new System.Drawing.Size(33, 36);
            this.btnDown.TabIndex = 2;
            this.btnDown.Text = " ↓";
            this.btnDown.UseVisualStyleBackColor = true;
            this.btnDown.Click += new System.EventHandler(this.btnDown_Click);
            // 
            // btnUp
            // 
            this.btnUp.Location = new System.Drawing.Point(322, 176);
            this.btnUp.Name = "btnUp";
            this.btnUp.Size = new System.Drawing.Size(33, 35);
            this.btnUp.TabIndex = 1;
            this.btnUp.Text = " ↑";
            this.btnUp.UseVisualStyleBackColor = true;
            this.btnUp.Click += new System.EventHandler(this.btnUp_Click);
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(129, 177);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(183, 121);
            this.listBox1.TabIndex = 0;
            this.listBox1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.listBox1_MouseUp);
            // 
            // errorProvider1
            // 
            this.errorProvider1.ContainerControl = this;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.deleteToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(108, 26);
            // 
            // deleteToolStripMenuItem
            // 
            this.deleteToolStripMenuItem.Name = "deleteToolStripMenuItem";
            this.deleteToolStripMenuItem.Size = new System.Drawing.Size(107, 22);
            this.deleteToolStripMenuItem.Text = "Delete";
            this.deleteToolStripMenuItem.Click += new System.EventHandler(this.deleteToolStripMenuItem_Click);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.groupBox2);
            this.panel2.Location = new System.Drawing.Point(12, 363);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(472, 72);
            this.panel2.TabIndex = 26;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.textBox1);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox2.Location = new System.Drawing.Point(0, 0);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(472, 72);
            this.groupBox2.TabIndex = 0;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Description:";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(7, 19);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(458, 46);
            this.textBox1.TabIndex = 0;
            this.textBox1.Text = "Display in action\'s SuperTip ,Click here!";
            this.textBox1.Click += new System.EventHandler(this.textBox1_Click);
            this.textBox1.Leave += new System.EventHandler(this.textBox1_Leave);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.btnIcon);
            this.groupBox3.Controls.Add(this.btnGroup);
            this.groupBox3.Controls.Add(this.label3);
            this.groupBox3.Controls.Add(this.label2);
            this.groupBox3.Controls.Add(this.btnDelete);
            this.groupBox3.Controls.Add(this.tbButtonName);
            this.groupBox3.Controls.Add(this.cbTemplateName);
            this.groupBox3.Controls.Add(this.lbADV_FType);
            this.groupBox3.Controls.Add(this.lbADV_Type);
            this.groupBox3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox3.Location = new System.Drawing.Point(3, 3);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(458, 313);
            this.groupBox3.TabIndex = 27;
            this.groupBox3.TabStop = false;
            // 
            // btnIcon
            // 
            this.btnIcon.Location = new System.Drawing.Point(106, 120);
            this.btnIcon.Name = "btnIcon";
            this.btnIcon.Size = new System.Drawing.Size(64, 64);
            this.btnIcon.TabIndex = 41;
            this.btnIcon.Tag = "Click here to search a perfect icon!";
            this.btnIcon.Text = "Choose an Icon";
            this.btnIcon.UseVisualStyleBackColor = true;
            this.btnIcon.Click += new System.EventHandler(this.btnIcon_Click);
            // 
            // btnGroup
            // 
            this.btnGroup.Location = new System.Drawing.Point(106, 92);
            this.btnGroup.Name = "btnGroup";
            this.btnGroup.Size = new System.Drawing.Size(194, 21);
            this.btnGroup.TabIndex = 40;
            this.btnGroup.Tag = "Click here to config your groups!";
            this.btnGroup.Text = "Settings...";
            this.btnGroup.UseVisualStyleBackColor = true;
            this.btnGroup.Click += new System.EventHandler(this.btnGroup_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(55, 146);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(31, 13);
            this.label3.TabIndex = 39;
            this.label3.Text = "Icon:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(33, 96);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(52, 13);
            this.label2.TabIndex = 38;
            this.label2.Text = "Category:";
            // 
            // btnDelete
            // 
            this.btnDelete.Location = new System.Drawing.Point(311, 57);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(75, 23);
            this.btnDelete.TabIndex = 37;
            this.btnDelete.Text = "Delete";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Visible = false;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // tbButtonName
            // 
            this.tbButtonName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.tbButtonName.FormattingEnabled = true;
            this.tbButtonName.Items.AddRange(new object[] {
            "New Button"});
            this.tbButtonName.Location = new System.Drawing.Point(106, 59);
            this.tbButtonName.Name = "tbButtonName";
            this.tbButtonName.Size = new System.Drawing.Size(194, 21);
            this.tbButtonName.TabIndex = 36;
            this.tbButtonName.SelectedIndexChanged += new System.EventHandler(this.tbButtonName_SelectedIndexChanged);
            // 
            // cbTemplateName
            // 
            this.cbTemplateName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbTemplateName.FormattingEnabled = true;
            this.cbTemplateName.Location = new System.Drawing.Point(106, 26);
            this.cbTemplateName.Name = "cbTemplateName";
            this.cbTemplateName.Size = new System.Drawing.Size(194, 21);
            this.cbTemplateName.TabIndex = 35;
            this.cbTemplateName.SelectedIndexChanged += new System.EventHandler(this.cbTemplateName_SelectedIndexChanged);
            // 
            // lbADV_FType
            // 
            this.lbADV_FType.AutoSize = true;
            this.lbADV_FType.Location = new System.Drawing.Point(31, 29);
            this.lbADV_FType.Name = "lbADV_FType";
            this.lbADV_FType.Size = new System.Drawing.Size(54, 13);
            this.lbADV_FType.TabIndex = 34;
            this.lbADV_FType.Text = "Template:";
            // 
            // lbADV_Type
            // 
            this.lbADV_Type.AutoSize = true;
            this.lbADV_Type.Location = new System.Drawing.Point(14, 62);
            this.lbADV_Type.Name = "lbADV_Type";
            this.lbADV_Type.Size = new System.Drawing.Size(71, 13);
            this.lbADV_Type.TabIndex = 33;
            this.lbADV_Type.Text = "Action Name:";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(472, 345);
            this.tabControl1.TabIndex = 28;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.groupBox3);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(464, 319);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "General";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.panel1);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(464, 319);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Advanced";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.groupBox4);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(464, 319);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Security";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.clbGroup);
            this.groupBox4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox4.Location = new System.Drawing.Point(3, 3);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(458, 313);
            this.groupBox4.TabIndex = 0;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Groups for action";
            // 
            // clbGroup
            // 
            this.clbGroup.CheckOnClick = true;
            this.clbGroup.Dock = System.Windows.Forms.DockStyle.Fill;
            this.clbGroup.FormattingEnabled = true;
            this.clbGroup.HorizontalScrollbar = true;
            this.clbGroup.Location = new System.Drawing.Point(3, 16);
            this.clbGroup.Name = "clbGroup";
            this.clbGroup.Size = new System.Drawing.Size(452, 294);
            this.clbGroup.TabIndex = 1;
            // 
            // frmAd_NewButton
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(504, 493);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.btnCNB_Create);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmAd_NewButton";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Create New Action  - RSystems FinanceTools v2";
            this.panel1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage3.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ComboBox cbFunctionType;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbMacroName;
        private System.Windows.Forms.Label lbADV_FName;
        private System.Windows.Forms.Button btnCNB_Create;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.Button btnDown;
        private System.Windows.Forms.Button btnUp;
        private System.Windows.Forms.ErrorProvider errorProvider1;
        private System.Windows.Forms.Button btnAddProcMacro;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem deleteToolStripMenuItem;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.CheckBox chkStop;
        private System.Windows.Forms.Label lblRef;
        private System.Windows.Forms.ComboBox cbReference;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button btnIcon;
        private System.Windows.Forms.Button btnGroup;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.ComboBox tbButtonName;
        private System.Windows.Forms.ComboBox cbTemplateName;
        private System.Windows.Forms.Label lbADV_FType;
        private System.Windows.Forms.Label lbADV_Type;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.CheckedListBox clbGroup;
        private System.Windows.Forms.CheckBox cbMsg;
    }
}