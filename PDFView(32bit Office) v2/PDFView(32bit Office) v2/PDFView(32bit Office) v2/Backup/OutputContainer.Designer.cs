using Utility.DataComponent;
using ExcelAddIn4.Common;
namespace ExcelAddIn4
{
    partial class OutputContainer
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.panel4 = new System.Windows.Forms.Panel();
            this.panel5 = new System.Windows.Forms.Panel();
            this.panel6 = new System.Windows.Forms.Panel();
            this.btnSetMax = new System.Windows.Forms.Button();
            this.btnTestJournal = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.panel10 = new System.Windows.Forms.Panel();
            this.panel7 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.button2 = new System.Windows.Forms.Button();
            this.tabPage7 = new System.Windows.Forms.TabPage();
            this.panel9 = new System.Windows.Forms.Panel();
            this.panel17 = new System.Windows.Forms.Panel();
            this.panel16 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.cbXMLOrText = new System.Windows.Forms.ComboBox();
            this.cbItems = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnTestCTF = new System.Windows.Forms.Button();
            this.CTF_btnSave = new System.Windows.Forms.Button();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.updateToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            CopyStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            PasteStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            InsertStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            RemoveStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dataGridViewColumnHeaderEditor1 = new Utility.DataComponent.DataGridViewColumnHeaderEditor(this.components);
            this.fbdAd_UpdateFolder = new System.Windows.Forms.FolderBrowserDialog();
            this.panel1.SuspendLayout();
            this.panel3.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.panel4.SuspendLayout();
            this.panel6.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.panel10.SuspendLayout();
            this.panel2.SuspendLayout();
            this.tabPage7.SuspendLayout();
            this.panel9.SuspendLayout();
            this.panel16.SuspendLayout();
            this.contextMenuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewColumnHeaderEditor1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.panel3);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1207, 464);
            this.panel1.TabIndex = 0;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.tabControl1);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel3.Location = new System.Drawing.Point(0, 0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(1207, 464);
            this.panel3.TabIndex = 1;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Controls.Add(this.tabPage7);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1207, 464);
            this.tabControl1.TabIndex = 1;
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.panel4);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1199, 438);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Journal";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.panel5);
            this.panel4.Controls.Add(this.panel6);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel4.Location = new System.Drawing.Point(3, 3);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(1193, 432);
            this.panel4.TabIndex = 5;
            // 
            // panel5
            // 
            this.panel5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel5.Location = new System.Drawing.Point(0, 35);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(1193, 397);
            this.panel5.TabIndex = 1;
            // 
            // panel6
            // 
            this.panel6.Controls.Add(this.btnSetMax);
            this.panel6.Controls.Add(this.btnTestJournal);
            this.panel6.Controls.Add(this.button3);
            this.panel6.Controls.Add(this.button1);
            this.panel6.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel6.Location = new System.Drawing.Point(0, 0);
            this.panel6.Name = "panel6";
            this.panel6.Size = new System.Drawing.Size(1193, 35);
            this.panel6.TabIndex = 0;
            // 
            // btnSetMax
            // 
            this.btnSetMax.Location = new System.Drawing.Point(626, 2);
            this.btnSetMax.Name = "btnSetMax";
            this.btnSetMax.Size = new System.Drawing.Size(75, 31);
            this.btnSetMax.TabIndex = 23;
            this.btnSetMax.Text = "Set Max";
            this.btnSetMax.UseVisualStyleBackColor = true;
            this.btnSetMax.Click += new System.EventHandler(this.btnSetMax_Click);
            // 
            // btnTestJournal
            // 
            this.btnTestJournal.Location = new System.Drawing.Point(707, 2);
            this.btnTestJournal.Name = "btnTestJournal";
            this.btnTestJournal.Size = new System.Drawing.Size(97, 31);
            this.btnTestJournal.TabIndex = 22;
            this.btnTestJournal.Text = "Test Journal";
            this.btnTestJournal.UseVisualStyleBackColor = true;
            this.btnTestJournal.Click += new System.EventHandler(this.btnTestJournal_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(545, 2);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 31);
            this.button3.TabIndex = 20;
            this.button3.Text = "Add Fields";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(464, 2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 31);
            this.button1.TabIndex = 19;
            this.button1.Text = "Save";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.panel10);
            this.tabPage4.Location = new System.Drawing.Point(4, 22);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Size = new System.Drawing.Size(1199, 438);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "Save Options";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // panel10
            // 
            this.panel10.Controls.Add(this.panel7);
            this.panel10.Controls.Add(this.panel2);
            this.panel10.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel10.Location = new System.Drawing.Point(0, 0);
            this.panel10.Name = "panel10";
            this.panel10.Size = new System.Drawing.Size(1199, 438);
            this.panel10.TabIndex = 0;
            // 
            // panel7
            // 
            this.panel7.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel7.Location = new System.Drawing.Point(0, 35);
            this.panel7.Name = "panel7";
            this.panel7.Size = new System.Drawing.Size(1199, 403);
            this.panel7.TabIndex = 1;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.button2);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1199, 35);
            this.panel2.TabIndex = 0;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(464, 2);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 31);
            this.button2.TabIndex = 20;
            this.button2.Text = "Save";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.btnSaveCriteria_Click);
            // 
            // tabPage7
            // 
            this.tabPage7.Controls.Add(this.panel9);
            this.tabPage7.Location = new System.Drawing.Point(4, 22);
            this.tabPage7.Name = "tabPage7";
            this.tabPage7.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage7.Size = new System.Drawing.Size(1199, 438);
            this.tabPage7.TabIndex = 7;
            this.tabPage7.Text = "XML/Text File";
            this.tabPage7.UseVisualStyleBackColor = true;
            // 
            // panel9
            // 
            this.panel9.Controls.Add(this.panel17);
            this.panel9.Controls.Add(this.panel16);
            this.panel9.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel9.Location = new System.Drawing.Point(3, 3);
            this.panel9.Name = "panel9";
            this.panel9.Size = new System.Drawing.Size(1193, 432);
            this.panel9.TabIndex = 0;
            // 
            // panel17
            // 
            this.panel17.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel17.Location = new System.Drawing.Point(0, 35);
            this.panel17.Name = "panel17";
            this.panel17.Size = new System.Drawing.Size(1193, 397);
            this.panel17.TabIndex = 4;
            // 
            // panel16
            // 
            this.panel16.Controls.Add(this.label2);
            this.panel16.Controls.Add(this.cbXMLOrText);
            this.panel16.Controls.Add(this.cbItems);
            this.panel16.Controls.Add(this.label1);
            this.panel16.Controls.Add(this.btnTestCTF);
            this.panel16.Controls.Add(this.CTF_btnSave);
            this.panel16.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel16.Location = new System.Drawing.Point(0, 0);
            this.panel16.Name = "panel16";
            this.panel16.Size = new System.Drawing.Size(1193, 35);
            this.panel16.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(169, 11);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(167, 13);
            this.label2.TabIndex = 35;
            this.label2.Text = "Choose a XML or Text File Profile:";
            // 
            // cbXMLOrText
            // 
            this.cbXMLOrText.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbXMLOrText.FormattingEnabled = true;
            this.cbXMLOrText.Items.AddRange(new object[] {
            "XML",
            "Text File"});
            this.cbXMLOrText.Location = new System.Drawing.Point(55, 8);
            this.cbXMLOrText.Name = "cbXMLOrText";
            this.cbXMLOrText.Size = new System.Drawing.Size(85, 21);
            this.cbXMLOrText.TabIndex = 34;
            this.cbXMLOrText.SelectedIndexChanged += new System.EventHandler(this.cbXMLOrText_SelectedIndexChanged);
            // 
            // cbItems
            // 
            this.cbItems.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbItems.FormattingEnabled = true;
            this.cbItems.Location = new System.Drawing.Point(388, 8);
            this.cbItems.Name = "cbItems";
            this.cbItems.Size = new System.Drawing.Size(292, 21);
            this.cbItems.TabIndex = 33;
            this.cbItems.SelectedIndexChanged += new System.EventHandler(this.cbItems_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(34, 13);
            this.label1.TabIndex = 32;
            this.label1.Text = "Type:";
            // 
            // btnTestCTF
            // 
            this.btnTestCTF.Location = new System.Drawing.Point(779, 2);
            this.btnTestCTF.Name = "btnTestCTF";
            this.btnTestCTF.Size = new System.Drawing.Size(75, 31);
            this.btnTestCTF.TabIndex = 31;
            this.btnTestCTF.Text = "Test";
            this.btnTestCTF.UseVisualStyleBackColor = true;
            this.btnTestCTF.Click += new System.EventHandler(this.btnTestCTF_Click);
            // 
            // CTF_btnSave
            // 
            this.CTF_btnSave.Location = new System.Drawing.Point(698, 2);
            this.CTF_btnSave.Name = "CTF_btnSave";
            this.CTF_btnSave.Size = new System.Drawing.Size(75, 31);
            this.CTF_btnSave.TabIndex = 22;
            this.CTF_btnSave.Text = "Save";
            this.CTF_btnSave.UseVisualStyleBackColor = true;
            this.CTF_btnSave.Click += new System.EventHandler(this.CTF_btnSave_Click);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1,
            this.updateToolStripMenuItem,
            this.CopyStripMenuItem,
            this.PasteStripMenuItem,
            this.InsertStripMenuItem,
            this.RemoveStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(226, 48);
            this.contextMenuStrip1.Text = "Consolidate by";
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(225, 22);
            this.toolStripMenuItem1.Text = "Consolidate/Unconsolidated";
            this.toolStripMenuItem1.Click += new System.EventHandler(this.toolStripMenuItem1_Click);
            // 
            // updateToolStripMenuItem
            // 
            this.updateToolStripMenuItem.Name = "updateToolStripMenuItem";
            this.updateToolStripMenuItem.Size = new System.Drawing.Size(225, 22);
            this.updateToolStripMenuItem.Text = "Update/Remove Update";
            this.updateToolStripMenuItem.Visible = false;
            this.updateToolStripMenuItem.Click += new System.EventHandler(this.updateToolStripMenuItem_Click);
            // 
            // CopyStripMenuItem
            // 
            this.CopyStripMenuItem.Name = "CopyStripMenuItem";
            this.CopyStripMenuItem.Size = new System.Drawing.Size(125, 22);
            this.CopyStripMenuItem.Text = "Copy";
            this.CopyStripMenuItem.Visible = false;
            this.CopyStripMenuItem.Click += new System.EventHandler(this.CopyStripMenuItem_Click);
            // 
            // PasteStripMenuItem
            // 
            this.PasteStripMenuItem.Name = "PasteStripMenuItem";
            this.PasteStripMenuItem.Size = new System.Drawing.Size(125, 22);
            this.PasteStripMenuItem.Text = "Paste";
            this.PasteStripMenuItem.Visible = false;
            this.PasteStripMenuItem.Click += new System.EventHandler(this.PasteStripMenuItem_Click);
            // 
            // InsertStripMenuItem
            // 
            this.InsertStripMenuItem.Name = "InsertStripMenuItem";
            this.InsertStripMenuItem.Size = new System.Drawing.Size(125, 22);
            this.InsertStripMenuItem.Text = "Insert";
            this.InsertStripMenuItem.Visible = false;
            this.InsertStripMenuItem.Click += new System.EventHandler(this.InsertStripMenuItem_Click);
            // 
            // RemoveStripMenuItem
            // 
            this.RemoveStripMenuItem.Name = "RemoveStripMenuItem";
            this.RemoveStripMenuItem.Size = new System.Drawing.Size(125, 22);
            this.RemoveStripMenuItem.Text = "Remove";
            this.RemoveStripMenuItem.Visible = false;
            this.RemoveStripMenuItem.Click += new System.EventHandler(this.RemoveStripMenuItem_Click);
            // 
            // dataGridViewColumnHeaderEditor1
            // 
            this.dataGridViewColumnHeaderEditor1.TargetControl = null;
            // 
            // OutputContainer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panel1);
            this.Name = "OutputContainer";
            this.Size = new System.Drawing.Size(1207, 464);
            this.panel1.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            this.panel6.ResumeLayout(false);
            this.tabPage4.ResumeLayout(false);
            this.panel10.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.tabPage7.ResumeLayout(false);
            this.panel9.ResumeLayout(false);
            this.panel16.ResumeLayout(false);
            this.panel16.PerformLayout();
            this.contextMenuStrip1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewColumnHeaderEditor1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Panel panel6;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.Panel panel10;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.FolderBrowserDialog fbdAd_UpdateFolder;
        private System.Windows.Forms.Button btnTestJournal;
        private System.Windows.Forms.TabPage tabPage7;
        private System.Windows.Forms.Panel panel9;
        private System.Windows.Forms.Panel panel17;
        private System.Windows.Forms.Panel panel16;
        private System.Windows.Forms.Button CTF_btnSave;
        private DataGridViewColumnHeaderEditor dataGridViewColumnHeaderEditor1;
        private System.Windows.Forms.Button btnTestCTF;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Panel panel7;
        private System.Windows.Forms.Label label1;
        public System.Windows.Forms.ComboBox cbItems;
        public System.Windows.Forms.ComboBox cbXMLOrText;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ToolStripMenuItem updateToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem CopyStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem PasteStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem InsertStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem RemoveStripMenuItem;
        private System.Windows.Forms.Button btnSetMax;
    }
}
