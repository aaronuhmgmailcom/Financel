namespace ExcelAddIn3
{
    partial class CreateSelectionList
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CreateSelectionList));
            this.cbFolders = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.cbColumnName = new System.Windows.Forms.ComboBox();
            this.cbColumnName2 = new System.Windows.Forms.ComboBox();
            this.lblColumnName2 = new System.Windows.Forms.Label();
            this.cbColumnName3 = new System.Windows.Forms.ComboBox();
            this.lblColumnName3 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txtFilter = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.cbOperator = new System.Windows.Forms.ComboBox();
            this.cbOutPut = new System.Windows.Forms.ComboBox();
            this.cbOutPut2 = new System.Windows.Forms.ComboBox();
            this.cbOperator2 = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.txtFilter2 = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.cbOutPut3 = new System.Windows.Forms.ComboBox();
            this.cbOperator3 = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.txtFilter3 = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // cbFolders
            // 
            this.cbFolders.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbFolders.FormattingEnabled = true;
            this.cbFolders.Location = new System.Drawing.Point(91, 48);
            this.cbFolders.Name = "cbFolders";
            this.cbFolders.Size = new System.Drawing.Size(172, 21);
            this.cbFolders.TabIndex = 10;
            this.cbFolders.SelectedIndexChanged += new System.EventHandler(this.cbFolders_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(17, 49);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(68, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "Table Name:";
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(800, 183);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 8;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(719, 183);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 7;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 85);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(76, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Column Name:";
            // 
            // cbColumnName
            // 
            this.cbColumnName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbColumnName.FormattingEnabled = true;
            this.cbColumnName.Location = new System.Drawing.Point(91, 81);
            this.cbColumnName.Name = "cbColumnName";
            this.cbColumnName.Size = new System.Drawing.Size(233, 21);
            this.cbColumnName.TabIndex = 11;
            // 
            // cbColumnName2
            // 
            this.cbColumnName2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbColumnName2.FormattingEnabled = true;
            this.cbColumnName2.Location = new System.Drawing.Point(91, 115);
            this.cbColumnName2.Name = "cbColumnName2";
            this.cbColumnName2.Size = new System.Drawing.Size(233, 21);
            this.cbColumnName2.TabIndex = 17;
            // 
            // lblColumnName2
            // 
            this.lblColumnName2.AutoSize = true;
            this.lblColumnName2.Location = new System.Drawing.Point(9, 117);
            this.lblColumnName2.Name = "lblColumnName2";
            this.lblColumnName2.Size = new System.Drawing.Size(76, 13);
            this.lblColumnName2.TabIndex = 16;
            this.lblColumnName2.Text = "Column Name:";
            // 
            // cbColumnName3
            // 
            this.cbColumnName3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbColumnName3.FormattingEnabled = true;
            this.cbColumnName3.Location = new System.Drawing.Point(91, 147);
            this.cbColumnName3.Name = "cbColumnName3";
            this.cbColumnName3.Size = new System.Drawing.Size(233, 21);
            this.cbColumnName3.TabIndex = 19;
            // 
            // lblColumnName3
            // 
            this.lblColumnName3.AutoSize = true;
            this.lblColumnName3.Location = new System.Drawing.Point(9, 151);
            this.lblColumnName3.Name = "lblColumnName3";
            this.lblColumnName3.Size = new System.Drawing.Size(76, 13);
            this.lblColumnName3.TabIndex = 18;
            this.lblColumnName3.Text = "Column Name:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(517, 84);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(32, 13);
            this.label3.TabIndex = 14;
            this.label3.Text = "Filter:";
            // 
            // txtFilter
            // 
            this.txtFilter.Location = new System.Drawing.Point(555, 82);
            this.txtFilter.Name = "txtFilter";
            this.txtFilter.Size = new System.Drawing.Size(187, 20);
            this.txtFilter.TabIndex = 15;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(330, 84);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(51, 13);
            this.label4.TabIndex = 20;
            this.label4.Text = "Operator:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(752, 85);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(51, 13);
            this.label5.TabIndex = 21;
            this.label5.Text = "IsOutPut:";
            // 
            // cbOperator
            // 
            this.cbOperator.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbOperator.FormattingEnabled = true;
            this.cbOperator.Location = new System.Drawing.Point(387, 81);
            this.cbOperator.Name = "cbOperator";
            this.cbOperator.Size = new System.Drawing.Size(113, 21);
            this.cbOperator.TabIndex = 22;
            // 
            // cbOutPut
            // 
            this.cbOutPut.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbOutPut.FormattingEnabled = true;
            this.cbOutPut.Location = new System.Drawing.Point(809, 81);
            this.cbOutPut.Name = "cbOutPut";
            this.cbOutPut.Size = new System.Drawing.Size(55, 21);
            this.cbOutPut.TabIndex = 23;
            this.cbOutPut.SelectedIndexChanged += new System.EventHandler(this.cbOutPut_SelectedIndexChanged);
            // 
            // cbOutPut2
            // 
            this.cbOutPut2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbOutPut2.FormattingEnabled = true;
            this.cbOutPut2.Location = new System.Drawing.Point(809, 114);
            this.cbOutPut2.Name = "cbOutPut2";
            this.cbOutPut2.Size = new System.Drawing.Size(55, 21);
            this.cbOutPut2.TabIndex = 29;
            this.cbOutPut2.SelectedIndexChanged += new System.EventHandler(this.cbOutPut2_SelectedIndexChanged);
            // 
            // cbOperator2
            // 
            this.cbOperator2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbOperator2.FormattingEnabled = true;
            this.cbOperator2.Location = new System.Drawing.Point(387, 114);
            this.cbOperator2.Name = "cbOperator2";
            this.cbOperator2.Size = new System.Drawing.Size(113, 21);
            this.cbOperator2.TabIndex = 28;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(752, 118);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(51, 13);
            this.label6.TabIndex = 27;
            this.label6.Text = "IsOutPut:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(330, 117);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(51, 13);
            this.label7.TabIndex = 26;
            this.label7.Text = "Operator:";
            // 
            // txtFilter2
            // 
            this.txtFilter2.Location = new System.Drawing.Point(555, 115);
            this.txtFilter2.Name = "txtFilter2";
            this.txtFilter2.Size = new System.Drawing.Size(187, 20);
            this.txtFilter2.TabIndex = 25;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(517, 117);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(32, 13);
            this.label8.TabIndex = 24;
            this.label8.Text = "Filter:";
            // 
            // cbOutPut3
            // 
            this.cbOutPut3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbOutPut3.FormattingEnabled = true;
            this.cbOutPut3.Location = new System.Drawing.Point(809, 147);
            this.cbOutPut3.Name = "cbOutPut3";
            this.cbOutPut3.Size = new System.Drawing.Size(55, 21);
            this.cbOutPut3.TabIndex = 35;
            this.cbOutPut3.SelectedIndexChanged += new System.EventHandler(this.cbOutPut3_SelectedIndexChanged);
            // 
            // cbOperator3
            // 
            this.cbOperator3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbOperator3.FormattingEnabled = true;
            this.cbOperator3.Location = new System.Drawing.Point(387, 147);
            this.cbOperator3.Name = "cbOperator3";
            this.cbOperator3.Size = new System.Drawing.Size(113, 21);
            this.cbOperator3.TabIndex = 34;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(752, 151);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(51, 13);
            this.label9.TabIndex = 33;
            this.label9.Text = "IsOutPut:";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(330, 150);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(51, 13);
            this.label10.TabIndex = 32;
            this.label10.Text = "Operator:";
            // 
            // txtFilter3
            // 
            this.txtFilter3.Location = new System.Drawing.Point(555, 148);
            this.txtFilter3.Name = "txtFilter3";
            this.txtFilter3.Size = new System.Drawing.Size(187, 20);
            this.txtFilter3.TabIndex = 31;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(517, 150);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(32, 13);
            this.label11.TabIndex = 30;
            this.label11.Text = "Filter:";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.groupBox1.Location = new System.Drawing.Point(0, 212);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(887, 351);
            this.groupBox1.TabIndex = 38;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Specifying Query Selection Criteria:";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.richTextBox1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 16);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(881, 332);
            this.panel1.TabIndex = 0;
            // 
            // richTextBox1
            // 
            this.richTextBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richTextBox1.Location = new System.Drawing.Point(0, 0);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(881, 332);
            this.richTextBox1.TabIndex = 0;
            this.richTextBox1.Text = resources.GetString("richTextBox1.Text");
            // 
            // CreateSelectionList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(887, 563);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.cbOutPut3);
            this.Controls.Add(this.cbOperator3);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.txtFilter3);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.cbOutPut2);
            this.Controls.Add(this.cbOperator2);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.txtFilter2);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.cbOutPut);
            this.Controls.Add(this.cbOperator);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.cbColumnName3);
            this.Controls.Add(this.lblColumnName3);
            this.Controls.Add(this.cbColumnName2);
            this.Controls.Add(this.lblColumnName2);
            this.Controls.Add(this.txtFilter);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.cbColumnName);
            this.Controls.Add(this.cbFolders);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "CreateSelectionList";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "CreateSelectionList - RSystems FinanceTools ";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.CreateSelectionList_FormClosing);
            this.groupBox1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cbFolders;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbColumnName;
        private System.Windows.Forms.ComboBox cbColumnName2;
        private System.Windows.Forms.Label lblColumnName2;
        private System.Windows.Forms.ComboBox cbColumnName3;
        private System.Windows.Forms.Label lblColumnName3;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtFilter;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cbOperator;
        private System.Windows.Forms.ComboBox cbOutPut;
        private System.Windows.Forms.ComboBox cbOutPut2;
        private System.Windows.Forms.ComboBox cbOperator2;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtFilter2;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox cbOutPut3;
        private System.Windows.Forms.ComboBox cbOperator3;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox txtFilter3;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.RichTextBox richTextBox1;
    }
}