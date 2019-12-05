namespace ExcelAddIn4
{
    partial class FileNameForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FileNameForm));
            this.panel1 = new System.Windows.Forms.Panel();
            this.cbTmpName = new System.Windows.Forms.ComboBox();
            this.btnDelTemp = new System.Windows.Forms.Button();
            this.cbFolders = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.cbTmpName);
            this.panel1.Controls.Add(this.btnDelTemp);
            this.panel1.Controls.Add(this.cbFolders);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.btnCancel);
            this.panel1.Controls.Add(this.btnSave);
            this.panel1.Controls.Add(this.textBox1);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(445, 212);
            this.panel1.TabIndex = 0;
            // 
            // cbTmpName
            // 
            this.cbTmpName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbTmpName.FormattingEnabled = true;
            this.cbTmpName.Location = new System.Drawing.Point(158, 97);
            this.cbTmpName.Name = "cbTmpName";
            this.cbTmpName.Size = new System.Drawing.Size(172, 21);
            this.cbTmpName.TabIndex = 7;
            this.cbTmpName.SelectedIndexChanged += new System.EventHandler(this.cbTmpName_SelectedIndexChanged);
            // 
            // btnDelTemp
            // 
            this.btnDelTemp.Location = new System.Drawing.Point(358, 95);
            this.btnDelTemp.Name = "btnDelTemp";
            this.btnDelTemp.Size = new System.Drawing.Size(75, 23);
            this.btnDelTemp.TabIndex = 6;
            this.btnDelTemp.Text = "Delete";
            this.btnDelTemp.UseVisualStyleBackColor = true;
            this.btnDelTemp.Visible = false;
            this.btnDelTemp.Click += new System.EventHandler(this.btnDelTemp_Click);
            // 
            // cbFolders
            // 
            this.cbFolders.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbFolders.FormattingEnabled = true;
            this.cbFolders.Location = new System.Drawing.Point(158, 47);
            this.cbFolders.Name = "cbFolders";
            this.cbFolders.Size = new System.Drawing.Size(172, 21);
            this.cbFolders.TabIndex = 5;
            this.cbFolders.SelectedIndexChanged += new System.EventHandler(this.cbFolder_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(33, 47);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(70, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Folder Name:";
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(264, 161);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(158, 161);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 2;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(158, 98);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(172, 20);
            this.textBox1.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(30, 101);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(85, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Template Name:";
            // 
            // FileNameForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(445, 212);
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FileNameForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Save/Amend Template  - RSystems FinanceTools v2";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cbFolders;
        private System.Windows.Forms.Button btnDelTemp;
        private System.Windows.Forms.ComboBox cbTmpName;

    }
}