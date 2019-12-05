namespace Utility.DataComponent
{
    partial class DataGridViewColumnHeaderEditor
    {
        /// <summary>
        /// 
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 
        /// </summary>
        /// <param name="disposing"></param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region

        /// <summary>
        /// 
        /// </summary>
        private void InitializeComponent()
        {
            this.rtbTitle = new System.Windows.Forms.RichTextBox();
            // 
            // rtbTitle
            // 
            this.rtbTitle.Location = new System.Drawing.Point(0, 0);
            this.rtbTitle.Name = "rtbTitle";
            this.rtbTitle.Size = new System.Drawing.Size(100, 96);
            this.rtbTitle.TabIndex = 0;
            this.rtbTitle.Text = "";
            this.rtbTitle.Multiline = false;
            this.rtbTitle.TabStop = false;
            this.rtbTitle.Visible = false;
            this.rtbTitle.KeyDown += new System.Windows.Forms.KeyEventHandler(this.rtbTitle_KeyDown);
            this.rtbTitle.Leave += new System.EventHandler(this.rtbTitle_Leave);
        }

        #endregion

        private System.Windows.Forms.RichTextBox rtbTitle;
    }
}
