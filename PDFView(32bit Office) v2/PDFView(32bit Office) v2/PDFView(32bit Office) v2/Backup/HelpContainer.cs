using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using System.Configuration;

namespace ExcelAddIn4
{
    public partial class HelpContainer : UserControl
    {
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
        public HelpContainer()
        {
            InitializeComponent();
            this.toolStripButton1.Image = ExcelAddIn4.Common.OleCreateConverter.PictureDispToImage(Globals.ThisAddIn.Application.CommandBars.GetImageMso("FileSave", 23, 22));
            this.toolStripButton2.Image = ExcelAddIn4.Common.OleCreateConverter.PictureDispToImage(Globals.ThisAddIn.Application.CommandBars.GetImageMso("SharingOpenNotesFolder", 26, 24));
            this.toolStrip1.Size = new System.Drawing.Size(543, 25);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            SaveToFile();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filename"></param>
        private void SaveHelp(string filename)
        {
            if (!string.IsNullOrEmpty(filename))
            {
                SqlConnection conn = null;
                try
                {
                    conn = new
                        SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                    conn.Open();
                    SqlCommand cmd = new SqlCommand("rsTemplateHelp_Del", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                    cmd.ExecuteNonQuery();
                    SqlCommand cmd2 = new SqlCommand("rsTemplateHelp_Ins", conn);
                    cmd2.CommandType = CommandType.StoredProcedure;
                    cmd2.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                    cmd2.Parameters.Add(new SqlParameter("@helpFileData", ft.GetData(filename)));
                    cmd2.Parameters.Add(new SqlParameter("@helpFileType", Path.GetExtension(filename)));
                    cmd2.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                finally
                {
                    if (conn != null)
                    {
                        conn.Close();
                    }
                }
            }
            else
            {
                MessageBox.Show("Please choose a help file !", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = false;
            fileDialog.Title = "Please choose a help file";
            fileDialog.Filter = "Help Files|*.htm;*.html;*.shtml;*.xhtml|All Files|*.*";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                SaveHelp(fileDialog.FileName);
                string Url = ft.OpenHelpFile().Replace("\\\\", "\\");
                webBrowser1.Navigate(Url);
                webBrowser1.Document.ExecCommand("EditMode", true, "");
            }
        }
        /// <summary>
        /// 
        /// </summary>
        private void SaveToFile()
        {
            try
            {
                StreamReader sr = new StreamReader(webBrowser1.DocumentStream);
                StreamWriter sw = new StreamWriter(webBrowser1.Url.LocalPath);
                sw.Write(sr.ReadToEnd());
                sw.Flush();
                sw.Close();
                sr.Close();
            }
            catch { }
            SqlConnection conn = null;
            try
            {
                conn = new
                    SqlConnection(ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.ToString());
                conn.Open();
                SqlCommand cmd = new SqlCommand("rsTemplateHelp_Del", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                cmd.ExecuteNonQuery();
                SqlCommand cmd2 = new SqlCommand("rsTemplateHelp_Ins", conn);
                cmd2.CommandType = CommandType.StoredProcedure;
                cmd2.Parameters.Add(new SqlParameter("@TemplateID", SessionInfo.UserInfo.File_ftid));
                cmd2.Parameters.Add(new SqlParameter("@helpFileData", ft.GetData(webBrowser1.Url.LocalPath)));
                cmd2.Parameters.Add(new SqlParameter("@helpFileType", Path.GetExtension(webBrowser1.Url.LocalPath)));
                cmd2.ExecuteNonQuery();
                MessageBox.Show("Success!", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
}
