
/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<LoginFrm>   
 * Author：Peter.uhm  (yanb@shinetechchina.com)
 * Modify date：2015.10
 * Version : 2.0.0.2
 */


using System;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelAddIn4
{
    public partial class LoginFrm : Form
    {
        internal object pSender;
        public LoginFrm()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbForm_CheckedChanged(object sender, EventArgs e)
        {
            if (this.cbForm.Checked == true)
            {
                this.panel2.Visible = true;
            }
            else
            {
                this.panel2.Visible = false;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            string formUserID = this.textBox1.Text;
            string formPassword = this.textBox2.Text;
            using (RSFinanceToolsEntities db = new RSFinanceToolsEntities())
            {
                try
                {
                    var wUserID = (from FT_user in db
                                   where FT_user.FormUserID == formUserID & FT_user.FormUserPassword == formPassword
                                   select FT_user.ft_id).First();

                    var SunIP = (from FT_user in db.FinTools_Users
                                 where FT_user.FormUserID == formUserID & FT_user.FormUserPassword == formPassword
                                 select FT_user.SUNUserIP).First();

                    var SunUID = (from FT_user in db.FinTools_Users
                                  where FT_user.FormUserID == formUserID & FT_user.FormUserPassword == formPassword
                                  select FT_user.SUNUserID).First();

                    var SunUpass = (from FT_user in db.FinTools_Users
                                    where FT_user.FormUserID == formUserID & FT_user.FormUserPassword == formPassword
                                    select FT_user.SUNUserPass).First();

                    if (SessionInfo.UserInfo == null)
                    {
                        SessionInfo.UserInfo = new UserInfo();
                    }
                    SessionInfo.UserInfo.ID = wUserID.ToString();
                    SessionInfo.UserInfo.SunUserIP = SunIP;
                    SessionInfo.UserInfo.SunUserID = DEncrypt.Decrypt(SunUID);
                    SessionInfo.UserInfo.SunUserPass = DEncrypt.Decrypt(SunUpass);

                    if (pSender != null)
                    {
                        ((Ribbon2)pSender).addfolders();
                    }
                    this.Close();
                    this.Dispose();
                }
                catch
                {
                    MessageBox.Show("The user id or the password is invalid.", "Message - RSystems FinanceTools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1_Click(null, null);
            }
        }
    }
}
