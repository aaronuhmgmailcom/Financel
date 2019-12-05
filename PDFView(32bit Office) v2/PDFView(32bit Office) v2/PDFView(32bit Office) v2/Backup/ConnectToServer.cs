using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.Data.SqlClient;
using ExcelAddIn4.Common;

namespace ExcelAddIn4
{
    public partial class ConnectToServer : Form
    {
        public bool hasSaved = false;
        public ConnectToServer()
        {
            InitializeComponent();
        }
        private void textBox1_Validating(object sender, CancelEventArgs e)
        {
            if (string.IsNullOrEmpty(this.txtServerName.Text))
                errorProvider1.SetError(txtServerName, "Server name can't be empty!");
        }
        private void button1_Click(object sender, EventArgs e)
        {
            string serverName = txtServerName.Text;
            string userid = this.txtUserID.Text;
            string pass = this.txtPass.Text;
            string connString = string.Format("Data Source={0};Initial Catalog=RSDataV2;User ID={1};Password={2}", serverName, userid, pass);

            SqlConnection sqlConn = new SqlConnection(connString);
            try
            {
                sqlConn.Open();
                MessageBox.Show("Test Connection succeeded!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("The test database connection failed! Please check the server name and database name is correct!");
                LogHelper.WriteLog(typeof(ConnectToServer), ex.Message);
            }
            finally
            {
                sqlConn.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string serverName = txtServerName.Text;
            string userid = this.txtUserID.Text;
            string pass = this.txtPass.Text;
            string connString = string.Format("Data Source={0};Initial Catalog=RSDataV2;User ID={1};Password={2}", serverName, userid, pass);

            SqlConnection sqlConn = new SqlConnection(connString);
            try
            {
                sqlConn.Open();

                string str = ConfigurationManager.ConnectionStrings["conRsTool"].ConnectionString.Replace("(local)", txtServerName.Text).Replace("sa", txtUserID.Text).Replace("as1", txtPass.Text); ;
                string str2 = ConfigurationManager.ConnectionStrings["RSFinanceToolsEntities"].ConnectionString.Replace("127.0.0.1", txtServerName.Text).Replace("sa", txtUserID.Text).Replace("as1", txtPass.Text); ;

                Finance_Tools.ConnectionStringsSave("conRsTool", str);
                Finance_Tools.ConnectionStringsSave("RSFinanceToolsEntities", str2);
                hasSaved = true;

                Finance_Tools.AppSettingSave("isSqlInitialize", "true");

                this.Close();

            }
            catch (Exception)
            {
                MessageBox.Show("The test database connection failed! Please check the server name and database name is correct!");
                hasSaved = false;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        private void txtUserID_Validating(object sender, CancelEventArgs e)
        {
            if (string.IsNullOrEmpty(this.txtUserID.Text))
                errorProvider2.SetError(txtUserID, "Login name can't be empty!");
        }

        private void txtPass_Validating(object sender, CancelEventArgs e)
        {
            if (string.IsNullOrEmpty(this.txtPass.Text))
                errorProvider3.SetError(txtPass, "Password can't be empty!");
        }
    }
}
