using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Security.Principal;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
namespace ExcelAddIn4.Common
{
    public class BasePage
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
        internal static RSFinanceToolsEntities db
        {
            get { return new RSFinanceToolsEntities(); }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="DesiredFunctionCode"></param>
        /// <param name="type"></param>
        /// <param name="C"></param>
        public static void VerifyButton(string DesiredFunctionCode, string type, RibbonControl C)
        {
            string userName = WindowsIdentity.GetCurrent().Name;
            var wUserID = (from FT_user in db.rsUsers
                           where FT_user.WindowsUserID == userName
                           select FT_user.ft_id).First();
            var admin = (from FT_user in db.rsUsers
                         where FT_user.WindowsUserID == userName
                         select FT_user.LoginType).First();
            if (admin == 0)
            {
                C.Visible = true;
                return;
            }
            System.Collections.Generic.List<string> list = ft.GetUserGroups(wUserID.ToString());
            for (int i = 0; i < list.Count; i++)
            {
                string groupid = list[i];
                bool groupdisable = (bool)ft.GetGroupDisableByID(int.Parse(groupid));
                if (groupdisable)
                    continue;
                DataTable dt = ft.GetGroupPermissionsView(groupid);
                for (int j = 0; j < dt.Rows.Count; j++)
                    if ((type == dt.Rows[j]["Per_Type"].ToString()) && (DesiredFunctionCode == dt.Rows[j]["PermissionName"].ToString()))
                    {
                        C.Visible = true;
                        return;
                    }
            }
            C.Visible = false;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="DesiredFunctionCode"></param>
        /// <param name="type"></param>
        /// <param name="C"></param>
        public static void VerifyDelButton(string DesiredFunctionCode, string type, Control C)
        {
            string userName = WindowsIdentity.GetCurrent().Name;
            var wUserID = (from FT_user in db.rsUsers
                           where FT_user.WindowsUserID == userName
                           select FT_user.ft_id).First();
            var admin = (from FT_user in db.rsUsers
                         where FT_user.WindowsUserID == userName
                         select FT_user.LoginType).First();
            string[] arr = { DesiredFunctionCode + ".xlsm - Delete", DesiredFunctionCode + ".xlsx - Delete", DesiredFunctionCode + ".xls - Delete" };
            if (admin == 0)
            {
                C.Visible = true;
                return;
            }
            System.Collections.Generic.List<string> list = ft.GetUserGroups(wUserID.ToString());
            for (int i = 0; i < list.Count; i++)
            {
                string groupid = list[i];
                bool groupdisable = (bool)ft.GetGroupDisableByID(int.Parse(groupid));
                if (groupdisable)
                    continue;
                DataTable dt = ft.GetGroupPermissionsView(groupid);
                for (int j = 0; j < dt.Rows.Count; j++)
                    if ((type == dt.Rows[j]["Per_Type"].ToString()) && (arr.Contains(dt.Rows[j]["PermissionName"].ToString())))
                    {
                        C.Visible = true;
                        return;
                    }
            }
            C.Visible = false;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="DesiredFunctionCode"></param>
        /// <returns></returns>
        public static bool VerifyAmendButton(string DesiredFunctionCode)
        {
            string userName = WindowsIdentity.GetCurrent().Name;
            var wUserID = (from FT_user in db.rsUsers
                           where FT_user.WindowsUserID == userName
                           select FT_user.ft_id).First();
            var admin = (from FT_user in db.rsUsers
                         where FT_user.WindowsUserID == userName
                         select FT_user.LoginType).First();
            string[] arr = { DesiredFunctionCode + ".xlsm - Save/Amend", DesiredFunctionCode + ".xlsx - Save/Amend", DesiredFunctionCode + ".xls - Save/Amend"
                           ,DesiredFunctionCode + ".xlsm - Delete", DesiredFunctionCode + ".xlsx - Delete", DesiredFunctionCode + ".xls - Delete" };
            if (admin == 0)
                return true;
            System.Collections.Generic.List<string> list = ft.GetUserGroups(wUserID.ToString());
            for (int i = 0; i < list.Count; i++)
            {
                string groupid = list[i];
                bool groupdisable = (bool)ft.GetGroupDisableByID(int.Parse(groupid));
                if (groupdisable)
                    continue;
                DataTable dt = ft.GetGroupPermissionsView(groupid);
                for (int j = 0; j < dt.Rows.Count; j++)
                    if (("0" == dt.Rows[j]["Per_Type"].ToString() || dt.Rows[j]["Per_Type"].ToString() == "5") && (arr.Contains(dt.Rows[j]["PermissionName"].ToString())))
                        return true;
            }
            return false;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="DesiredFunctionCode"></param>
        /// <param name="C"></param>
        public static void VerifyWriteButton(string DesiredFunctionCode, Control C)
        {
            string userName = WindowsIdentity.GetCurrent().Name;
            var wUserID = (from FT_user in db.rsUsers
                           where FT_user.WindowsUserID == userName
                           select FT_user.ft_id).First();
            var admin = (from FT_user in db.rsUsers
                         where FT_user.WindowsUserID == userName
                         select FT_user.LoginType).First();
            string[] arr = { DesiredFunctionCode + ".xlsm - Write", DesiredFunctionCode + ".xlsx - Write", DesiredFunctionCode + ".xls - Write" ,
                            DesiredFunctionCode + ".xlsm - Save/Amend", DesiredFunctionCode + ".xlsx - Save/Amend", DesiredFunctionCode + ".xls - Save/Amend"
                           ,DesiredFunctionCode + ".xlsm - Delete", DesiredFunctionCode + ".xlsx - Delete", DesiredFunctionCode + ".xls - Delete" };
            if (admin == 0)
            {
                C.Visible = true;
                return;
            }
            System.Collections.Generic.List<string> list = ft.GetUserGroups(wUserID.ToString());
            for (int i = 0; i < list.Count; i++)
            {
                string groupid = list[i];
                bool groupdisable = (bool)ft.GetGroupDisableByID(int.Parse(groupid));
                if (groupdisable)
                    continue;
                DataTable dt = ft.GetGroupPermissionsView(groupid);
                for (int j = 0; j < dt.Rows.Count; j++)
                    if (("1" == dt.Rows[j]["Per_Type"].ToString() || "0" == dt.Rows[j]["Per_Type"].ToString() || dt.Rows[j]["Per_Type"].ToString() == "5") && (arr.Contains(dt.Rows[j]["PermissionName"].ToString())))
                    {
                        C.Visible = true;
                        return;
                    }
            }
            C.Visible = false;
        }
    }
}
