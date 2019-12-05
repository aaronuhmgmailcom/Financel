using System;
using System.Collections.Generic;
using System.Web;
using System.Text;
using System.Web.Script.Serialization;

namespace ExcelAddIn3
{
    public static class XmlDataHelper
    {
        private static string FormatDateTime(DateTime? dt)
        {
            string result = string.Empty;
            if (dt != null)
                result = string.Format("{0:yyyy/MM/dd hh:mm tt}", dt.Value);
            return result;
        }
    }

}