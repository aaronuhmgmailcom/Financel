
/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<SessionInfo>   
 * Author：Peter.uhm  (yanb@shinetechchina.com)
 * Modify date：2016.04
 * Modify date：2016.09
 * Version : 2.0.0.2
 */
using System;
using System.Collections.Generic;
using System.Xml.Linq;
using System.Web;
using System.Collections.Specialized;

namespace ExcelAddIn4
{
    internal class SessionInfo
    {
        public static UserInfo UserInfo
        {
            get;
            set;
        }
    }
}