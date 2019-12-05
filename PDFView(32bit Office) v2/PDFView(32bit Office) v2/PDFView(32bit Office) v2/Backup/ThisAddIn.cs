/*  
 * Module ID：<ExcelAddIn4>   
 * Function：<ThisAddIn>   
 * Author：Peter.uhm  (yanb@shinetechchina.com)
 * Modify date：2013.3
 * Modify date：2016.04
 * Version : 2.0.0.2
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Excel.Controls;
using Microsoft.Office.Tools;
using System.IO;
using System.Diagnostics;
using ExcelAddIn4.Common;

namespace ExcelAddIn4
{
    public partial class ThisAddIn
    {
        /// <summary>
        /// 
        /// </summary>
        internal static Finance_Tools ft
        {
            get { return new Finance_Tools(); }
        }
        /// <summary>
        /// LogHelper.WriteStartUpLog(typeof(ThisAddIn), "Finance Tool Start up");
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            DirectoryInfo mypath = new DirectoryInfo(Finance_Tools.strRSDataCache);
            if (mypath.Exists)
            { }
            else
            {
                mypath.Create();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Finance_Tools.totalRunTimes++;
            if (!Finance_Tools.CompareFileCountWithConfigCount()) Finance_Tools.UpdateFolderFileXMLStructure();
            ft.deleteCache(Finance_Tools.strRSDataCache);
            this.CustomTaskPanes.Dispose();
            if (Ribbon2.notifyIcon1 != null)
                Ribbon2.notifyIcon1.Dispose();
        }
        #region VSTO generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
