using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Excel.Extensions;
using Microsoft.Office.Tools;

namespace ExcelAddIn2
{
    public partial class ThisAddIn
    {
        public CustomTaskPane _MyCustomTaskPane = null;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Form1 f = new Form1();
            //f.Show();
            //System.Configuration.

            UCForTaskPane taskPane = new UCForTaskPane();
            _MyCustomTaskPane = this.CustomTaskPanes.Add(taskPane, "My Task Pane");
            _MyCustomTaskPane.Width = 200;
            _MyCustomTaskPane.Visible = true;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
