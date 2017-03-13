using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelTools
{
    public partial class ThisWorkbook
    {
        private ActionsPaneControl aPane;
        public static Excel.Application app;
        private void CreateActionsPane()
        {

            ThisWorkbook.app = Application;
            Excel.Application app = Application;
            //Add the user control to the actions pane
            aPane = new ActionsPaneControl();
            ActionsPane.Controls.Add(aPane);
            aPane.Width = 200;
            aPane.Visible = true;
        }
        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            CreateActionsPane();
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
        }

        #endregion

    }
}
