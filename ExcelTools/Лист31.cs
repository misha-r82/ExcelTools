﻿using System;
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
    public partial class Лист3
    {

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Таблица1.Change += new Microsoft.Office.Tools.Excel.ListObjectChangeHandler(this.Таблица1_Change);

        }


        #endregion

        private void Таблица1_Change(Excel.Range targetRange, ListRanges changedRanges)
        {

        }
    }
}
