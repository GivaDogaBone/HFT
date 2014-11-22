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

namespace GeollyExcelWorkbook
{
    public partial class Tabelle2
    {
        private void Tabelle2_Startup(object sender, System.EventArgs e)
        {
        }

        private void Tabelle2_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Tabelle2_Startup);
            this.Shutdown += new System.EventHandler(Tabelle2_Shutdown);
        }

        #endregion

    }
}
