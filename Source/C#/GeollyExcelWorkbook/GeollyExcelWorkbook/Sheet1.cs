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
using MongoDB.Bson;
using MongoDB.Driver;
using MongoDB.Driver.Builders;

namespace GeollyExcelWorkbook
{
    public partial class Sheet1
    {
        Microsoft.Office.Tools.Excel.NamedRange evntRangeAccountName = null; // A2
        Microsoft.Office.Tools.Excel.NamedRange evntRangeProductName = null; // B2
        Microsoft.Office.Tools.Excel.NamedRange evntRangeInstrumentName = null; // C2

        private void Sheet1_Startup(object sender, System.EventArgs e)
        {
            evntRangeAccountName = Globals.Sheet1.Controls.AddNamedRange(Globals.Sheet1.Range["A2", missing], "evntRangeAccountName");
            evntRangeProductName = Globals.Sheet1.Controls.AddNamedRange(Globals.Sheet1.Range["B2", missing], "evntRangeProductName");
            evntRangeInstrumentName = Globals.Sheet1.Controls.AddNamedRange(Globals.Sheet1.Range["C2", missing], "evntRangeInstrumentName");

            var connectionString = "mongodb://localhost";
            var client = new MongoClient(connectionString);
            var server = client.GetServer();
            var database = server.GetDatabase("test");
            var collection = database.GetCollection<Account>("account");

            var query = Query<Account>.EQ(account => account.Account_Name, "GQ");

            if (collection.Count(query) == 0)
            {
                var account = new Account { Start_Trading = true, Account_Name = evntRangeAccountName.Cells.Value2.ToString(), Product_Name = evntRangeProductName.Cells.Value2.ToString(), Instrument_Name = evntRangeInstrumentName.Cells.Value2.ToString() };
                collection.Insert(account);
            }
        }

        private void Sheet1_Shutdown(object sender, System.EventArgs e)
        {

        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Sheet1_Startup);
            this.Shutdown += new System.EventHandler(Sheet1_Shutdown);
        }

        #endregion

    }
}
