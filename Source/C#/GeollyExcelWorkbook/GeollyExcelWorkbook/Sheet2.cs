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
    public partial class Sheet2
    {
        Microsoft.Office.Tools.Excel.NamedRange evntRangeLong = null; // Q10
        Microsoft.Office.Tools.Excel.NamedRange evntRangeShort = null; // R10
        Microsoft.Office.Tools.Excel.NamedRange evntRangeTrendSlow = null; // Y12
        Microsoft.Office.Tools.Excel.NamedRange evntRangeTrendFast = null; // Y13

        private void Sheet2_Startup(object sender, System.EventArgs e)
        {
            evntRangeLong = Globals.Sheet2.Controls.AddNamedRange(Globals.Sheet2.Range["Q10", missing], "evntRangeLong");
            evntRangeShort = Globals.Sheet2.Controls.AddNamedRange(Globals.Sheet2.Range["R10", missing], "evntRangeShort");
            evntRangeTrendSlow = Globals.Sheet2.Controls.AddNamedRange(Globals.Sheet2.Range["Y12", missing], "evntRangeTrendSlow");
            evntRangeTrendFast = Globals.Sheet2.Controls.AddNamedRange(Globals.Sheet2.Range["Y13", missing], "evntRangeTrendFast");         
            GeollyExcelWorkbook.Globals.ThisWorkbook.SheetCalculate += new Microsoft.Office.Interop.Excel.WorkbookEvents_SheetCalculateEventHandler(CalculateEvent);
        }

        private void Sheet2_Shutdown(object sender, System.EventArgs e)
        {
            GeollyExcelWorkbook.Globals.ThisWorkbook.SheetCalculate -= new Microsoft.Office.Interop.Excel.WorkbookEvents_SheetCalculateEventHandler(CalculateEvent);
        }

        private void CalculateEvent(Object Sh)
        {
            var connectionString = "mongodb://localhost";
            var client = new MongoClient(connectionString);
            var server = client.GetServer();
            var database = server.GetDatabase("test");

            var entitiesCollection = database.GetCollection<Entity>("entities");
            var trendCollection = database.GetCollection<Trend>("trend");
            var takeProfitCollection = database.GetCollection<TakeProfit>("takeProfit");
            var cutLossCollection = database.GetCollection<CutLoss>("cutLoss");
            var ordersCollection = database.GetCollection<Order>("orders");

            var entitiesQuery = Query<Entity>.EQ(entity => entity.Start_Trading, true);
            var trendQuery = Query<Trend>.EQ(trend => trend.Start_Trading, true);
            var takeProfitQuery = Query<TakeProfit>.EQ(takeProfit => takeProfit.Start_Trading, true);
            var cutLossQuery = Query<CutLoss>.EQ(cutLoss => cutLoss.Start_Trading, true);
            var ordersQuery = Query<Order>.EQ(order => order.Start_Trading, true);

            if (evntRangeLong.Cells.Value2.ToString() == "1")
            {
                var entitiesUpdate = Update<Entity>.Set(entity => entity.Order_Type, "Buy"); // update modifiers
                entitiesCollection.Update(entitiesQuery, entitiesUpdate);
            }
            if (evntRangeShort.Cells.Value2.ToString() == "1")
            {
                var entitiesUpdate = Update<Entity>.Set(entity => entity.Order_Type, "Sell"); // update modifiers
                entitiesCollection.Update(entitiesQuery, entitiesUpdate);
            }

            if ( (entitiesCollection.Find(entitiesQuery).SetLimit(1).FirstOrDefault().Order_Type == "Buy" ||
                  entitiesCollection.Find(entitiesQuery).SetLimit(1).FirstOrDefault().Order_Type.Equals("Buy")) && 
                 (evntRangeTrendSlow.Cells.Value2 > -1 || evntRangeTrendFast.Cells.Value2 > -1) )
            {
                var trendUpdate = Update<Trend>.Set(trend => trend.Order_Type, "Buy"); // update modifiers
                trendCollection.Update(trendQuery, trendUpdate);
            }

            if ( (entitiesCollection.Find(entitiesQuery).SetLimit(1).FirstOrDefault().Order_Type == "Sell" ||
                  entitiesCollection.Find(entitiesQuery).SetLimit(1).FirstOrDefault().Order_Type.Equals("Sell")) &&
                 (evntRangeTrendSlow.Cells.Value2 < 1 || evntRangeTrendFast.Cells.Value2 < 1) )
            {
                var trendUpdate = Update<Trend>.Set(trend => trend.Order_Type, "Sell"); // update modifiers
                trendCollection.Update(trendQuery, trendUpdate);
            }

            if (ordersCollection.Find(ordersQuery).SetLimit(1).FirstOrDefault().Filled_Order == "Buy" ||
                ordersCollection.Find(ordersQuery).SetLimit(1).FirstOrDefault().Filled_Order.Equals("Buy"))
            {
                if (evntRangeTrendSlow.Cells.Value2 >= 0 && evntRangeTrendFast.Cells.Value2 >= 0) //Take Profit
                {
                    var takeProfitUpdate = Update<TakeProfit>.Set(takeProfit => takeProfit.Order_Type, "Take Profit"); // update modifiers
                    var cutLossUpdate = Update<CutLoss>.Set(cutLoss => cutLoss.Order_Type, "No Order"); // update modifiers

                    takeProfitCollection.Update(takeProfitQuery, takeProfitUpdate);
                    cutLossCollection.Update(cutLossQuery, cutLossUpdate);
                }
                else if (evntRangeTrendSlow.Cells.Value2 == -1 && evntRangeTrendFast.Cells.Value2 == -1) // Cut Loss
                {
                    var takeProfitUpdate = Update<TakeProfit>.Set(takeProfit => takeProfit.Order_Type, "No Order"); // update modifiers
                    var cutLossUpdate = Update<CutLoss>.Set(cutLoss => cutLoss.Order_Type, "Cut Loss"); // update modifiers

                    takeProfitCollection.Update(takeProfitQuery, takeProfitUpdate);
                    cutLossCollection.Update(cutLossQuery, cutLossUpdate);
                }
                else // No Order
                {
                    var takeProfitUpdate = Update<TakeProfit>.Set(takeProfit => takeProfit.Order_Type, "No Order"); // update modifiers
                    var cutLossUpdate = Update<CutLoss>.Set(cutLoss => cutLoss.Order_Type, "No Order"); // update modifiers

                    takeProfitCollection.Update(takeProfitQuery, takeProfitUpdate);
                    cutLossCollection.Update(cutLossQuery, cutLossUpdate);
                }
            }

            if (ordersCollection.Find(ordersQuery).SetLimit(1).FirstOrDefault().Filled_Order == "Sell" ||
                ordersCollection.Find(ordersQuery).SetLimit(1).FirstOrDefault().Filled_Order.Equals("Sell"))
            {
                if (evntRangeTrendSlow.Cells.Value2 <= 0 && evntRangeTrendFast.Cells.Value2 <= 0) //Take Profit
                {
                    var takeProfitUpdate = Update<TakeProfit>.Set(takeProfit => takeProfit.Order_Type, "Take Profit"); // update modifiers
                    var cutLossUpdate = Update<CutLoss>.Set(cutLoss => cutLoss.Order_Type, "No Order"); // update modifiers

                    takeProfitCollection.Update(takeProfitQuery, takeProfitUpdate);
                    cutLossCollection.Update(cutLossQuery, cutLossUpdate);
                }
                else if (evntRangeTrendSlow.Cells.Value2 == 1 && evntRangeTrendFast.Cells.Value2 == 1) // Cut Loss
                {
                    var takeProfitUpdate = Update<TakeProfit>.Set(takeProfit => takeProfit.Order_Type, "No Order"); // update modifiers
                    var cutLossUpdate = Update<CutLoss>.Set(cutLoss => cutLoss.Order_Type, "Cut Loss"); // update modifiers

                    takeProfitCollection.Update(takeProfitQuery, takeProfitUpdate);
                    cutLossCollection.Update(cutLossQuery, cutLossUpdate);
                }
                else // No Order
                {
                    var takeProfitUpdate = Update<TakeProfit>.Set(takeProfit => takeProfit.Order_Type, "No Order"); // update modifiers
                    var cutLossUpdate = Update<CutLoss>.Set(cutLoss => cutLoss.Order_Type, "No Order"); // update modifiers

                    takeProfitCollection.Update(takeProfitQuery, takeProfitUpdate);
                    cutLossCollection.Update(cutLossQuery, cutLossUpdate);
                }
            }     
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Sheet2_Startup);
            this.Shutdown += new System.EventHandler(Sheet2_Shutdown);
        }

        #endregion

    }
}
