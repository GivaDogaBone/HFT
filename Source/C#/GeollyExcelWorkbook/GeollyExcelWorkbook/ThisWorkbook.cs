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
    public partial class ThisWorkbook
    {
        private ObjectId entitiesIdentity;
        private ObjectId trendIdentity;
        private ObjectId takeProfitIdentity;
        private ObjectId cutLossIdentity;

        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            var connectionString = "mongodb://localhost";
            var client = new MongoClient(connectionString);
            var server = client.GetServer();
            var database = server.GetDatabase("test");

            var entitiesCollection = database.GetCollection<Entity>("entities");
            var trendCollection = database.GetCollection<Trend>("trend");
            var takeProfitCollection = database.GetCollection<TakeProfit>("takeProfit");
            var cutLossCollection = database.GetCollection<CutLoss>("cutLoss");

            var entitiesQuery = Query<Entity>.EQ(entity => entity.Id, entitiesIdentity);
            var trendQuery = Query<Trend>.EQ(trend => trend.Id, trendIdentity);
            var takeProfitQuery = Query<TakeProfit>.EQ(takeProfit => takeProfit.Id, takeProfitIdentity);
            var cutLossQuery = Query<CutLoss>.EQ(cutLoss => cutLoss.Id, cutLossIdentity);

            var entitiesDeleteQuery = Query<Entity>.EQ(entity => entity.Start_Trading, true);
            var trendDeleteQuery = Query<Trend>.EQ(trend => trend.Start_Trading, true);
            var takeProfitDeleteQuery = Query<TakeProfit>.EQ(takeProfit => takeProfit.Start_Trading, true);
            var cutLossDeleteQuery = Query<CutLoss>.EQ(cutLoss => cutLoss.Start_Trading, true);

            entitiesCollection.Remove(entitiesDeleteQuery);
            trendCollection.Remove(trendDeleteQuery);
            takeProfitCollection.Remove(takeProfitDeleteQuery);
            cutLossCollection.Remove(cutLossDeleteQuery);
            
            if (entitiesCollection.Count(entitiesQuery) == 0)
            {
                var entity = new Entity { Start_Trading = true, Order_Type = "No Order" };
                entitiesCollection.Insert(entity);
                entitiesIdentity = entity.Id;
            }

            if (trendCollection.Count(trendQuery) == 0)
            {
                var trend = new Trend { Start_Trading = true, Order_Type = "No Order" };
                trendCollection.Insert(trend);
                trendIdentity = trend.Id;
            }

            if (takeProfitCollection.Count(takeProfitQuery) == 0)
            {
                var takeProfit = new TakeProfit { Start_Trading = true, Order_Type = "No Order" };
                takeProfitCollection.Insert(takeProfit);
                takeProfitIdentity = takeProfit.Id;
            }

            if (cutLossCollection.Count(cutLossQuery) == 0)
            {
                var cutLoss = new CutLoss { Start_Trading = true, Order_Type = "No Order" };
                cutLossCollection.Insert(cutLoss);
                cutLossIdentity = cutLoss.Id;
            }       
        
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
            var connectionString = "mongodb://localhost";
            var client = new MongoClient(connectionString);
            var server = client.GetServer();
            var database = server.GetDatabase("test");

            var accountCollection = database.GetCollection<Entity>("account");
            var entitiesCollection = database.GetCollection<Entity>("entities");
            var trendCollection = database.GetCollection<Trend>("trend");
            var takeProfitCollection = database.GetCollection<TakeProfit>("takeProfit");
            var cutLossCollection = database.GetCollection<CutLoss>("cutLoss");

            var accountQuery = Query<Account>.EQ(account => account.Start_Trading, true);
            var entitiesQuery = Query<Entity>.EQ(entity => entity.Start_Trading, true);
            var trendQuery = Query<Trend>.EQ(trend => trend.Start_Trading, true);
            var takeProfitQuery = Query<TakeProfit>.EQ(takeProfit => takeProfit.Start_Trading, true);
            var cutLossQuery = Query<CutLoss>.EQ(cutLoss => cutLoss.Start_Trading, true);

            accountCollection.Remove(accountQuery);
            entitiesCollection.Remove(entitiesQuery);
            trendCollection.Remove(trendQuery);
            takeProfitCollection.Remove(takeProfitQuery);
            cutLossCollection.Remove(cutLossQuery);

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
