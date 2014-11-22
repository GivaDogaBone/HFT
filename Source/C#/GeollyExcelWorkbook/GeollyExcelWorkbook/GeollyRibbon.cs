using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using MongoDB.Bson;
using MongoDB.Driver;
using MongoDB.Driver.Builders;

namespace GeollyExcelWorkbook
{
    public class Entity
    {
        public ObjectId Id { get; set; }
        public Boolean Start_Trading { get; set; }
        public string Order_Type { get; set; }
    }

    public class Trend
    {
        public ObjectId Id { get; set; }
        public Boolean Start_Trading { get; set; }
        public string Order_Type { get; set; }
    }

    public class TakeProfit
    {
        public ObjectId Id { get; set; }
        public Boolean Start_Trading { get; set; }
        public string Order_Type { get; set; }
    }

    public class CutLoss
    {
        public ObjectId Id { get; set; }
        public Boolean Start_Trading { get; set; }
        public string Order_Type { get; set; }
    }

    public class Account
    {
        public ObjectId Id { get; set; }
        public Boolean Start_Trading { get; set; }
        public string Account_Name { get; set; }
        public string Product_Name { get; set; }
        public string Instrument_Name { get; set; }
    }

    public class Order
    {
        public ObjectId Id { get; set; }
        public Boolean Start_Trading { get; set; }
        public String Filled_Order { get; set; }
        public Boolean Order_Executed { get; set; }
    }

    public partial class GeollyRibbon
    {
        private ObjectId entitiesIdentity;
        private ObjectId trendIdentity;
        private ObjectId takeProfitIdentity;
        private ObjectId cutLossIdentity;

        private void GeollyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
       
        }

        private void Reset_TT_API_Button_Click(object sender, RibbonControlEventArgs e)
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

                MessageBox.Show("Start_Trading_Click collection.Count(query) was Zero for Entity " + entitiesIdentity.ToString());
            }
            else
            {
                MessageBox.Show("Start_Trading_Click collection.Count(query) is not Zero for Entity " + entitiesIdentity.ToString());
            }

            if (trendCollection.Count(trendQuery) == 0)
            {
                var trend = new Trend { Start_Trading = true, Order_Type = "No Order" };
                trendCollection.Insert(trend);
                trendIdentity = trend.Id;

                MessageBox.Show("Start_Trading_Click collection.Count(query) was Zero for Trend " + trendIdentity.ToString());
            }
            else
            {
                MessageBox.Show("Start_Trading_Click collection.Count(query) is not Zero for Trend " + trendIdentity.ToString());
            }

            if (takeProfitCollection.Count(takeProfitQuery) == 0)
            {
                var takeProfit = new TakeProfit { Start_Trading = true, Order_Type = "No Order" };
                takeProfitCollection.Insert(takeProfit);
                takeProfitIdentity = takeProfit.Id;

                MessageBox.Show("Start_Trading_Click collection.Count(query) was Zero for TakeProfit " + takeProfitIdentity.ToString());
            }
            else
            {
                MessageBox.Show("Start_Trading_Click collection.Count(query) is not Zero for TakeProfit " + takeProfitIdentity.ToString());
            }

            if (cutLossCollection.Count(cutLossQuery) == 0)
            {
                var cutLoss = new CutLoss { Start_Trading = true, Order_Type = "No Order" };
                cutLossCollection.Insert(cutLoss);
                cutLossIdentity = cutLoss.Id;

                MessageBox.Show("Start_Trading_Click collection.Count(query) was Zero for CutLoss " + cutLossIdentity.ToString());
            }
            else
            {
                MessageBox.Show("Start_Trading_Click collection.Count(query) is not Zero for CutLoss " + cutLossIdentity.ToString());
            }

        }

        private void STOP_TT_API_Button_Click(object sender, RibbonControlEventArgs e)
        {
            var connectionString = "mongodb://localhost";
            var client = new MongoClient(connectionString);
            var server = client.GetServer();
            var database = server.GetDatabase("test");

            var entitiesCollection = database.GetCollection<Entity>("entities");
            var trendCollection = database.GetCollection<Trend>("trend");
            var takeProfitCollection = database.GetCollection<TakeProfit>("takeProfit");
            var cutLossCollection = database.GetCollection<CutLoss>("cutLoss");

            var entitiesQuery = Query<Entity>.EQ(entity => entity.Start_Trading, true);
            var trendQuery = Query<Trend>.EQ(trend => trend.Start_Trading, true);
            var takeProfitQuery = Query<TakeProfit>.EQ(takeProfit => takeProfit.Start_Trading, true);
            var cutLossQuery = Query<CutLoss>.EQ(cutLoss => cutLoss.Start_Trading, true);

            entitiesCollection.Remove(entitiesQuery);
            trendCollection.Remove(trendQuery);
            takeProfitCollection.Remove(takeProfitQuery);
            cutLossCollection.Remove(cutLossQuery);

            MessageBox.Show("Stop_Trading_Click !!!");

        }

        private void BuyButton_Click(object sender, RibbonControlEventArgs e)
        {
            var connectionString = "mongodb://localhost";
            var client = new MongoClient(connectionString);
            var server = client.GetServer();
            var database = server.GetDatabase("test");
            var entitiesCollection = database.GetCollection<Entity>("entities");
            var trendCollection = database.GetCollection<Trend>("trend");

            var entityQuery = Query<Entity>.EQ(entity => entity.Start_Trading, true);
            var entityUpdate = Update<Entity>.Set(entity => entity.Order_Type, "Buy"); // update modifiers
            entitiesCollection.Update(entityQuery, entityUpdate);

            var trendQuery = Query<Trend>.EQ(trend => trend.Start_Trading, true);
            var trendUpdate = Update<Trend>.Set(trend => trend.Order_Type, "Buy"); // update modifiers
            trendCollection.Update(trendQuery, trendUpdate);

            MessageBox.Show("Buy_Click !!!");

        }

        private void SellButton_Click(object sender, RibbonControlEventArgs e)
        {
            var connectionString = "mongodb://localhost";
            var client = new MongoClient(connectionString);
            var server = client.GetServer();
            var database = server.GetDatabase("test");
            var entitiesCollection = database.GetCollection<Entity>("entities");
            var trendCollection = database.GetCollection<Trend>("trend");

            var entitiesQuery = Query<Entity>.EQ(entity => entity.Start_Trading, true);
            var entitiesUpdate = Update<Entity>.Set(entity => entity.Order_Type, "Sell"); // update modifiers
            entitiesCollection.Update(entitiesQuery, entitiesUpdate);

            var trendQuery = Query<Trend>.EQ(trend => trend.Start_Trading, true);
            var trendUpdate = Update<Trend>.Set(trend => trend.Order_Type, "Sell"); // update modifiers
            trendCollection.Update(trendQuery, trendUpdate);

            MessageBox.Show("Sell_Click !!!");

        }

        private void TakeProfitButton_Click(object sender, RibbonControlEventArgs e)
        {
            var connectionString = "mongodb://localhost";
            var client = new MongoClient(connectionString);
            var server = client.GetServer();
            var database = server.GetDatabase("test");

            var takeProfitCollection = database.GetCollection<TakeProfit>("takeProfit");
            var cutLossCollection = database.GetCollection<CutLoss>("cutLoss");

            var takeProfitQuery = Query<TakeProfit>.EQ(takeProfit => takeProfit.Start_Trading, true);
            var cutLossQuery = Query<CutLoss>.EQ(cutLoss => cutLoss.Start_Trading, true);

            var takeProfitUpdate = Update<TakeProfit>.Set(takeProfit => takeProfit.Order_Type, "Take Profit"); // update modifiers
            var cutLossUpdate = Update<CutLoss>.Set(cutLoss => cutLoss.Order_Type, "No Order"); // update modifiers

            takeProfitCollection.Update(takeProfitQuery, takeProfitUpdate);
            cutLossCollection.Update(cutLossQuery, cutLossUpdate);

            MessageBox.Show("TakeProfitButton_Click !!!");

        }

        private void CutLossButton_Click(object sender, RibbonControlEventArgs e)
        {
            var connectionString = "mongodb://localhost";
            var client = new MongoClient(connectionString);
            var server = client.GetServer();
            var database = server.GetDatabase("test");

            var takeProfitCollection = database.GetCollection<TakeProfit>("takeProfit");
            var cutLossCollection = database.GetCollection<CutLoss>("cutLoss");

            var takeProfitQuery = Query<TakeProfit>.EQ(takeProfit => takeProfit.Start_Trading, true);
            var cutLossQuery = Query<CutLoss>.EQ(cutLoss => cutLoss.Start_Trading, true);

            var takeProfitUpdate = Update<TakeProfit>.Set(takeProfit => takeProfit.Order_Type, "No Order"); // update modifiers
            var cutLossUpdate = Update<CutLoss>.Set(cutLoss => cutLoss.Order_Type, "Cut Loss"); // update modifiers

            takeProfitCollection.Update(takeProfitQuery, takeProfitUpdate);
            cutLossCollection.Update(cutLossQuery, cutLossUpdate);

            MessageBox.Show("CutLossButton_Click !!!");

        }

    }
}
