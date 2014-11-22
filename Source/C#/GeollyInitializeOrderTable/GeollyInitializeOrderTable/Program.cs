using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using MongoDB.Bson;
using MongoDB.Driver;
using MongoDB.Driver.Builders;

namespace GeollyInitializeOrderTable
{
    public class Order
    {
        public ObjectId Id { get; set; }
        public Boolean Start_Trading { get; set; }
        public String Filled_Order { get; set; }
        public Boolean Order_Executed { get; set; }
    }

    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            var connectionString = "mongodb://localhost";
            var client = new MongoClient(connectionString);
            var server = client.GetServer();
            var database = server.GetDatabase("test");
            var collection = database.GetCollection<Order>("orders");

            var query = Query<Order>.EQ(order => order.Start_Trading, true);

            collection.Remove(query);
            var tradeOrder = new Order { Start_Trading = true, Filled_Order = "No Order", Order_Executed = false };
            collection.Insert(tradeOrder);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
