using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using MongoDB.Bson;
using MongoDB.Driver;
using MongoDB.Driver.Builders;

namespace GeollyTTAPIConsoleApplication
{
    public class Entity
    {
        public ObjectId Id { get; set; }
        public Boolean Start_Trading { get; set; }
        public string Order_Type { get; set; }
        public Boolean Order_Executed { get; set; }
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

    class Program
    {
        static void Main(string[] args)
        {
            // Dictates whether TT API will be started on its own thread
            bool startOnSeparateThread = false;

            if (startOnSeparateThread)
            {
                // Start TT API on a separate thread
                TTAPIFunctions tf = new TTAPIFunctions();
                Thread workerThread = new Thread(tf.Start);
                workerThread.Name = "TT API Thread";
                workerThread.Start();
            }
            else
            {
                // Start the TT API on the same thread
                using (TTAPIFunctions tf = new TTAPIFunctions())
                {
                    tf.Start();
                }
            }
        }
    }
}
