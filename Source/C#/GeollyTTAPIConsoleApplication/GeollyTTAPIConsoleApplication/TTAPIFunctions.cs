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
    using TradingTechnologies.TTAPI;
    using TradingTechnologies.TTAPI.Tradebook;

    /// <summary>
    /// Main TT API class
    /// </summary>
    class TTAPIFunctions : IDisposable
    {
        /// <summary>
        /// Declare the TTAPI objects
        /// </summary>
        private XTraderModeTTAPI m_apiInstance = null;
        private WorkerDispatcher m_disp = null;
        private WorkerDispatcher m_disp_buy = null;
        private WorkerDispatcher m_disp_sell = null;
        private bool m_disposed = false;
        private object m_lock = new object();
        private InstrumentLookupSubscription m_req = null;
        private PriceSubscription m_ps = null;
        private InstrumentTradeSubscription m_ts = null;
        private Thread workerThread = null;
        private string m_orderKey = null;
        private string m_orderType = null;

        /// <summary>
        /// Default constructor
        /// </summary>
        public TTAPIFunctions()
        {
        }

        /// <summary>
        /// Create and start the Dispatcher
        /// </summary>
        public void Start()
        {
            // Attach a WorkerDispatcher to the current thread
            m_disp = Dispatcher.AttachWorkerDispatcher();
            m_disp.BeginInvoke(new Action(Init));
            m_disp.Run();
        }

        /// <summary>
        /// Initialize TT API
        /// </summary>
        public void Init()
        {
            // Use "Follow X_TRADER" Login Mode
            ApiInitializeHandler h = new ApiInitializeHandler(ttApiInitComplete);
            TTAPI.CreateXTraderModeTTAPI(Dispatcher.Current, h);
        }

        /// <summary>
        /// Event notification for status of TT API initialization
        /// </summary>
        public void ttApiInitComplete(TTAPI api, ApiCreationException ex)
        {
            if (ex == null)
            {
                // Authenticate your credentials
                m_apiInstance = (XTraderModeTTAPI)api;
                m_apiInstance.ConnectionStatusUpdate += new EventHandler<ConnectionStatusUpdateEventArgs>(m_apiInstance_ConnectionStatusUpdate);
                m_apiInstance.Start();
            }
            else
            {
                Console.WriteLine("TT API Initialization Failed: {0}", ex.Message);
                Dispose();
            }
        }

        /// <summary>
        /// Event notification for status of authentication
        /// </summary>
        public void m_apiInstance_ConnectionStatusUpdate(object sender, ConnectionStatusUpdateEventArgs e)
        {
            if (e.Status.IsSuccess)
            {
                // Add code here to begin working with the TT API
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

                try
                {
                    Console.WriteLine("Initial Entities Table   : " + entitiesCollection.Find(entitiesQuery).SetLimit(1).FirstOrDefault().Order_Type);
                    Console.WriteLine("Initial Trend Table      : " + trendCollection.Find(trendQuery).SetLimit(1).FirstOrDefault().Order_Type);
                    Console.WriteLine("Initial TakeProfit Table : " + takeProfitCollection.Find(takeProfitQuery).SetLimit(1).FirstOrDefault().Order_Type);
                    Console.WriteLine("Initial CutLoss Table    : " + cutLossCollection.Find(cutLossQuery).SetLimit(1).FirstOrDefault().Order_Type);
                    Console.WriteLine("Initial Orders Table     : " + ordersCollection.Find(ordersQuery).SetLimit(1).FirstOrDefault().Filled_Order);
                    while (entitiesCollection.Count(entitiesQuery) > 0)
                    {
                        if (!ordersCollection.Find(ordersQuery).SetLimit(1).FirstOrDefault().Filled_Order.Equals(entitiesCollection.Find(entitiesQuery).SetLimit(1).FirstOrDefault().Order_Type) ||
                            ordersCollection.Find(ordersQuery).SetLimit(1).FirstOrDefault().Filled_Order != entitiesCollection.Find(entitiesQuery).SetLimit(1).FirstOrDefault().Order_Type)
                        {
                            if ((entitiesCollection.Find(entitiesQuery).SetLimit(1).FirstOrDefault().Order_Type.Equals("Buy") ||
                                  entitiesCollection.Find(entitiesQuery).SetLimit(1).FirstOrDefault().Order_Type == "Buy") &&
                                 (trendCollection.Find(trendQuery).SetLimit(1).FirstOrDefault().Order_Type.Equals("Buy") ||
                                  trendCollection.Find(trendQuery).SetLimit(1).FirstOrDefault().Order_Type == "Buy"))
                            {
                                Console.WriteLine(GetTimestamp(DateTime.Now) + " : I am Buying !!!");
                                update_Order_Type_Table("Buy");
                                // Start TT API on a separate thread
                                // Call the SendOrder function with a Buy request.
                                workerThread = new Thread(Buy_Order_Dispatcher);
                                workerThread.Name = "Buy Order from Geolly";
                                workerThread.Start();
                                // Loop until worker thread activates. 
                                while (!workerThread.IsAlive) ;
                                // Put the main thread to sleep for 60000 millisecond to 
                                // allow the worker thread to do some work:
                                Thread.Sleep(60000);
                                // Request that the worker thread stop itself:
                                workerThread.Abort();
                                // Use the Join method to block the current thread  
                                // until the object's thread terminates.
                                workerThread.Join();
                                workerThread = null;
                                m_disp_buy = null;
                                m_req = null;
                                m_ps = null;
                                m_ts = null;
                                m_orderKey = null;
                                m_orderType = null;
                            }
                            else if ((entitiesCollection.Find(entitiesQuery).SetLimit(1).FirstOrDefault().Order_Type.Equals("Sell") ||
                                      entitiesCollection.Find(entitiesQuery).SetLimit(1).FirstOrDefault().Order_Type == "Sell") &&
                                     (trendCollection.Find(trendQuery).SetLimit(1).FirstOrDefault().Order_Type.Equals("Sell") ||
                                      trendCollection.Find(trendQuery).SetLimit(1).FirstOrDefault().Order_Type == "Sell"))
                            {
                                Console.WriteLine(GetTimestamp(DateTime.Now) + " : I am Selling !!!");
                                update_Order_Type_Table("Sell");
                                // Start TT API on a separate thread
                                // Call the SendOrder function with a Sell request.
                                workerThread = new Thread(Sell_Order_Dispatcher);
                                workerThread.Name = "Sell Order from Geolly";
                                workerThread.Start();
                                // Loop until worker thread activates. 
                                while (!workerThread.IsAlive) ;
                                // Put the main thread to sleep for 60000 millisecond to 
                                // allow the worker thread to do some work:
                                Thread.Sleep(60000);
                                // Request that the worker thread stop itself:
                                workerThread.Abort();
                                // Use the Join method to block the current thread  
                                // until the object's thread terminates.
                                workerThread.Join();
                                workerThread = null;
                                m_disp_sell = null;
                                m_req = null;
                                m_ps = null;
                                m_ts = null;
                                m_orderKey = null;
                                m_orderType = null;
                            }
                        }
                        else
                        {
                            if ((ordersCollection.Find(ordersQuery).SetLimit(1).FirstOrDefault().Filled_Order == "Buy" || ordersCollection.Find(ordersQuery).SetLimit(1).FirstOrDefault().Filled_Order.Equals("Buy")) && 
                                (takeProfitCollection.Find(takeProfitQuery).SetLimit(1).FirstOrDefault().Order_Type == "Take Profit" || takeProfitCollection.Find(takeProfitQuery).SetLimit(1).FirstOrDefault().Order_Type.Equals("Take Profit")))
                            {
                                Console.WriteLine(GetTimestamp(DateTime.Now) + " : I am Selling to take Profit!!!");
                                flush_Order_Tables("No Order");
                                // Start TT API on a separate thread
                                // Call the SendOrder function with a Sell request.
                                workerThread = new Thread(Sell_Order_Dispatcher);
                                workerThread.Name = "Sell Order from Geolly";
                                workerThread.Start();
                                // Loop until worker thread activates. 
                                while (!workerThread.IsAlive) ;
                                // Put the main thread to sleep for 60000 millisecond to 
                                // allow the worker thread to do some work:
                                Thread.Sleep(60000);
                                // Request that the worker thread stop itself:
                                workerThread.Abort();
                                // Use the Join method to block the current thread  
                                // until the object's thread terminates.
                                workerThread.Join();
                                workerThread = null;
                                m_disp_sell = null;
                                m_req = null;
                                m_ps = null;
                                m_ts = null;
                                m_orderKey = null;
                                m_orderType = null;
                            }
                            if ((ordersCollection.Find(ordersQuery).SetLimit(1).FirstOrDefault().Filled_Order == "Buy" || ordersCollection.Find(ordersQuery).SetLimit(1).FirstOrDefault().Filled_Order.Equals("Buy")) &&
                                (cutLossCollection.Find(cutLossQuery).SetLimit(1).FirstOrDefault().Order_Type == "Cut Loss" || cutLossCollection.Find(cutLossQuery).SetLimit(1).FirstOrDefault().Order_Type.Equals("Cut Loss")))
                            {
                                Console.WriteLine(GetTimestamp(DateTime.Now) + " : I am Selling to cut Loss!!!");
                                flush_Order_Tables("No Order");
                                // Start TT API on a separate thread
                                // Call the SendOrder function with a Sell request.
                                workerThread = new Thread(Sell_Order_Dispatcher);
                                workerThread.Name = "Sell Order from Geolly";
                                workerThread.Start();
                                // Loop until worker thread activates. 
                                while (!workerThread.IsAlive) ;
                                // Put the main thread to sleep for 60000 millisecond to 
                                // allow the worker thread to do some work:
                                Thread.Sleep(60000);
                                // Request that the worker thread stop itself:
                                workerThread.Abort();
                                // Use the Join method to block the current thread  
                                // until the object's thread terminates.
                                workerThread.Join();
                                workerThread = null;
                                m_disp_sell = null;
                                m_req = null;
                                m_ps = null;
                                m_ts = null;
                                m_orderKey = null;
                                m_orderType = null;
                            }
                            if ((ordersCollection.Find(ordersQuery).SetLimit(1).FirstOrDefault().Filled_Order == "Sell" || ordersCollection.Find(ordersQuery).SetLimit(1).FirstOrDefault().Filled_Order.Equals("Sell")) &&
                                (takeProfitCollection.Find(takeProfitQuery).SetLimit(1).FirstOrDefault().Order_Type == "Take Profit" || takeProfitCollection.Find(takeProfitQuery).SetLimit(1).FirstOrDefault().Order_Type.Equals("Take Profit")))
                            {
                                Console.WriteLine(GetTimestamp(DateTime.Now) + " : I am Buying to take Profit!!!");
                                flush_Order_Tables("No Order");
                                // Start TT API on a separate thread
                                // Call the SendOrder function with a Buy request.
                                workerThread = new Thread(Buy_Order_Dispatcher);
                                workerThread.Name = "Buy Order from Geolly";
                                workerThread.Start();
                                // Loop until worker thread activates. 
                                while (!workerThread.IsAlive) ;
                                // Put the main thread to sleep for 60000 millisecond to 
                                // allow the worker thread to do some work:
                                Thread.Sleep(60000);
                                // Request that the worker thread stop itself:
                                workerThread.Abort();
                                // Use the Join method to block the current thread  
                                // until the object's thread terminates.
                                workerThread.Join();
                                workerThread = null;
                                m_disp_buy = null;
                                m_req = null;
                                m_ps = null;
                                m_ts = null;
                                m_orderKey = null;
                                m_orderType = null;
                            }
                            if ((ordersCollection.Find(ordersQuery).SetLimit(1).FirstOrDefault().Filled_Order == "Sell" || ordersCollection.Find(ordersQuery).SetLimit(1).FirstOrDefault().Filled_Order.Equals("Sell")) &&
                                (cutLossCollection.Find(cutLossQuery).SetLimit(1).FirstOrDefault().Order_Type == "Cut Loss" || cutLossCollection.Find(cutLossQuery).SetLimit(1).FirstOrDefault().Order_Type.Equals("Cut Loss")))
                            {
                                Console.WriteLine(GetTimestamp(DateTime.Now) + " : I am Buying to cut Loss!!!");
                                flush_Order_Tables("No Order");
                                // Start TT API on a separate thread
                                // Call the SendOrder function with a Buy request.
                                workerThread = new Thread(Buy_Order_Dispatcher);
                                workerThread.Name = "Buy Order from Geolly";
                                workerThread.Start();
                                // Loop until worker thread activates. 
                                while (!workerThread.IsAlive) ;
                                // Put the main thread to sleep for 60000 millisecond to 
                                // allow the worker thread to do some work:
                                Thread.Sleep(60000);
                                // Request that the worker thread stop itself:
                                workerThread.Abort();
                                // Use the Join method to block the current thread  
                                // until the object's thread terminates.
                                workerThread.Join();
                                workerThread = null;
                                m_disp_buy = null;
                                m_req = null;
                                m_ps = null;
                                m_ts = null;
                                m_orderKey = null;
                                m_orderType = null;
                            }
                        }
                     }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Something went wrong : {0}", ex.Message);                  
                    Dispose();
                }

/*
                m_orderType = "Buy";
                // lookup an instrument
                m_req = new InstrumentLookupSubscription(m_apiInstance.Session, Dispatcher.Current,
                    new ProductKey(MarketKey.Cme, ProductType.Future, "ES"),
                    "Dec14");
                m_req.Update += new EventHandler<InstrumentLookupSubscriptionEventArgs>(m_req_Update);
                m_req.Start();
*/
            }
            else
            {
                Console.WriteLine("TT Login failed: {0}", e.Status.StatusMessage);
                Dispose();
            }
        }

        /// <summary>
        /// Create and start the Dispatcher for Buy Order Threads
        /// </summary>
        private void Buy_Order_Dispatcher()
        {
            // Attach a WorkerDispatcher to the current thread
            m_disp_buy = Dispatcher.AttachWorkerDispatcher();
            m_disp_buy.BeginInvoke(new Action(Buy_Order));
            m_disp_buy.Run();
        }

        /// <summary>
        /// Create and start the Dispatcher for Sell Order Threads
        /// </summary>
        private void Sell_Order_Dispatcher()
        {
            // Attach a WorkerDispatcher to the current thread
            m_disp_sell = Dispatcher.AttachWorkerDispatcher();
            m_disp_sell.BeginInvoke(new Action(Sell_Order));
            m_disp_sell.Run();
        }

        public void Buy_Order()
        {
            var connectionString = "mongodb://localhost";
            var client = new MongoClient(connectionString);
            var server = client.GetServer();
            var database = server.GetDatabase("test");
            var accountCollection = database.GetCollection<Account>("account");
            var accountQuery = Query<Account>.EQ(account => account.Start_Trading, true);

            // Call the SendOrder function with a Buy request.
            m_orderKey = "";
            m_orderType = "Buy";

            // lookup an instrument
            /*
            m_req = new InstrumentLookupSubscription(m_apiInstance.Session, Dispatcher.Current,
                new ProductKey(MarketKey.Cme, ProductType.Future, "ES"),
                "Dec14");
            */
            m_req = new InstrumentLookupSubscription(m_apiInstance.Session, Dispatcher.Current,
                new ProductKey(MarketKey.Cme, ProductType.Future, accountCollection.Find(accountQuery).SetLimit(1).FirstOrDefault().Product_Name),
                accountCollection.Find(accountQuery).SetLimit(1).FirstOrDefault().Instrument_Name);
            m_req.Update += new EventHandler<InstrumentLookupSubscriptionEventArgs>(m_req_Update);
            m_req.Start();

        }

        private void Sell_Order()
        {
            var connectionString = "mongodb://localhost";
            var client = new MongoClient(connectionString);
            var server = client.GetServer();
            var database = server.GetDatabase("test");
            var accountCollection = database.GetCollection<Account>("account");
            var accountQuery = Query<Account>.EQ(account => account.Start_Trading, true);

            // Call the SendOrder function with a Sell request.
            m_orderKey = "";
            m_orderType = "Sell";

            // lookup an instrument
            /*
            m_req = new InstrumentLookupSubscription(m_apiInstance.Session, Dispatcher.Current,
                new ProductKey(MarketKey.Cme, ProductType.Future, "ES"),
                "Dec14");
            */
            m_req = new InstrumentLookupSubscription(m_apiInstance.Session, Dispatcher.Current,
                new ProductKey(MarketKey.Cme, ProductType.Future, accountCollection.Find(accountQuery).SetLimit(1).FirstOrDefault().Product_Name),
                accountCollection.Find(accountQuery).SetLimit(1).FirstOrDefault().Instrument_Name);
            m_req.Update += new EventHandler<InstrumentLookupSubscriptionEventArgs>(m_req_Update);
            m_req.Start();

        }
      
        private static String GetTimestamp(DateTime value)
        {
            return value.ToString("yyyy:MM:dd:HH:mm:ss:ffff");
        }

        void update_Order_Type_Table(String orderType)
        {
            var connectionString = "mongodb://localhost";
            var client = new MongoClient(connectionString);
            var server = client.GetServer();
            var database = server.GetDatabase("test");
            var collection = database.GetCollection<Entity>("orders");
            var query = Query<Order>.EQ(order => order.Start_Trading, true);
            var update = Update<Order>.Set(order => order.Filled_Order, orderType); // update modifiers
            collection.Update(query, update);
        }

        void flush_Order_Tables(String orderType)
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

            var entitiesUpdate = Update<Entity>.Set(entity => entity.Order_Type, orderType); // update modifiers
            var trendUpdate = Update<Trend>.Set(trend => trend.Order_Type, orderType); // update modifiers
            var takeProfitUpdate = Update<TakeProfit>.Set(takeProfit => takeProfit.Order_Type, orderType); // update modifiers
            var cutLossUpdate = Update<CutLoss>.Set(cutLoss => cutLoss.Order_Type, orderType); // update modifiers
            var ordersUpdate = Update<Order>.Set(order => order.Filled_Order, orderType); // update modifiers

            entitiesCollection.Update(entitiesQuery, entitiesUpdate);
            trendCollection.Update(trendQuery, trendUpdate);
            takeProfitCollection.Update(takeProfitQuery, takeProfitUpdate);
            cutLossCollection.Update(cutLossQuery, cutLossUpdate);
            ordersCollection.Update(ordersQuery, ordersUpdate);
        }
              
        void update_Order_Executed_Table(Boolean orderExecuted)
        {
            var connectionString = "mongodb://localhost";
            var client = new MongoClient(connectionString);
            var server = client.GetServer();
            var database = server.GetDatabase("test");
            var collection = database.GetCollection<Entity>("orders");
            var query = Query<Order>.EQ(order => order.Start_Trading, true);
            var update = Update<Order>.Set(order => order.Order_Executed, orderExecuted); // update modifiers
            collection.Update(query, update);
        }

        /// <summary>
        /// Event notification for instrument lookup
        /// </summary>
        void m_req_Update(object sender, InstrumentLookupSubscriptionEventArgs e)
        {
            if (e.Instrument != null && e.Error == null)
            {
                // Instrument was found
                Console.WriteLine("Found: {0}", e.Instrument.Name);

                // Subscribe for Inside Market Data
                m_ps = new PriceSubscription(e.Instrument, Dispatcher.Current);
                m_ps.Settings = new PriceSubscriptionSettings(PriceSubscriptionType.InsideMarket);
                m_ps.FieldsUpdated += new FieldsUpdatedEventHandler(m_ps_FieldsUpdated);
                m_ps.Start();

                // Create a TradeSubscription to listen for order / fill events only for orders submitted through it
                m_ts = new InstrumentTradeSubscription(m_apiInstance.Session, Dispatcher.Current, e.Instrument, true, true, false, false);
                m_ts.OrderUpdated += new EventHandler<OrderUpdatedEventArgs>(m_ts_OrderUpdated);
                m_ts.OrderAdded += new EventHandler<OrderAddedEventArgs>(m_ts_OrderAdded);
                m_ts.OrderDeleted += new EventHandler<OrderDeletedEventArgs>(m_ts_OrderDeleted);
                m_ts.OrderFilled += new EventHandler<OrderFilledEventArgs>(m_ts_OrderFilled);
                m_ts.OrderRejected += new EventHandler<OrderRejectedEventArgs>(m_ts_OrderRejected);
                m_ts.Start();
            }
            else if (e.IsFinal)
            {
                // Instrument was not found and TT API has given up looking for it
                Console.WriteLine("Cannot find instrument: {0}", e.Error.Message);
                Dispose();
            }
        }

        /// <summary>
        /// Event notification for price update
        /// </summary>
        void m_ps_FieldsUpdated(object sender, FieldsUpdatedEventArgs e)
        {
            if (e.Error == null)
            {
                    // Make sure that there is a valid bid
                    if (e.Fields.GetBestBidPriceField().HasValidValue)
                    {
                        if (m_orderKey == "")
                        {
                            // If there is no order working, submit one through the first valid order feed.
                            // You should use the order feed that is valid for your purposes.
                            OrderProfile op = new OrderProfile(e.Fields.Instrument.GetValidOrderFeeds()[0], e.Fields.Instrument);
                            if (m_orderType == "Buy")
                            {
                                op.BuySell = BuySell.Buy;
                            }
                            else if (m_orderType == "Sell")
                            {
                                op.BuySell = BuySell.Sell;
                            }
                            else
                            {
                                op.BuySell = BuySell.Unknown;
                            }
                            var connectionString = "mongodb://localhost";
                            var client = new MongoClient(connectionString);
                            var server = client.GetServer();
                            var database = server.GetDatabase("test");
                            var accountCollection = database.GetCollection<Account>("account");
                            var accountQuery = Query<Account>.EQ(account => account.Start_Trading, true);
                            var ordersCollection = database.GetCollection<Order>("orders");
                            var ordersQuery = Query<Order>.EQ(order => order.Start_Trading, true);

                            //op.AccountName = "15610035";
                            op.AccountName = accountCollection.Find(accountQuery).SetLimit(1).FirstOrDefault().Account_Name;
                            op.AccountType = AccountType.A1;
                            op.OrderQuantity = Quantity.FromInt(e.Fields.Instrument, 1);
                            op.OrderType = OrderType.Limit;
                            op.LimitPrice = e.Fields.GetBestBidPriceField().Value;

                            if (!m_ts.SendOrder(op))
                            {
                                update_Order_Executed_Table(false);
                                Console.WriteLine("Send new order failed.  {0}", op.RoutingStatus.Message);
                                Dispose();
                            }
                            else
                            {
                                m_orderKey = op.SiteOrderKey;
                                update_Order_Executed_Table(true);
                                Console.WriteLine("Send new order succeeded." + m_orderKey);
                            }
                        }
                        else if (m_ts.Orders.ContainsKey(m_orderKey) &&
                            m_ts.Orders[m_orderKey].LimitPrice != e.Fields.GetBestBidPriceField().Value)
                        {
                            // If there is a working order, reprice it if its price is not the same as the bid
                            OrderProfileBase op = m_ts.Orders[m_orderKey].GetOrderProfile();
                            op.LimitPrice = e.Fields.GetBestBidPriceField().Value;
                            op.Action = OrderAction.Change;

                            if (!m_ts.SendOrder(op))
                            {
                                Console.WriteLine("Send change order failed.  {0}", op.RoutingStatus.Message);
                            }
                            else
                            {
                                Console.WriteLine("Send change order succeeded.");
                            }
                        }
                    }
            }
            else
            {
                if (e.Error.IsRecoverableError == false)
                {
                    Console.WriteLine("Unrecoverable price subscription error: {0}", e.Error.Message);
                    Dispose();
                }
            }
        }

        /// <summary>
        /// Event notification for order rejected
        /// </summary>
        void m_ts_OrderRejected(object sender, OrderRejectedEventArgs e)
        {
            Console.WriteLine("Order was rejected.");
        }

        /// <summary>
        /// Event notification for order filled
        /// </summary>
        void m_ts_OrderFilled(object sender, OrderFilledEventArgs e)
        {
            if (e.FillType == FillType.Full)
            {
                Console.WriteLine("Order was fully filled for {0} at {1}.", e.Fill.Quantity, e.Fill.MatchPrice);
            }
            else
            {
                Console.WriteLine("Order was partially filled for {0} at {1}.", e.Fill.Quantity, e.Fill.MatchPrice);
            }

            Console.WriteLine("Average Buy Price = {0} : Net Position = {1} : P&L = {2}", m_ts.ProfitLossStatistics.BuyAveragePrice,
                m_ts.ProfitLossStatistics.NetPosition, m_ts.ProfitLoss.AsPrimaryCurrency);
        }

        /// <summary>
        /// Event notification for order deleted
        /// </summary>
        void m_ts_OrderDeleted(object sender, OrderDeletedEventArgs e)
        {
            Console.WriteLine("Order was deleted.");
        }

        /// <summary>
        /// Event notification for order added
        /// </summary>
        void m_ts_OrderAdded(object sender, OrderAddedEventArgs e)
        {
            Console.WriteLine("Order was added with price of {0}.", e.Order.LimitPrice);
        }

        /// <summary>
        /// Event notification for order update
        /// </summary>
        void m_ts_OrderUpdated(object sender, OrderUpdatedEventArgs e)
        {
            Console.WriteLine("Order was updated with price of {0}.", e.NewOrder.LimitPrice);
        }

        /// <summary>
        /// Shuts down the TT API
        /// </summary>
        public void Dispose()
        {
            lock (m_lock)
            {
                if (!m_disposed)
                {
                    // Unattached callbacks and dispose of all subscriptions
                    if (m_req != null)
                    {
                        m_req.Update -= m_req_Update;
                        m_req.Dispose();
                        m_req = null;
                    }
                    if (m_ps != null)
                    {
                        m_ps.FieldsUpdated -= m_ps_FieldsUpdated;
                        m_ps.Dispose();
                        m_ps = null;
                    }
                    if (m_ts != null)
                    {
                        m_ts.OrderUpdated -= m_ts_OrderUpdated;
                        m_ts.OrderAdded -= m_ts_OrderAdded;
                        m_ts.OrderDeleted -= m_ts_OrderDeleted;
                        m_ts.OrderFilled -= m_ts_OrderFilled;
                        m_ts.OrderRejected -= m_ts_OrderRejected;
                        m_ts.Dispose();
                        m_ts = null;
                    }

                    // Begin shutdown the TT API
                    TTAPI.ShutdownCompleted += new EventHandler(TTAPI_ShutdownCompleted);
                    TTAPI.Shutdown();

                    m_disposed = true;
                }
            }
        }

        /// <summary>
        /// Event notification for completion of TT API shutdown
        /// </summary>
        public void TTAPI_ShutdownCompleted(object sender, EventArgs e)
        {
            // Shutdown the Dispatcher
            if (m_disp != null)
            {
                m_disp.BeginInvokeShutdown();
                m_disp = null;
            }
            if (m_disp_buy != null)
            {
                m_disp_buy.BeginInvokeShutdown();
                m_disp_buy = null;
            }
            if (m_disp_sell != null)
            {
                m_disp_sell.BeginInvokeShutdown();
                m_disp_sell = null;
            }

            // Dispose of any other objects / resources
            if (workerThread != null)
            {
                // Request that the worker thread stop itself:
                workerThread.Abort();
                // Use the Join method to block the current thread  
                // until the object's thread terminates.
                workerThread.Join();
                workerThread = null;
            }
            m_orderKey = null;
            m_orderType = null;
        }
    }
}
