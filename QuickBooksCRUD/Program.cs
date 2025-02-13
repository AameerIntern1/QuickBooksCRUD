using System;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using Newtonsoft.Json;
using System.Text;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Interop.QBFC16;

namespace QuickBooksCRUD
{
    public class Program
    {


        //static void Main(string[] args)
        //{
        //    try
        //    {
        //        // Create QBFC session manager
        //        QBSessionManager qbSessionManager = new QBSessionManager();
        //        qbSessionManager.OpenConnection("", "QuickBooks Invoice Sample");
        //        qbSessionManager.BeginSession("", ENOpenMode.omDontCare);

        //        // Create a new Invoice
        //        IInvoiceAdd invoiceAdd = (IInvoiceAdd)qbSessionManager.CreateMsgSetRequest("US", 13, 0).AppendInvoiceAddRq();

        //        // Set Customer Ref
        //        ICustomerRef customerRef = invoiceAdd.CustomerRef;
        //        customerRef.ListID.SetValue("1"); // Replace with your customer ListID

        //        // Set the Invoice Date
        //        invoiceAdd.TxnDate.SetValue(DateTime.Now.ToString("yyyy-MM-dd"));

        //        // Add an item to the Invoice
        //        IInvoiceLineAdd line = invoiceAdd.InvoiceLineAdd.Add();
        //        IItemRef itemRef = line.ItemRef;
        //        itemRef.ListID.SetValue("1"); // Replace with your item ListID
        //        line.Quantity.SetValue(2);
        //        line.Rate.SetValue(15.50); // Item price

        //        // Add the Invoice to QuickBooks
        //        IMsgSetResponse response = qbSessionManager.DoRequests(qbSessionManager.CreateMsgSetRequest("US", 13, 0));

        //        // Check the response status
        //        if (response.ResponseList.GetAt(0).StatusCode != 0)
        //        {
        //            Console.WriteLine("Error: " + response.ResponseList.GetAt(0).StatusMessage);
        //        }
        //        else
        //        {
        //            Console.WriteLine("Invoice created successfully!");
        //        }

        //        // End the session and close the connection
        //        qbSessionManager.EndSession();
        //        qbSessionManager.CloseConnection();
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Error: " + ex.Message);
        //    }
        //}
        public static async Task Main(string[] args)
        {
            QuickBooks quickBooks = new QuickBooks();
          

            //quickBooks.DoInvoiceAdd();
            //quickBooks.DoItemAdd(data);
            //quickBooks.GetAccount();
            //quickBooks.GetItems()
            //quickBooks.GetInvoices();
            //quickBooks.GetCompanyInfo();
            //quickBooks.GetCategory();
            //quickBooks.GetClasses();

            await RabbitMQ();

        }
        public static async Task RabbitMQ()
        {
            var factory = new ConnectionFactory() { HostName = "localhost" };
            var connection = await factory.CreateConnectionAsync();
            var channel = await connection.CreateChannelAsync();
            string queueName = "invoiceQueue";

            await channel.QueueDeclareAsync(
                queue: queueName,
                durable: true,
                exclusive: false,
                autoDelete: false,
                arguments: null);

            Console.WriteLine($"Waiting for messages in {queueName}...");

            var consumer = new AsyncEventingBasicConsumer(channel);
            consumer.ReceivedAsync += async (model, ea) =>
            {
                var body = ea.Body.ToArray();
                var message = Encoding.UTF8.GetString(body);

                try
                {
                    //var queueData = JsonConvert.DeserializeObject<Dictionary<string,decimal>>(message);
                     Dictionary<string, List<Journal>>?  queueData = JsonConvert.DeserializeObject<Dictionary <string,List<Journal>>>(message);

                    //Dictionary<string, decimal>? queueData = JsonConvert.DeserializeObject<Dictionary<string, decimal>>(message);


                    if (queueData != null)
                    {
                        //foreach (var journal in queueData) {
                        //    Console.WriteLine($"{journal.Key} ");
                        //    foreach (var data in journal.Value)
                        //    {
                        //        Console.WriteLine($"{data.Account}  {data.EarnedAmount}  {data.UnEarnedAmount}  {data.AccountReceivable}  {data.Cash}");
                        //    }
                        //}
                        Console.WriteLine("Received data:");
                        Console.WriteLine($"No of data Received :{queueData.Count}");

                        QuickBooks quickBooks = new QuickBooks();
                        var accountList = quickBooks.GetJournal();
                        quickBooks.DailyJournalAdd(queueData, accountList);
                        //var list = quickBooks.GetInvoices1(queueData);
                        //foreach (var mod in list)
                        //{
                        //    Console.WriteLine($"Modifying Invoice - TxnID: {mod.TaxId}, EditSequence: {mod.EditSequenceID}, Item: {mod.Item}, Old Price: {mod.OldPrice}, New Price: {mod.NewPrice}");
                        //}
                        //quickBooks.DailyInvoiceAdd(queueData,list);
                        //quickBooks.GetInvoices(queueData);



                        Console.WriteLine("Press Enter To Exit..");
                    }
                    else
                    {
                        Console.WriteLine("Failed to parse queue data.");
                    }
                }
                catch (JsonException ex)
                {
                    Console.WriteLine($"Error deserializing message: {ex.Message}");
                }

                Console.WriteLine(new string('*', 100));
            };

            await channel.BasicConsumeAsync(queue: queueName, autoAck: false, consumer: consumer);

            // To keep the program running
            Console.ReadLine();

            await channel.CloseAsync();
            await connection.CloseAsync();
        }



    }
}


