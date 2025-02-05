using System;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using Newtonsoft.Json;
using System.Text;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace QuickBooksCRUD
{
    public class Program
    {
        public static async Task  Main(string[] args)
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
            string queueName = "ItemQueue";

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
                    // Deserialize dynamically into a dictionary
                    var queueData = JsonConvert.DeserializeObject<Dictionary<string, List<ItemModel>>>(message);


                    if (queueData != null)
                    {
                        Console.WriteLine("Received data:");

                        //foreach (var category in queueData)
                        //{
                        //    Console.WriteLine($"{category.Key}:");
                        //    foreach (var item in category.Value)
                        //    {
                        //        Console.WriteLine($"  {item.Item}: {item.Price}");
                        //    }
                        //}


                        QuickBooks quickBooks = new QuickBooks();
                        // You can process each category individually, for example:
                        quickBooks.DoInvoiceAdd(queueData);

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


