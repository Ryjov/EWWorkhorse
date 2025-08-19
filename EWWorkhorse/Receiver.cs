using System.Text;

namespace EWWorkhorse;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;

public class Receiver
{
    public static async Task Main()
    {
        ConnectionFactory factory = new();
        factory.Uri = new Uri(uriString: "amqp://guest:guest@localhost:1011");
        factory.ClientProvidedName = "Rabbit receiver app";

        try
        {
            var cnn = await factory.CreateConnectionAsync();
            var channel = await cnn.CreateChannelAsync();

            string exchangeName = "EWExchange";
            string routingKey = "ew-routing-key";
            string queueName = "EWFileQueue";

            await channel.ExchangeDeclareAsync(exchangeName, ExchangeType.Direct);
            await channel.QueueDeclareAsync(queueName, durable: false, exclusive: false, autoDelete: false, arguments: null);
            await channel.QueueBindAsync(queueName, exchangeName, routingKey, arguments: null);
            await channel.BasicQosAsync(prefetchSize: 0, prefetchCount: 2, global: false);

            var consumer = new AsyncEventingBasicConsumer(channel);
            consumer.ReceivedAsync += async (sender, args) =>
            {
                var body = args.Body.ToArray();
                //Console.WriteLine("Message Received: " + Encoding.UTF8.GetString(body));

                //var worker = new Worker();// todo
                //worker.ExecuteTask;// todo

                await channel.BasicAckAsync(args.DeliveryTag, multiple: true);
            };

            string consumerTag = await channel.BasicConsumeAsync(queueName, autoAck: false, consumer);

            Console.ReadLine();

            await channel.BasicCancelAsync(consumerTag);

            await channel.CloseAsync();
            await cnn.CloseAsync();
        }
        catch (Exception ex)
        {
            string s = ex.Message;
        }
    }
}