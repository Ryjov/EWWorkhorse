using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using System.Collections;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace EWWorkhorse
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        CancellationToken _cancellationToken;
        const string _queueName = "EWFileQueue";

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            ConnectionFactory factory = new();
            factory.Uri = new Uri(uriString: "amqp://guest:guest@localhost:1011");
            factory.ClientProvidedName = "Rabbit receiver";

            try
            {
                var cnn = await factory.CreateConnectionAsync();
                var channel = await cnn.CreateChannelAsync();

                string exchangeName = "EWExchange";
                string routingKey = "ew-routing-key";

                await channel.ExchangeDeclareAsync(exchangeName, ExchangeType.Direct);
                await channel.QueueDeclareAsync(_queueName, durable: false, exclusive: false, autoDelete: false, arguments: null);
                await channel.QueueBindAsync(_queueName, exchangeName, routingKey, arguments: null);
                await channel.BasicQosAsync(prefetchSize: 0, prefetchCount: 1, global: false);

                var consumer = new AsyncEventingBasicConsumer(channel);

                consumer.ReceivedAsync += async (sender, args) =>
                {
                    Console.WriteLine("Job Received. Delivery tag: " + args.DeliveryTag);

                    var props = args.BasicProperties;
                    var replyProps = new BasicProperties { CorrelationId = props.CorrelationId };

                    try
                    {
                        (byte[] wordBytes, byte[] excelBytes) = ParseEnvelope(args.Body.ToArray());
                        byte[] resultBytes = Replacer.ReplaceBytes(wordBytes, excelBytes);
                        await channel.BasicPublishAsync(string.Empty, props.ReplyTo!, true, replyProps, resultBytes);
                        await channel.BasicAckAsync(args.DeliveryTag, multiple: false);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Unexpected error: " + ex.Message);
                        await channel.BasicNackAsync(args.DeliveryTag, multiple: false, requeue: false);
                    }
                };

                string consumerTag = await channel.BasicConsumeAsync(_queueName, false, consumer);

                try { await Task.Delay(Timeout.Infinite, stoppingToken); }
                catch (OperationCanceledException) { }
                finally
                {
                    await channel.CloseAsync();
                    await cnn.CloseAsync();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Worker failed");
            }
        }
        private static (byte[] wordBytes, byte[] excelBytes) ParseEnvelope(byte[] body)
        {
            int wordLength = BitConverter.ToInt32(body, 0);
            return (body.AsSpan(4, wordLength).ToArray(), body[(4 + wordLength)..]);
        }
    }
}
