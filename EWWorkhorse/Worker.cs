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
        private const string _queueName = "EWFileQueue";
        private const string _exchangeName = "EWExchange";
        private const string _routingKey = "ew-routing-key";
        private const string _deadLetterExchangeName = "EWFileDLX";
        private const string _deadLetterQueueName = "EWFileDLQ";
        private readonly ILogger<Worker> _logger;
        CancellationToken _cancellationToken;

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

                await channel.ExchangeDeclareAsync(_exchangeName, ExchangeType.Direct);
                await channel.ExchangeDeclareAsync(_deadLetterExchangeName, ExchangeType.Direct);

                await channel.QueueDeclareAsync(_queueName, durable: false, exclusive: false, autoDelete: false,
                    arguments: new Dictionary<string, object?> { { "x-dead-letter-exchange", _deadLetterExchangeName } });
                await channel.QueueDeclareAsync(_deadLetterQueueName, durable: false, exclusive: false, autoDelete: false, arguments: null);

                await channel.QueueBindAsync(_queueName, _exchangeName, _routingKey, arguments: null);
                await channel.QueueBindAsync(_deadLetterQueueName, _deadLetterExchangeName, _queueName);

                await channel.BasicQosAsync(prefetchSize: 0, prefetchCount: 1, global: false);

                var consumer = new AsyncEventingBasicConsumer(channel);

                consumer.ReceivedAsync += async (sender, args) =>
                {
                    Console.WriteLine("Job Received. Delivery tag: " + args.DeliveryTag);

                    var props = args.BasicProperties;

                    if (props.ReplyTo is null)
                    {
                        await channel.BasicNackAsync(args.DeliveryTag, multiple: false, requeue: false);
                        return;
                    }

                    var replyProps = new BasicProperties { CorrelationId = props.CorrelationId };

                    try
                    {
                        (byte[] wordBytes, byte[] excelBytes) = ParseEnvelope(args.Body.ToArray());
                        byte[] resultBytes = Replacer.ReplaceBytes(wordBytes, excelBytes);
                        await channel.BasicPublishAsync(string.Empty, props.ReplyTo, true, replyProps, resultBytes);
                        await channel.BasicAckAsync(args.DeliveryTag, multiple: false);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Unexpected error: " + ex.Message);

                        try
                        {
                            await channel.BasicPublishAsync(string.Empty, props.ReplyTo, true, new BasicProperties { CorrelationId = props.CorrelationId, Type = "error" },
                                Encoding.UTF8.GetBytes(ex.Message));
                        }
                        catch (Exception replyEx)
                        {
                            Console.WriteLine("Failed to send error reply: " + replyEx.Message);
                        }

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
