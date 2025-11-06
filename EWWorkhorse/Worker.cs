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
                //await channel.ExchangeDeclareAsync(string.Empty, ExchangeType.Direct);
                await channel.QueueDeclareAsync(_queueName, durable: false, exclusive: false, autoDelete: false, arguments: null);
                await channel.QueueBindAsync(_queueName, exchangeName, routingKey, arguments: null);
                await channel.BasicQosAsync(prefetchSize: 0, prefetchCount: 2, global: false);

                var consumer = new AsyncEventingBasicConsumer(channel);
                while (!stoppingToken.IsCancellationRequested)
                {
                    if (_logger.IsEnabled(LogLevel.Information))
                    {
                        _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
                    }
                    //await Task.Delay(5000, stoppingToken);

                    var wordBody = default(byte[]);
                    var excelBody = default(byte[]);

                    consumer.ReceivedAsync += async (sender, args) =>
                    {
                        if (args.DeliveryTag % 2 == 0)
                            excelBody = args.Body.ToArray();
                        else
                            wordBody = args.Body.ToArray();

                        Console.WriteLine("Message Received. Delivery tag: " + args.DeliveryTag);

                        IReadOnlyBasicProperties props = args.BasicProperties;
                        var replyProps = new BasicProperties
                        {
                            CorrelationId = props.CorrelationId
                        };

                        if (wordBody != null && excelBody != null)
                        {
                            try
                            {
                                using (MemoryStream excelMeme = new MemoryStream())
                                using (MemoryStream wordMeme = new MemoryStream())
                                {
                                    wordMeme.Write(wordBody, 0, (int)wordBody.Length);
                                    WordprocessingDocument wordDoc = WordprocessingDocument.Open(wordMeme, true);

                                    excelMeme.Write(excelBody, 0, (int)excelBody.Length);
                                    SpreadsheetDocument excDoc = SpreadsheetDocument.Open(excelMeme, true);

                                    var result = Replacer.ReplaceFile(wordDoc, excDoc);
                                    var wordBytes = default(byte[]);
                                    using (StreamReader sr = new StreamReader(result.MainDocumentPart.GetStream()))
                                    {
                                        //var w = await sr.ReadToEndAsync();
                                        using (var memstream = new MemoryStream())
                                        {
                                            sr.BaseStream.CopyTo(memstream);
                                            wordBytes = memstream.ToArray();
                                        }
                                    }
                                    await channel.BasicPublishAsync(string.Empty, props.ReplyTo!, true, replyProps, wordBytes);

                                    await channel.BasicAckAsync(args.DeliveryTag, multiple: true);
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Unexpected error: " + ex.Message);
                            }
                            finally
                            {
                                wordBody = default(byte[]);
                                excelBody = default(byte[]);
                            }
                        }

                    };

                    string consumerTag = await channel.BasicConsumeAsync(_queueName, false, consumer);

                    Console.ReadLine();

                    //??
                    await channel.BasicCancelAsync(consumerTag);

                    await channel.CloseAsync();
                    await cnn.CloseAsync();
                }
            }
            catch (Exception ex)
            {
                string s = ex.Message;
            }
        }
    }
}
