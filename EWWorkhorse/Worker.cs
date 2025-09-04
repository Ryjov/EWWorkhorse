using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using System.Collections;
using System.Text;
using System.Text.RegularExpressions;

namespace EWWorkhorse
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
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
                while (!stoppingToken.IsCancellationRequested)
                {
                    if (_logger.IsEnabled(LogLevel.Information))
                    {
                        _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
                    }
                    await Task.Delay(5000, stoppingToken);

                    var wordBody = default(byte[]);
                    var excelBody = default(byte[]);

                    consumer.ReceivedAsync += async (sender, args) =>
                    {
                        if (args.DeliveryTag % 2 == 0)
                            excelBody = args.Body.ToArray();
                        else
                            wordBody = args.Body.ToArray();

                        Console.WriteLine("Message Received. Delivery tag: " + args.DeliveryTag);

                        // todo: add Replace

                        await channel.BasicAckAsync(args.DeliveryTag, multiple: true);

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

                                    var result = await Replace(wordDoc, excDoc);
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

                    string consumerTag = await channel.BasicConsumeAsync(queueName, autoAck: false, consumer);

                    Console.ReadLine();

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

        public async Task<WordprocessingDocument> Replace(WordprocessingDocument wdoc, SpreadsheetDocument excDoc)
        {
            var wordBody = wdoc.MainDocumentPart.Document.Body;
            var paragraphs = wordBody.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();
            Regex markerRegEx = new Regex(@"<#\d+#[A-Z]+\d+>");

            foreach (var paragraph in paragraphs)
            {
                foreach (var run in paragraph.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>())
                {
                    foreach (var text in run.Elements<DocumentFormat.OpenXml.Wordprocessing.Text>())
                    {
                        MatchCollection markerMatches = markerRegEx.Matches(text.Text);

                        foreach (Match match in markerMatches)
                        {
                            Regex sheetRegEx = new Regex(@"#\d+#");
                            Regex cellRegEx = new Regex(@"#[A-Z]+\d+>");
                            int sheetIndex = Int32.Parse(sheetRegEx.Match(match.Value).Value.Trim('#'));
                            string cellIndex = cellRegEx.Match(match.Value).Value.Trim('#', '>');
                            WorkbookPart wbPart = excDoc.WorkbookPart;
                            Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.SheetId == sheetIndex);
                            WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
                            Cell cell = wsPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == cellIndex);

                            var value = cell.InnerText;

                            if (cell.DataType is not null)
                            {
                                if (cell.DataType.Value == CellValues.SharedString)
                                {
                                    var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                                    value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;

                                    text.Text = text.Text.Replace(match.Value, value);
                                }
                            }
                            else
                            {
                                text.Text = text.Text.Replace(match.Value, value);
                            }
                        }
                    }
                }
            }

            wdoc.Save();
            return wdoc;
        }
    }
}
