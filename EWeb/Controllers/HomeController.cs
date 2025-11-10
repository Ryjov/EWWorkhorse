using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using EWeb.Models;
using EWeb.RPC;
using Microsoft.AspNetCore.Connections;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Channels;
using System.Xml;
using System.Xml.Serialization;

namespace EWeb.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        IWebHostEnvironment _appEnvironment;
        CancellationToken _cancellationToken;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> AddFile(IFormFileCollection uploadedFiles)
        {
            var file = await InvokeAsync(uploadedFiles);
            var beginningTagLength = 1348;
            var endingTagLength = 22;
            file = file.Substring(1348, file.Length - beginningTagLength - endingTagLength);

            var stream = new MemoryStream();
            using (WordprocessingDocument doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
            {
                MainDocumentPart mainPart = doc.AddMainDocumentPart();

                new Document(new Body()).Save(mainPart);

                Body body = mainPart.Document.Body;
                body.InnerXml = file;

                mainPart.Document.Save();
            }
            stream.Seek(0, SeekOrigin.Begin);

            return File(stream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "result.docx");
        }

        [HttpGet]
        public async Task GetResult()//FileStreamResult
        {
            ConnectionFactory factory = new();
            factory.Uri = new Uri(uriString: "amqp://guest:guest@localhost:1011");
            factory.ClientProvidedName = "Rabbit receiver app";


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
            while (!_cancellationToken.IsCancellationRequested)
            {
                consumer.ReceivedAsync += async (sender, args) =>
                {
                    await channel.BasicAckAsync(args.DeliveryTag, multiple: true);
                    var wordBody = args.Body.ToArray();
                };
            }
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        private static async Task<string> InvokeAsync(IFormFileCollection files)
        {
            var rpcClient = new RPCClient();
            await rpcClient.StartAsync();
            var resultFile = await rpcClient.CallAsync(files);

            return resultFile;
        }
    }
}