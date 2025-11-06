using DocumentFormat.OpenXml.Packaging;
using EWeb.Models;
using EWeb.RPC;
using Microsoft.AspNetCore.Connections;
using Microsoft.AspNetCore.Mvc;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Channels;

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

            using (var memStream = new MemoryStream())
            {
                WordprocessingDocument wordDoc = WordprocessingDocument.Create(memStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
                wordDoc.AddMainDocumentPart();

                //get the main part of the document which contains CustomXMLParts
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;

                //delete all CustomXMLParts in the document. If needed only specific CustomXMLParts can be deleted using the CustomXmlParts IEnumerable
                mainPart.DeleteParts<CustomXmlPart>(mainPart.CustomXmlParts);

                byte[] buf = (new UTF8Encoding()).GetBytes(file);
                memStream.Write(buf, 0, buf.Length);
                //add new CustomXMLPart with data from new XML file
                CustomXmlPart myXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                myXmlPart.FeedData(memStream);

                wordDoc.Save();
                var result = memStream.ToArray();

                return File(result, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "result.docx");
            }
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