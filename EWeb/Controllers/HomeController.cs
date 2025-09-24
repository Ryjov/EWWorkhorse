using DocumentFormat.OpenXml.Packaging;
using EWeb.Models;
using Microsoft.AspNetCore.Connections;
using Microsoft.AspNetCore.Mvc;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using System.Diagnostics;
using System.IO;
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
            ConnectionFactory factory = new();
            factory.Uri = new Uri(uriString: "amqp://guest:guest@localhost:1011");// add appsettings
            factory.ClientProvidedName = "EW filebytes sender app";

            var cnn = await factory.CreateConnectionAsync();
            var channel = await cnn.CreateChannelAsync();

            string exchangeName = "EWExchange";
            string routingKey = "ew-routing-key";
            string queueName = "EWFileQueue";

            channel.ExchangeDeclareAsync(exchangeName, ExchangeType.Direct);
            channel.QueueDeclareAsync(queueName, durable: false, exclusive: false, autoDelete: false, arguments: null);
            channel.QueueBindAsync(queueName, exchangeName, routingKey, arguments: null);

            byte[] wordBytes = default(byte[]);
            byte[] excDoc = default(byte[]);

            foreach (var uploadedFile in uploadedFiles)
            {
                if (uploadedFile.ContentType == "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                {
                    using (var reader = new StreamReader(uploadedFile.OpenReadStream()))
                    {
                        using (var mem = new MemoryStream())
                        {
                            reader.BaseStream.CopyTo(mem);
                            wordBytes = mem.ToArray();
                        }
                    }
                }
                else if (uploadedFile.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                {
                    using (var reader = new StreamReader(uploadedFile.OpenReadStream()))
                    {
                        using (var mem = new MemoryStream())
                        {
                            reader.BaseStream.CopyTo(mem);
                            excDoc = mem.ToArray();
                        }
                    }
                }
            }

            //var fileBatch = channel.CreateBasicPublishBatch();
            await channel.BasicPublishAsync(exchangeName, routingKey, true, wordBytes, _cancellationToken);// need to roll into one?
            await channel.BasicPublishAsync(exchangeName, routingKey, true, excDoc, _cancellationToken);

            return RedirectToAction("Index");
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
    }
}