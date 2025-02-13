using DocumentFormat.OpenXml.Packaging;
using EWeb.Models;
using Microsoft.AspNetCore.Connections;
using Microsoft.AspNetCore.Mvc;
using RabbitMQ.Client;
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
            factory.Uri = new Uri(uriString: "amqp://guest:guest@localhost:5672");// add appsettings
            factory.ClientProvidedName = "EW filebytes sender app";

            var cnn = await factory.CreateConnectionAsync();
            var channel = await cnn.CreateChannelAsync();

            string exchangeName = "EWExchange";
            string routingKey = "ew-routing-key";
            string queueName = "EWFileQueue";

            channel.ExchangeDeclareAsync(exchangeName, ExchangeType.Direct);
            channel.QueueDeclareAsync(queueName, durable: false, exclusive: false, autoDelete: false, arguments: null);
            channel.QueueBindAsync(queueName, exchangeName, routingKey, arguments: null);

            byte[] wordBytes;
            byte[] excDoc;

            foreach (var uploadedFile in uploadedFiles)
            {
                if (uploadedFile.ContentType == "")
                {
                    using (var reader = new StreamReader(uploadedFile.OpenReadStream()))
                    {
                        var s = await reader.ReadToEndAsync();
                    }
                }
                else if (uploadedFile.ContentType == "")
                {
                    using (var reader = new StreamReader(uploadedFile.OpenReadStream()))
                    {
                        excDoc = await reader.ReadToEndAsync();
                    }
                }
            }

            await channel.BasicPublishAsync(exchangeName, routingKey, true, null, wordBytes, _cancellationToken);
            await channel.BasicPublishAsync(exchangeName, routingKey, true, null, excDoc, _cancellationToken);

            return RedirectToAction("Index");
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