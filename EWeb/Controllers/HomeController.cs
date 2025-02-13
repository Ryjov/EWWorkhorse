using EWeb.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.IO;

namespace EWeb.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        IWebHostEnvironment _appEnvironment;

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
            factory.ClientProvidedName = "Rabbit sender app";

            IConnection cnn = factory.CreateConnection();
            IModel channel = cnn.CreateModel();

            string exchangeName = "DemoExchange";
            string routingKey = "demo-routing-key";
            string queueName = "DemoQueue";
            
            channel.ExchangeDeclare(exchangeName, ExchangeType.Direct);
            channel.QueueDeclare(queueName, durable: false, exclusive: false, autoDelete: false, arguments: null);
            channel.QueueBind(queueName, exchangeName, routingKey, arguments: null);

            byte[] wordBytes;
            byte[] excDoc;
            
            foreach (var uploadedFile in uploadedFiles)
            {
                if (uploadedFile.class == WordProcessingDocument)
                {
                    wordBytes = uploadedFile.ReadAllBytes(uploadedFile.File);
                }
                else if (uploadedFile.class == SpreadsheetDocument)
                {
                    Workbook wb = new Workbook();
                    excDoc = uploadedFile.File;
                }
            }
            
            channel.BasicPublish(exchangeName, routingKey, basicProperties: null, wordBytes);
            channel.BasicPublish(exchangeName, routingKey, basicProperties: null, excDoc);
            
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
