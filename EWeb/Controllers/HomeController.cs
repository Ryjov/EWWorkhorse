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
        public async Task<IActionResult> ExecuteMerge(IFormFileCollection uploadedFiles)
        {
            var stream = new MemoryStream();
            await Executor.ExecuteAsync(uploadedFiles, stream);

            return File(stream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "result.docx");
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}