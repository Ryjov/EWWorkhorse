using EWeb.Models;
using EWeb.RPC;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace EWeb.Controllers
{
    public class HomeController : Controller
    {
        private readonly RPCClient _rpcClient;
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger, RPCClient rpcClient)
        {
            _logger = logger;
            _rpcClient = rpcClient;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> ExecuteMerge(IFormFileCollection uploadedFiles)
        {
            var resultBytes = await _rpcClient.CallAsync(uploadedFiles, HttpContext.RequestAborted);

            return File(resultBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "result.docx");
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}