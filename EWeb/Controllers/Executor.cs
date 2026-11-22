using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using EWeb.RPC;
using Microsoft.AspNetCore.Mvc;

namespace EWeb.Controllers
{
    internal static class Executor
    {
        internal static async Task<byte[]> ExecuteAsync(IFormFileCollection uploadedFiles)
        {
            var rpcClient = new RPCClient();
            await rpcClient.StartAsync();
            return await rpcClient.CallAsync(uploadedFiles);
        }
    }
}
