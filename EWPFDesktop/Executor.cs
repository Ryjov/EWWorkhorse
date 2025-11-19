using EWPFDesktop.RPC;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWPFDesktop
{
    internal static class Executor
    {
        internal static async Task ExecuteAsync(byte[] wordBytes, byte[] excBytes, MemoryStream stream)
        {
            var rpcClient = new RPCClient();
            await rpcClient.StartAsync();
            var resultFile = await rpcClient.CallAsync(wordBytes, excBytes);

            return resultFile;
        }
    }
}
