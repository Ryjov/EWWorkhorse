using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using EWPFDesktop.RPC;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;

namespace EWPFDesktop
{
    internal static class Executor
    {
        internal static async Task<byte[]> ExecuteAsync(byte[] wordBytes, byte[] excBytes)
        {
            var rpcClient = new RPCClient();
            await rpcClient.StartAsync();
            return await rpcClient.CallAsync(wordBytes, excBytes);
        }
    }
}
