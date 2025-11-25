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
        internal static async Task<string> ExecuteAsync(byte[] wordBytes, byte[] excBytes, MemoryStream stream)
        {
            var rpcClient = new RPCClient();
            await rpcClient.StartAsync();
            var file = await rpcClient.CallAsync(wordBytes, excBytes);

            var beginningTagLength = 1348;
            var endingTagLength = 22;
            file = file.Substring(1348, file.Length - beginningTagLength - endingTagLength);

            using (WordprocessingDocument doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
            {
                MainDocumentPart mainPart = doc.AddMainDocumentPart();

                new DocumentFormat.OpenXml.Wordprocessing.Document(new Body()).Save(mainPart);

                Body body = mainPart.Document.Body;
                body.InnerXml = file;

                mainPart.Document.Save();
            }
            stream.Seek(0, SeekOrigin.Begin);

            return file;
        }

        internal static void ReadFileBytes(ref byte[] array, string filePath)
        {
            using (StreamReader sr = new StreamReader(filePath))
            {
                using (var mem = new MemoryStream())
                {
                    sr.BaseStream.CopyTo(mem);
                    array = mem.ToArray();
                }
            }
        }
    }
}
