using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using EWeb.RPC;
using Microsoft.AspNetCore.Mvc;

namespace EWeb.Controllers
{
    public static class Executor
    {
        public static async Task ExecuteAsync(IFormFileCollection uploadedFiles, MemoryStream stream)
        {
            var file = await InvokeRPCAsync(uploadedFiles);
            var beginningTagLength = 1348;
            var endingTagLength = 22;
            file = file.Substring(1348, file.Length - beginningTagLength - endingTagLength);

            using (WordprocessingDocument doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
            {
                MainDocumentPart mainPart = doc.AddMainDocumentPart();

                new Document(new Body()).Save(mainPart);

                Body body = mainPart.Document.Body;
                body.InnerXml = file;

                mainPart.Document.Save();
            }
            stream.Seek(0, SeekOrigin.Begin);
        }

        private static async Task<string> InvokeRPCAsync(IFormFileCollection files)
        {
            var rpcClient = new RPCClient();
            await rpcClient.StartAsync();
            var resultFile = await rpcClient.CallAsync(files);

            return resultFile;
        }
    }
}
