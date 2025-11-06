using Google.Protobuf;
using Grpc.Core;

namespace EWWorkhorse.RPC.Services
{
    public class FileExchangeService : FileExchanger.FileExchangerBase
    {
        public FileExchangeService() 
        {
        }

        public override Task<ProcessingResult> ProcessFiles(ProcessingRequest request, ServerCallContext context)
        {
            return Task.FromResult(new ProcessingResult
            {
                ResultWordBytes = ByteString.CopyFrom(Replacer.ReplaceBytes(request.WordBytes.ToByteArray(), request.ExcelBytes.ToByteArray()))
            });
        }
    }
}
