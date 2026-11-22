using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using System.Collections.Concurrent;

namespace EWeb.RPC
{
    public class RPCClient
    {
        private readonly IConnectionFactory _connectionFactory;
        private readonly ConcurrentDictionary<string, TaskCompletionSource<byte[]>> _callbackMapper = new();

        const string _queueName = "EWFileQueue";
        private string? _replyQueueName;
        private IConnection? _connection;
        private IChannel? _channel;

        public async Task StartAsync()
        {
            ConnectionFactory factory = new();
            factory.Uri = new Uri(uriString: "amqp://guest:guest@localhost:1011");// add appsettings
            factory.ClientProvidedName = "Web sender";

            _connection = await factory.CreateConnectionAsync();
            _channel = await _connection.CreateChannelAsync();

            var queueDeclareResult = await _channel.QueueDeclareAsync();
            _replyQueueName = queueDeclareResult.QueueName;
            var replyConsumer = new AsyncEventingBasicConsumer(_channel);

            replyConsumer.ReceivedAsync += (model, ea) =>
            {
                string? correlationId = ea.BasicProperties.CorrelationId;

                if (false == string.IsNullOrEmpty(correlationId))
                {
                    if (_callbackMapper.TryRemove(correlationId, out var tcs))
                    {
                        tcs.TrySetResult(ea.Body.ToArray());
                    }
                }

                return Task.CompletedTask;
            };

            var res = await _channel.BasicConsumeAsync(_replyQueueName, true, replyConsumer);
        }

        public async Task<byte[]> CallAsync(IFormFileCollection files, CancellationToken cancellationToken = default)
        {
            if (_channel is null)
            {
                throw new InvalidOperationException();
            }

            string correlationId = Guid.NewGuid().ToString();
            var props = new BasicProperties
            {
                CorrelationId = correlationId,
                ReplyTo = _replyQueueName
            };

            var tcs = new TaskCompletionSource<byte[]>(
                    TaskCreationOptions.RunContinuationsAsynchronously);
            _callbackMapper.TryAdd(correlationId, tcs);

            byte[] wordBytes = default;
            byte[] excDoc = default;

            foreach (var uploadedFile in files)
                if (uploadedFile.ContentType == "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    using (var reader = new StreamReader(uploadedFile.OpenReadStream()))
                    using (var mem = new MemoryStream())
                    {
                        reader.BaseStream.CopyTo(mem);
                        wordBytes = mem.ToArray();
                    }
                else if (uploadedFile.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    using (var reader = new StreamReader(uploadedFile.OpenReadStream()))
                    using (var mem = new MemoryStream())
                    {
                        reader.BaseStream.CopyTo(mem);
                        excDoc = mem.ToArray();
                    }

            byte[] envelope = BuildEnvelope(wordBytes, excDoc);

            await _channel.BasicPublishAsync(string.Empty, _queueName, true, props, envelope);

            using CancellationTokenRegistration ctr =
            cancellationToken.Register(() =>
            {
                _callbackMapper.TryRemove(correlationId, out _);
                tcs.SetCanceled();
            });

            return await tcs.Task;
        }

        private static byte[] BuildEnvelope(byte[] wordBytes, byte[] excelBytes)
        {
            using var payload = new MemoryStream();
            payload.Write(BitConverter.GetBytes(wordBytes.Length));
            payload.Write(wordBytes);
            payload.Write(excelBytes);
            return payload.ToArray();
        }
    }
}
