using Microsoft.AspNetCore.Http;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using System.Collections.Concurrent;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;

namespace EWPFDesktop.RPC
{
    public class RPCClient
    {
        private readonly IConnectionFactory _connectionFactory;
        private readonly ConcurrentDictionary<string, TaskCompletionSource<string>> _callbackMapper = new();

        const string _queueName = "EWFileQueue";
        private string? _replyQueueName;
        private IConnection? _connection;
        private IChannel? _channel;

        public async Task StartAsync()
        {
            ConnectionFactory factory = new();
            factory.Uri = new Uri(uriString: "amqp://guest:guest@localhost:1011");// add appsettings
            factory.ClientProvidedName = "Desktop sender";

            _connection = await factory.CreateConnectionAsync();
            _channel = await _connection.CreateChannelAsync();

            var queueDeclareResult = await _channel.QueueDeclareAsync();
            _replyQueueName = queueDeclareResult.QueueName;
            var replyConsumer = new AsyncEventingBasicConsumer(_channel);

            replyConsumer.ReceivedAsync += (model, ea) =>
            {
                var result = default(byte[]);

                string? correlationId = ea.BasicProperties.CorrelationId;

                if (false == string.IsNullOrEmpty(correlationId))
                {
                    if (_callbackMapper.TryRemove(correlationId, out var tcs))
                    {
                        result = ea.Body.ToArray();
                        var response = Encoding.UTF8.GetString(result);
                        tcs.TrySetResult(response);
                    }
                }

                return Task.CompletedTask;
            };

            var res = await _channel.BasicConsumeAsync(_replyQueueName, true, replyConsumer);
        }

        public async Task<string> CallAsync(byte[] wordBytes, byte[] excelBytes, CancellationToken cancellationToken = default)
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

            var tcs = new TaskCompletionSource<string>(
                    TaskCreationOptions.RunContinuationsAsynchronously);
            _callbackMapper.TryAdd(correlationId, tcs);

            //var fileBatch = channel.CreateBasicPublishBatch();
            await _channel.BasicPublishAsync(string.Empty, _queueName, true, props, wordBytes);// need to roll into one?
            await _channel.BasicPublishAsync(string.Empty, _queueName, true, props, excelBytes);

            using CancellationTokenRegistration ctr =
            cancellationToken.Register(() =>
            {
                _callbackMapper.TryRemove(correlationId, out _);
                tcs.SetCanceled();
            });

            return await tcs.Task;
        }
    }
}
