using Microsoft.AspNetCore.Http;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using System.Collections.Concurrent;
using System.IO;
using System.Text;

namespace EWPFDesktop.RPC
{
    public class RPCClient
    {
        private readonly IConnectionFactory _connectionFactory;
        private readonly ConcurrentDictionary<string, TaskCompletionSource<byte[]>> _callbackMapper = new();
        private static readonly TimeSpan ReplyTimeout = TimeSpan.FromSeconds(30);

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
                string? correlationId = ea.BasicProperties.CorrelationId;

                if (false == string.IsNullOrEmpty(correlationId) && _callbackMapper.TryRemove(correlationId, out var tcs))
                    if (ea.BasicProperties.Type == "error")
                        tcs.TrySetException(new InvalidOperationException("Workhorse failed: " + Encoding.UTF8.GetString(ea.Body.ToArray())));
                    else
                        tcs.TrySetResult(ea.Body.ToArray());

                return Task.CompletedTask;
            };

            var res = await _channel.BasicConsumeAsync(_replyQueueName, true, replyConsumer);
        }

        public ValueTask DisposeAsync() => new(CloseAsync());

        private async Task CloseAsync()
        {
            if (_channel is null && _connection is null) return;
            if (_channel is not null) { await _channel.CloseAsync(); _channel = null; }
            if (_connection is not null) { await _connection.CloseAsync(); _connection = null; }
        }

        public async Task<byte[]> CallAsync(byte[] wordBytes, byte[] excelBytes, CancellationToken cancellationToken = default)
        {
            if (_channel is null)
            {
                throw new InvalidOperationException();
            }

            string correlationId = Guid.NewGuid().ToString();
            var props = new BasicProperties
            {
                CorrelationId = correlationId,
                ReplyTo = _replyQueueName,
                DeliveryMode = DeliveryModes.Persistent
            };

            var tcs = new TaskCompletionSource<byte[]>(
                    TaskCreationOptions.RunContinuationsAsynchronously);
            _callbackMapper.TryAdd(correlationId, tcs);

            var envelope = BuildEnvelope(wordBytes, excelBytes);
            await _channel.BasicPublishAsync(string.Empty, _queueName, true, props, envelope);

            using CancellationTokenRegistration ctr =
            cancellationToken.Register(() =>
            {
                _callbackMapper.TryRemove(correlationId, out _);
                tcs.SetCanceled();
            });

            try
            {
                return await tcs.Task.WaitAsync(ReplyTimeout, cancellationToken);
            }
            catch
            {
                _callbackMapper.TryRemove(correlationId, out _);
                throw;
            }
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
