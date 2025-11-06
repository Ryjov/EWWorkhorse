using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using System.Collections.Concurrent;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;

namespace EWeb.RPC
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

            //string exchangeName = "EWExchange";
            //string routingKey = "ew-routing-key";

            //callback queue
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
                        var body = ea.Body.ToArray();
                        var response = Encoding.UTF8.GetString(body);
                        tcs.TrySetResult(response);
                    }
                }

                return Task.CompletedTask;
            };

            var res = await _channel.BasicConsumeAsync(_replyQueueName, true, replyConsumer);

            //_channel.ExchangeDeclareAsync(exchangeName, ExchangeType.Direct);
            //_channel.QueueDeclareAsync(_queueName, durable: false, exclusive: false, autoDelete: false, arguments: null);
            //_channel.QueueBindAsync(_queueName, exchangeName, routingKey, arguments: null);
        }

        public async Task<string> CallAsync(byte[] wordBytes, byte[] excBytes, CancellationToken cancellationToken = default)
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

            ConnectionFactory factory = new();
            factory.Uri = new Uri(uriString: "amqp://guest:guest@localhost:1011");// add appsettings
            factory.ClientProvidedName = "EW filebytes sender app";

            var cnn = factory.CreateConnectionAsync().Result;
            var channel = cnn.CreateChannelAsync().Result;

            string exchangeName = "EWPFExchange";
            string routingKey = "ew-routing-key";
            string queueName = "EWFileQueue";

            channel.ExchangeDeclareAsync(exchangeName, ExchangeType.Direct);
            channel.QueueDeclareAsync(queueName, durable: false, exclusive: false, autoDelete: false, arguments: null);
            channel.QueueBindAsync(queueName, exchangeName, routingKey, arguments: null);

            

            //var fileBatch = channel.CreateBasicPublishBatch();
            channel.BasicPublishAsync(exchangeName, routingKey, true, wordBytes, cancellationToken);// need to roll into one?
            channel.BasicPublishAsync(exchangeName, routingKey, true, excBytes, cancellationToken);

            using (WebClient wc = new WebClient())
            {
                //wc.DownloadProgressChanged += wc_DownloadProgressChanged;
                wc.DownloadFileAsync(new System.Uri("http://url"),
                 "Result location");
            }

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
