using EWWorkhorse;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using System.Text;

var builder = Host.CreateApplicationBuilder(args);
builder.Services.AddHostedService<Worker>();

var host = builder.Build();
host.Run();