using EWWorkhorse;
using EWWorkhorse.RPC.Services;
using Microsoft.AspNetCore.Builder;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using System.Text;

var builder = Host.CreateApplicationBuilder(args);
              //WebApplication.CreateBuilder(args);

builder.Services.AddHostedService<Worker>();
//builder.Services.AddGrpc();

var host = builder.Build();

//host.MapGrpcService<FileExchangeService>();

host.Run();