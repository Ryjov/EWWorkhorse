using EWWorkhorse;
using Microsoft.AspNetCore.Builder;

var builder = Host.CreateApplicationBuilder(args);

builder.Services.AddHostedService<Worker>();

var host = builder.Build();

host.Run();