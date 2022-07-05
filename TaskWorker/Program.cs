using TaskWorker.Interfaces;
using TaskWorker.Services;
using EmailSender.Interfaces;
using EmailSender.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System;
using System.IO;
using Serilog;
using Serilog.Events;

namespace TaskWorker
{
    class Program
    {
        //TODO refactor to a worker instead of depending on OS Scheduled Tasks
        static void Main(string[] args)
        {
            Log.Logger = new LoggerConfiguration()
              .MinimumLevel.Debug()
              .WriteTo.File("log.txt")  // log file.
              .WriteTo.Console(restrictedToMinimumLevel: LogEventLevel.Information)
              .CreateLogger();

            Console.WriteLine("Task Started");
            var host = CreateHostBuilder(args).Build();
            host.Services.GetService<IBusiness>().RunBusinessTask();
            Console.WriteLine("Task Ended");
        }

        private static IHostBuilder CreateHostBuilder(string[] args)
        {
            var hostBuilder = Host.CreateDefaultBuilder(args)
                .ConfigureAppConfiguration((context, builder) =>
                {
                    builder.SetBasePath(Directory.GetCurrentDirectory());
                })
                .ConfigureServices((context, services) =>
                {
                    services.AddScoped<IBusiness, BusinessService>();
                    services.AddScoped<IEmailSender, EmailSenderService>();
                    services.AddSingleton(Log.Logger);
                }).UseSerilog()
                ;
            return hostBuilder;
        }

    }
}
