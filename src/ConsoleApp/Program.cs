using ConsoleApp.Services;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

namespace ConsoleApp
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            var builder = Host.CreateDefaultBuilder(args)
              .ConfigureServices((hostContext, services) =>
               {
                   services.AddScoped<FileService>();
                   services.AddScoped<ExcelService>();
               })
              .UseConsoleLifetime();

            var host = builder.Build();

            using (var scope = host.Services.CreateScope())
            {
                var serviceProvider = scope.ServiceProvider;

                var fileService = serviceProvider.GetRequiredService<FileService>();
                var excelAnalysisedFileFactory = serviceProvider.GetRequiredService<ExcelService>();

                var sources = await fileService.ReadSourceAsync();
                var bytes = await excelAnalysisedFileFactory.CreateAsync(sources);
                await fileService.WriteExcelFileAsync(bytes);
            }
        }
    }
}