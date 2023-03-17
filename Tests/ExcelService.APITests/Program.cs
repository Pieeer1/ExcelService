using ExcelService.Interfaces;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
namespace ExcelService.APITests;

public class Program
{
    public static void Main(string[] args)
    {
        var host = new HostBuilder()
            .ConfigureFunctionsWorkerDefaults()
            .ConfigureServices(services =>
            {
                services.AddScoped<IExcel, Excel>();
                services.AddLogging();
            })
            .Build();

        host.Run();
    }
}


