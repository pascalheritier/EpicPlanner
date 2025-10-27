using EpicPlanner.Core;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using NLog;
using NLog.Config;
using NLog.Extensions.Logging;
using OfficeOpenXml;

namespace EpicPlanner.Planner;

internal class Program
{
    private const string AppSettingsFileName = "appsettings.json";
    private const string LogConfigFileName = "NLog.config";

    static async Task Main(string[] args)
    {
        try
        {
            ExcelPackage.License.SetNonCommercialPersonal("Adonite");

            IServiceCollection services = new ServiceCollection();
            ConfigureServices(services);
            IServiceProvider serviceProvider = services.BuildServiceProvider();

            var runner = serviceProvider.GetRequiredService<PlanningRunner>();
            await runner.RunAsync();

            AppConfiguration config = serviceProvider.GetRequiredService<AppConfiguration>();
            LogManager.GetCurrentClassLogger().Log(NLog.LogLevel.Info, "✅ Planning completed");
            LogManager.GetCurrentClassLogger().Log(NLog.LogLevel.Info, $"Excel: {Path.GetFullPath(config.FileConfiguration.OutputFilePath)}");
            LogManager.GetCurrentClassLogger().Log(NLog.LogLevel.Info, $"Gantt: {Path.GetFullPath(config.FileConfiguration.OutputPngFilePath)}");
        }
        catch (Exception ex)
        {
            LogManager.GetCurrentClassLogger().Log(NLog.LogLevel.Fatal, $"Critical app failure: {ex.Message}{Environment.NewLine}{ex.StackTrace}");
            LogManager.GetCurrentClassLogger().Log(NLog.LogLevel.Fatal, "❌ Error: " + ex.Message);
            LogManager.GetCurrentClassLogger().Log(NLog.LogLevel.Fatal, ex.StackTrace);
        }
    }

    private static void ConfigureServices(IServiceCollection _Services)
    {
        IConfiguration config = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile(AppSettingsFileName, optional: true, reloadOnChange: true)
            .Build();

        _Services.AddSingleton(_ => GetAppConfiguration(config));
        _Services.AddSingleton<PlanningDataProvider>();
        _Services.AddTransient<PlanningRunner>();

        _Services.AddLogging(loggingBuilder =>
        {
            loggingBuilder.ClearProviders();
            loggingBuilder.SetMinimumLevel(Microsoft.Extensions.Logging.LogLevel.Trace);
            loggingBuilder.AddNLog(GetLogConfiguration());
        });
    }

    private static LoggingConfiguration GetLogConfiguration()
    {
        var stream = typeof(Program).Assembly.GetManifestResourceStream($"{typeof(Program).Namespace}.{LogConfigFileName}");
        if (stream is null)
        {
            throw new InvalidOperationException($"Embedded log configuration '{LogConfigFileName}' not found.");
        }

        using var reader = new StreamReader(stream);
        string xml = reader.ReadToEnd();
        return XmlLoggingConfiguration.CreateFromXmlString(xml);
    }

    private static AppConfiguration GetAppConfiguration(IConfiguration _Configuration)
    {
        AppConfiguration appConfiguration = new();
        _Configuration.Bind(appConfiguration);
        return appConfiguration;
    }
}
