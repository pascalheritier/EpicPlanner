using EpicPlanner.Core;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using NLog;
using NLog.Config;
using NLog.Extensions.Logging;
using OfficeOpenXml;

namespace EpicPlanner.Checker;

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

            var runner = serviceProvider.GetRequiredService<CheckingRunner>();
            await runner.RunAsync();

            AppConfiguration config = serviceProvider.GetRequiredService<AppConfiguration>();
            string basePath = config.FileConfiguration.OutputFilePath;
            string directory = Path.GetDirectoryName(basePath) ?? string.Empty;
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(basePath);
            string extension = Path.GetExtension(basePath);
            if (string.IsNullOrWhiteSpace(extension))
            {
                extension = ".xlsx";
            }
            string comparisonPath = Path.Combine(directory, $"{fileNameWithoutExtension}_Comparison{extension}");

            LogManager.GetCurrentClassLogger().Log(NLog.LogLevel.Info, "✅ Check completed");
            LogManager.GetCurrentClassLogger().Log(NLog.LogLevel.Info, $"Comparison report: {Path.GetFullPath(comparisonPath)}");
        }
        catch (Exception ex)
        {
            LogManager.GetCurrentClassLogger().Log(NLog.LogLevel.Fatal, $"Critical app failure: {ex.Message}{Environment.NewLine}{ex.StackTrace}");
            LogManager.GetCurrentClassLogger().Log(NLog.LogLevel.Fatal, "❌ Error: " + ex.Message);
            LogManager.GetCurrentClassLogger().Log(NLog.LogLevel.Fatal, ex.StackTrace);
        }
    }

    private static void ConfigureServices(IServiceCollection services)
    {
        IConfiguration config = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile(AppSettingsFileName, optional: true, reloadOnChange: true)
            .Build();

        services.AddSingleton(_ => GetAppConfiguration(config));
        services.AddSingleton<PlanningDataProvider>();
        services.AddTransient<CheckingRunner>();

        services.AddLogging(loggingBuilder =>
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

    private static AppConfiguration GetAppConfiguration(IConfiguration configuration)
    {
        AppConfiguration appConfiguration = new();
        configuration.Bind(appConfiguration);
        return appConfiguration;
    }
}
