using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using NLog;
using NLog.Config;
using NLog.Extensions.Logging;
using OfficeOpenXml;

namespace EpicPlanner;

internal class Program
{
    #region App Config

    private static string AppSettingsFileName = "appsettings.json";
    private static string LogConfigFileName = "NLog.config";

    #endregion

    #region Main

    static async Task Main(string[] _Args)
    {
        try
        {
            ExcelPackage.License.SetNonCommercialPersonal("Adonite"); //This will also set the Author property to the name provided in the argument.
            IServiceCollection services = new ServiceCollection();
            ConfigureServices(services);
            IServiceProvider serviceProvider = services.BuildServiceProvider();

            var planner = serviceProvider.GetRequiredService<Planner>();
            await planner.RunAsync();

            AppConfiguration config = serviceProvider.GetRequiredService<AppConfiguration>();
            LogManager.GetCurrentClassLogger().Log(NLog.LogLevel.Info, "✅ Done");
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

    #endregion

    #region Services

    private static void ConfigureServices(IServiceCollection _Services)
    {
        IConfiguration config = new ConfigurationBuilder()
        .SetBasePath(Directory.GetCurrentDirectory()) //From NuGet Package Microsoft.Extensions.Configuration.Json
        .AddJsonFile(AppSettingsFileName, optional: true, reloadOnChange: true)
        .Build();

        _Services.AddSingleton<AppConfiguration>(_X => GetAppConfiguration(config));
        _Services.AddLogging(loggingBuilder =>
        {
            // configure Logging with NLog
            loggingBuilder.ClearProviders();
            loggingBuilder.SetMinimumLevel(Microsoft.Extensions.Logging.LogLevel.Trace);
            loggingBuilder.AddNLog(GetLogConfiguration());
        });
        _Services.AddTransient<Planner>();
    }

    #endregion

    #region Application configuration

    private static LoggingConfiguration GetLogConfiguration()
    {
        var stream = typeof(Program).Assembly.GetManifestResourceStream("EpicPlanner." + LogConfigFileName);
        string xml;
        using (var reader = new StreamReader(stream))
        {
            xml = reader.ReadToEnd();
        }
        return XmlLoggingConfiguration.CreateFromXmlString(xml);
    }

    private static AppConfiguration GetAppConfiguration(IConfiguration _Configuration)
    {
        AppConfiguration appConfiguration = new();
        _Configuration.Bind(appConfiguration);
        return appConfiguration;
    }

    #endregion
}