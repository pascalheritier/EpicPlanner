using EpicPlanner.Core.Configuration;
using EpicPlanner.Core.Planner;
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

            PlanningMode mode = ResolvePlanningMode(args);

            IServiceCollection services = new ServiceCollection();
            ConfigureServices(services);
            IServiceProvider serviceProvider = services.BuildServiceProvider();

            var runner = serviceProvider.GetRequiredService<PlanningRunner>();
            await runner.RunAsync(mode);

            AppConfiguration config = serviceProvider.GetRequiredService<AppConfiguration>();
            LogManager.GetCurrentClassLogger().Log(NLog.LogLevel.Info, "✅ Planning completed");
            if (mode == PlanningMode.Standard)
            {
                LogManager.GetCurrentClassLogger().Log(NLog.LogLevel.Info, $"Excel: {Path.GetFullPath(config.FileConfiguration.OutputFilePath)}");
            }
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
        _Services.AddPlannerCore();

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

    private static PlanningMode ResolvePlanningMode(string[] _Args)
    {
        if (_Args.Length > 1 && _Args[0].Equals("--mode", StringComparison.OrdinalIgnoreCase))
        {
            if (TryParsePlanningMode(_Args[1], out PlanningMode fromFlag))
            {
                return fromFlag;
            }
        }

        foreach (string arg in _Args)
        {
            if (arg.StartsWith("--mode=", StringComparison.OrdinalIgnoreCase))
            {
                string candidate = arg[("--mode=").Length..];
                if (TryParsePlanningMode(candidate, out PlanningMode modeFromOption))
                {
                    return modeFromOption;
                }
            }
            else if (TryParsePlanningMode(arg, out PlanningMode directMode))
            {
                return directMode;
            }
        }

        Console.WriteLine("Select the planning mode:");
        Console.WriteLine("  1 - Standard planning (sprint planning)");
        Console.WriteLine("  2 - Analysis planning (strategic planning)");

        while (true)
        {
            Console.Write("Enter your choice: ");
            string? input = Console.ReadLine();
            if (input is null)
            {
                Console.WriteLine("No input detected. Defaulting to standard planning.");
                return PlanningMode.Standard;
            }

            if (TryParsePlanningMode(input, out PlanningMode parsed))
            {
                return parsed;
            }

            Console.WriteLine("Invalid selection. Please type '1' for Standard planning or '2' for Analysis planning.");
        }
    }

    private static bool TryParsePlanningMode(string? _strInput, out PlanningMode _enumMode)
    {
        _enumMode = PlanningMode.Standard;
        if (string.IsNullOrWhiteSpace(_strInput))
        {
            return false;
        }

        string value = _strInput.Trim();

        if (value.Equals("1", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("standard", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("plan", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("planning", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("sprint", StringComparison.OrdinalIgnoreCase))
        {
            _enumMode = PlanningMode.Standard;
            return true;
        }

        if (value.Equals("2", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("analysis", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("strategic", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("strategy", StringComparison.OrdinalIgnoreCase))
        {
            _enumMode = PlanningMode.Analysis;
            return true;
        }

        return false;
    }
}
