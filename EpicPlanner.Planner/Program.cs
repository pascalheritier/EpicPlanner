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

    static async Task Main(string[] _Args)
    {
        try
        {
            ExcelPackage.License.SetNonCommercialPersonal("Adonite");

            EnumPlanningMode mode = ResolvePlanningMode(_Args);

            IServiceCollection services = new ServiceCollection();
            ConfigureServices(services);
            IServiceProvider serviceProvider = services.BuildServiceProvider();

            var runner = serviceProvider.GetRequiredService<PlanningRunner>();
            await runner.RunAsync(mode);

            AppConfiguration config = serviceProvider.GetRequiredService<AppConfiguration>();
            LogManager.GetCurrentClassLogger().Log(NLog.LogLevel.Info, "✅ Planning completed");
            if (mode == EnumPlanningMode.Standard || mode == EnumPlanningMode.StrategicEpic)
            {
                LogManager.GetCurrentClassLogger().Log(
                    NLog.LogLevel.Info,
                    $"Excel: {Path.GetFullPath(ResolveExcelOutputPath(config, mode))}");
            }
            LogManager.GetCurrentClassLogger().Log(
                NLog.LogLevel.Info,
                $"Gantt: {Path.GetFullPath(ResolvePngOutputPath(config, mode))}");
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

    private static EnumPlanningMode ResolvePlanningMode(string[] _Args)
    {
        if (_Args.Length > 1 && _Args[0].Equals("--mode", StringComparison.OrdinalIgnoreCase))
        {
            if (TryParsePlanningMode(_Args[1], out EnumPlanningMode fromFlag))
            {
                return fromFlag;
            }
        }

        foreach (string arg in _Args)
        {
            if (arg.StartsWith("--mode=", StringComparison.OrdinalIgnoreCase))
            {
                string candidate = arg[("--mode=").Length..];
                if (TryParsePlanningMode(candidate, out EnumPlanningMode modeFromOption))
                {
                    return modeFromOption;
                }
            }
            else if (TryParsePlanningMode(arg, out EnumPlanningMode directMode))
            {
                return directMode;
            }
        }

        Console.WriteLine("Select the planning mode:");
        Console.WriteLine("  1 - Standard planning (sprint planning)");
        Console.WriteLine("  2 - Analysis planning (strategic planning)");
        Console.WriteLine("  3 - Strategic epic planning (version-based epic scheduling)");

        while (true)
        {
            Console.Write("Enter your choice: ");
            string? input = Console.ReadLine();
            if (input is null)
            {
                Console.WriteLine("No input detected. Defaulting to standard planning.");
                return EnumPlanningMode.Standard;
            }

            if (TryParsePlanningMode(input, out EnumPlanningMode parsed))
            {
                return parsed;
            }

            Console.WriteLine("Invalid selection. Please type '1' for Standard planning, '2' for Analysis planning, or '3' for Strategic epic planning.");
        }
    }

    private static bool TryParsePlanningMode(string? _strInput, out EnumPlanningMode _enumMode)
    {
        _enumMode = EnumPlanningMode.Standard;
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
            _enumMode = EnumPlanningMode.Standard;
            return true;
        }

        if (value.Equals("2", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("analysis", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("strategic", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("strategy", StringComparison.OrdinalIgnoreCase))
        {
            _enumMode = EnumPlanningMode.Analysis;
            return true;
        }

        if (value.Equals("3", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("strategic-epic", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("epic", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("themes", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("strategic-epic-planning", StringComparison.OrdinalIgnoreCase))
        {
            _enumMode = EnumPlanningMode.StrategicEpic;
            return true;
        }

        return false;
    }

    private static string ResolveExcelOutputPath(AppConfiguration _Config, EnumPlanningMode _Mode)
    {
        return _Mode == EnumPlanningMode.StrategicEpic &&
            !string.IsNullOrWhiteSpace(_Config.FileConfiguration.StrategicOutputFilePath)
            ? _Config.FileConfiguration.StrategicOutputFilePath!
            : _Config.FileConfiguration.OutputFilePath;
    }

    private static string ResolvePngOutputPath(AppConfiguration _Config, EnumPlanningMode _Mode)
    {
        return _Mode == EnumPlanningMode.StrategicEpic &&
            !string.IsNullOrWhiteSpace(_Config.FileConfiguration.StrategicOutputPngFilePath)
            ? _Config.FileConfiguration.StrategicOutputPngFilePath!
            : _Config.FileConfiguration.OutputPngFilePath;
    }
}
