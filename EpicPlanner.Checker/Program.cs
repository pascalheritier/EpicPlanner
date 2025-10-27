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

            CheckerMode mode = ResolveMode(args);

            IServiceCollection services = new ServiceCollection();
            ConfigureServices(services);
            IServiceProvider serviceProvider = services.BuildServiceProvider();

            var runner = serviceProvider.GetRequiredService<CheckingRunner>();
            string reportPath = await runner.RunAsync(mode);

            string modeLabel = mode == CheckerMode.Comparison ? "Comparison" : "Epic states";
            LogManager.GetCurrentClassLogger().Log(NLog.LogLevel.Info, "✅ Check completed");
            LogManager.GetCurrentClassLogger().Log(NLog.LogLevel.Info, $"{modeLabel} report: {Path.GetFullPath(reportPath)}");
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
        _Services.AddTransient<CheckingRunner>();

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

    private static CheckerMode ResolveMode(string[] _Args)
    {
        if (_Args.Length > 1 && _Args[0].Equals("--mode", StringComparison.OrdinalIgnoreCase))
        {
            if (TryParseMode(_Args[1], out var modeFromFlag))
            {
                return modeFromFlag;
            }
        }

        foreach (string arg in _Args)
        {
            if (arg.StartsWith("--mode=", StringComparison.OrdinalIgnoreCase))
            {
                string candidate = arg[("--mode=").Length..];
                if (TryParseMode(candidate, out var mode))
                {
                    return mode;
                }
            }
            else if (TryParseMode(arg, out var directMode))
            {
                return directMode;
            }
        }

        Console.WriteLine("Select the checker mode:");
        Console.WriteLine("  1 - Comparison (sprint planning)");
        Console.WriteLine("  2 - Epic states (sprint review)");

        while (true)
        {
            Console.Write("Enter your choice: ");
            string? input = Console.ReadLine();
            if (input is null)
            {
                Console.WriteLine("No input detected. Defaulting to Comparison mode.");
                return CheckerMode.Comparison;
            }

            if (TryParseMode(input, out var parsedMode))
            {
                return parsedMode;
            }

            Console.WriteLine("Invalid selection. Please type '1' for Comparison or '2' for Epic states.");
        }
    }

    private static bool TryParseMode(string? _strInput, out CheckerMode _enumMode)
    {
        _enumMode = CheckerMode.Comparison;
        if (string.IsNullOrWhiteSpace(_strInput))
        {
            return false;
        }

        string value = _strInput.Trim();

        if (value.Equals("1", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("comparison", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("planning", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("plan", StringComparison.OrdinalIgnoreCase))
        {
            _enumMode = CheckerMode.Comparison;
            return true;
        }

        if (value.Equals("2", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("epic", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("epicstates", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("epic-state", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("review", StringComparison.OrdinalIgnoreCase))
        {
            _enumMode = CheckerMode.EpicStates;
            return true;
        }

        return false;
    }
}
