using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using NLog;
using NLog.Config;
using NLog.Extensions.Logging;
using OfficeOpenXml;

namespace EpicPlanner
{
    internal class Program
    {
        #region App Config

        private static string AppSettingsFileName = "appsettings.json";
        private static string LogConfigFileName = "NLog.config";

        #endregion

        #region Main

        static async Task Main(string[] args)
        {
            // EPPlus license (non-commercial)
            ExcelPackage.License.SetNonCommercialPersonal("My Name"); //This will also set the Author property to the name provided in the argument.

            // Input/outputs (adjust paths as needed)
            string inputPath = "[Athena] Planification_des_Epics.xlsx";
            string outputXlsx = @"D:\Users\pascal.heritier\OneDrive - Watchout\Bureau\epics_strict_with_spillover_v5_approved.xlsx";
            string outputPng = @"D:\Users\pascal.heritier\OneDrive - Watchout\Bureau\epics_strict_with_spillover_v5_gantt_sprints.png";

            try
            {
                IServiceCollection services = new ServiceCollection();
                ConfigureServices(services);
                IServiceProvider serviceProvider = services.BuildServiceProvider();

                var planner = serviceProvider.GetRequiredService<Planner>();
                await planner.RunAsync(inputPath, outputXlsx, outputPng);
                Console.WriteLine("✅ Done");
                Console.WriteLine($"Excel: {Path.GetFullPath(outputXlsx)}");
                Console.WriteLine($"Gantt: {Path.GetFullPath(outputPng)}");
            }
            catch (Exception ex)
            {
                LogManager.GetCurrentClassLogger().Log(NLog.LogLevel.Fatal, $"Critical app failure: {ex.Message}{Environment.NewLine}{ex.StackTrace}");
                Console.WriteLine("❌ Error: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        #endregion

        #region Services

        private static void ConfigureServices(IServiceCollection services)
        {
            IConfiguration config = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory()) //From NuGet Package Microsoft.Extensions.Configuration.Json
            .AddJsonFile(AppSettingsFileName, optional: true, reloadOnChange: true)
            .Build();

            services.AddSingleton<AppConfiguration>(_X => GetAppConfiguration(config));
            services.AddLogging(loggingBuilder =>
            {
                // configure Logging with NLog
                loggingBuilder.ClearProviders();
                loggingBuilder.SetMinimumLevel(Microsoft.Extensions.Logging.LogLevel.Trace);
                loggingBuilder.AddNLog(GetLogConfiguration());
            });
            services.AddTransient<Planner>();
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

        private static AppConfiguration GetAppConfiguration(IConfiguration configuration)
        {
            AppConfiguration appConfiguration = new();
            configuration.Bind(appConfiguration);
            return appConfiguration;
        }

        #endregion
    }
}
