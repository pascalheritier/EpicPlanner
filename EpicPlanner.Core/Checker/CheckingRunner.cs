using EpicPlanner.Core.Checker.Reports;
using EpicPlanner.Core.Checker.Services;
using EpicPlanner.Core.Checker.Simulation;
using EpicPlanner.Core.Configuration;
using EpicPlanner.Core.Shared.Services;
using System.Threading.Tasks;

namespace EpicPlanner.Core.Checker;

public class CheckingRunner
{
    #region Members

    private readonly PlanningDataProvider m_DataProvider;
    private readonly AppConfiguration m_AppConfiguration;
    private readonly EpicAnalysisDataLoader m_AnalysisDataLoader;
    private readonly EpicAnalysisHtmlGenerator m_AnalysisHtmlGenerator;

    #endregion

    #region Constructor

    public CheckingRunner(
        PlanningDataProvider _DataProvider,
        AppConfiguration _AppConfiguration,
        EpicAnalysisDataLoader _AnalysisDataLoader,
        EpicAnalysisHtmlGenerator _AnalysisHtmlGenerator)
    {
        m_DataProvider = _DataProvider;
        m_AppConfiguration = _AppConfiguration;
        m_AnalysisDataLoader = _AnalysisDataLoader;
        m_AnalysisHtmlGenerator = _AnalysisHtmlGenerator;
    }

    #endregion

    #region Run

    public async Task<string> RunAsync(EnumCheckerMode _enumMode)
    {
        CheckerPlanningSnapshot snapshot = await m_DataProvider.LoadCheckerSnapshotAsync();
        CheckerSimulator simulator = snapshot.CreateCheckerSimulator();
        simulator.Run();
        string outputPath = ResolveOutputPath(_enumMode);
        simulator.ExportCheckerReport(outputPath, _enumMode);

        if (_enumMode == EnumCheckerMode.EpicStates)
        {
            await GenerateEpicAnalysisHtmlAsync();
        }

        return outputPath;
    }

    private async Task GenerateEpicAnalysisHtmlAsync()
    {
        string htmlPath = m_AppConfiguration.EpicAnalysisReportConfiguration.OutputHtmlPath;
        if (string.IsNullOrWhiteSpace(htmlPath)) return;

        EpicAnalysisReportModel model = await m_AnalysisDataLoader.LoadAsync();
        m_AnalysisHtmlGenerator.Generate(model, htmlPath);
    }

    #endregion

    #region Helpers

    private string ResolveOutputPath(EnumCheckerMode _enumMode)
    {
        string basePath = m_AppConfiguration.FileConfiguration.OutputFilePath;
        if (string.IsNullOrWhiteSpace(basePath))
        {
            basePath = Path.Combine(Directory.GetCurrentDirectory(), "CheckerReport.xlsx");
        }

        string directory = Path.GetDirectoryName(basePath) ?? string.Empty;
        if (string.IsNullOrWhiteSpace(directory))
        {
            directory = Directory.GetCurrentDirectory();
        }

        string fileName = Path.GetFileNameWithoutExtension(basePath);
        if (string.IsNullOrWhiteSpace(fileName))
        {
            fileName = "CheckerReport";
        }

        string extension = Path.GetExtension(basePath);
        if (string.IsNullOrWhiteSpace(extension))
        {
            extension = ".xlsx";
        }

        Directory.CreateDirectory(directory);

        string suffix = _enumMode == EnumCheckerMode.Comparison ? "_Comparison" : "_EpicStates";
        return Path.Combine(directory, fileName + suffix + extension);
    }

    #endregion
}
