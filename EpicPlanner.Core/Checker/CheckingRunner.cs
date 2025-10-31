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

    #endregion

    #region Constructor

    public CheckingRunner(PlanningDataProvider _DataProvider, AppConfiguration _AppConfiguration)
    {
        m_DataProvider = _DataProvider;
        m_AppConfiguration = _AppConfiguration;
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
        return outputPath;
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
