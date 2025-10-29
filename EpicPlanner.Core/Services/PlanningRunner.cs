namespace EpicPlanner.Core;

public class PlanningRunner
{
    #region Members

    private readonly PlanningDataProvider m_DataProvider;
    private readonly AppConfiguration m_AppConfiguration;

    #endregion

    #region Constructor

    public PlanningRunner(PlanningDataProvider _DataProvider, AppConfiguration _AppConfiguration)
    {
        m_DataProvider = _DataProvider;
        m_AppConfiguration = _AppConfiguration;
    }

    #endregion
    
    #region Run

    public async Task RunAsync(PlanningMode _Mode)
    {
        bool includePlannedHours = _Mode == PlanningMode.Standard;

        PlanningSnapshot snapshot = await m_DataProvider.LoadAsync(_bIncludePlannedHours: includePlannedHours);
        Simulator simulator = snapshot.CreateSimulator(_Mode == PlanningMode.Analysis ? IsAnalysisEpic : null);
        simulator.Run();

        if (_Mode == PlanningMode.Standard)
        {
            simulator.ExportPlanningExcel(m_AppConfiguration.FileConfiguration.OutputFilePath);
        }

        simulator.ExportGanttSprintBased(m_AppConfiguration.FileConfiguration.OutputPngFilePath);
    }

    #endregion

    #region Helpers

    private static bool IsAnalysisEpic(Epic _Epic)
    {
        if (_Epic is null)
        {
            return false;
        }

        string? state = _Epic.State;

        if (string.IsNullOrWhiteSpace(state))
        {
            return false;
        }

        return state.Contains("analysis", StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
