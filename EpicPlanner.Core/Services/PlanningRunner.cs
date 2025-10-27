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

    public async Task RunAsync()
    {
        PlanningSnapshot snapshot = await m_DataProvider.LoadAsync(_bIncludePlannedHours: false);
        Simulator simulator = snapshot.CreateSimulator();
        simulator.Run();
        simulator.ExportExcel(m_AppConfiguration.FileConfiguration.OutputFilePath);
        simulator.ExportGanttSprintBased(m_AppConfiguration.FileConfiguration.OutputPngFilePath);
    }

    #endregion
}
