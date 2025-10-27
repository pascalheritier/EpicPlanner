namespace EpicPlanner.Core;

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

    public async Task RunAsync()
    {
        PlanningSnapshot snapshot = await m_DataProvider.LoadAsync(_bIncludePlannedHours: true);
        Simulator simulator = snapshot.CreateSimulator();
        simulator.Run();
        simulator.ExportComparisonReport(m_AppConfiguration.FileConfiguration.OutputFilePath);
    } 

    #endregion
}
