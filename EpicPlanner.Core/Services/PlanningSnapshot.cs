namespace EpicPlanner.Core;

public class PlanningSnapshot
{
    #region Members

    private readonly List<Epic> m_Epics;
    private readonly Dictionary<int, Dictionary<string, ResourceCapacity>> m_SprintCapacities;
    private readonly DateTime m_InitialSprintStart;
    private readonly int m_SprintDays;
    private readonly int m_MaxSprintCount;
    private readonly int m_SprintOffset;
    private readonly Dictionary<string, double> m_PlannedHours;

    #endregion

    #region Constructor

    public PlanningSnapshot(
        List<Epic> _Epics,
        Dictionary<int, Dictionary<string, ResourceCapacity>> _SprintCapacities,
        DateTime _InitialSprintStart,
        int _iSprintDays,
        int _iMaxSprintCount,
        int _iSprintOffset,
        Dictionary<string, double> _PlannedHours)
    {
        m_Epics = _Epics;
        m_SprintCapacities = _SprintCapacities;
        m_InitialSprintStart = _InitialSprintStart;
        m_SprintDays = _iSprintDays;
        m_MaxSprintCount = _iMaxSprintCount;
        m_SprintOffset = _iSprintOffset;
        m_PlannedHours = _PlannedHours;
    }

    #endregion

    #region Create Simulator

    public Simulator CreateSimulator()
    {
        return new Simulator(
            m_Epics,
            m_SprintCapacities,
            m_InitialSprintStart,
            m_SprintDays,
            m_MaxSprintCount,
            m_SprintOffset,
            m_PlannedHours);
    } 

    #endregion
}