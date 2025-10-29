using System;
using System.Linq;

namespace EpicPlanner.Core;

public class PlanningSnapshot
{
    #region Members

    private readonly List<Epic> m_Epics;
    private readonly Dictionary<int, Dictionary<string, ResourceCapacity>> m_SprintCapacities;
    private readonly DateTime m_InitialSprintStart;
    private readonly int m_iSprintDays;
    private readonly int m_iMaxSprintCount;
    private readonly int m_iSprintOffset;
    private readonly Dictionary<string, double> m_PlannedHours;
    private readonly IReadOnlyList<SprintEpicSummary> m_EpicSummaries;
    private readonly Dictionary<string, double> m_PlannedCapacityByEpic;

    #endregion

    #region Constructor

    public PlanningSnapshot(
        List<Epic> _Epics,
        Dictionary<int, Dictionary<string, ResourceCapacity>> _SprintCapacities,
        DateTime _InitialSprintStart,
        int _iSprintDays,
        int _iMaxSprintCount,
        int _iSprintOffset,
        Dictionary<string, double> _PlannedHours,
        IReadOnlyList<SprintEpicSummary> _EpicSummaries,
        IReadOnlyDictionary<string, double> _PlannedCapacityByEpic)
    {
        m_Epics = _Epics;
        m_SprintCapacities = _SprintCapacities;
        m_InitialSprintStart = _InitialSprintStart;
        m_iSprintDays = _iSprintDays;
        m_iMaxSprintCount = _iMaxSprintCount;
        m_iSprintOffset = _iSprintOffset;
        m_PlannedHours = _PlannedHours;
        m_EpicSummaries = _EpicSummaries;
        m_PlannedCapacityByEpic = _PlannedCapacityByEpic != null
            ? new Dictionary<string, double>(_PlannedCapacityByEpic, StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
    }

    #endregion

    #region Create Simulator

    public Simulator CreateSimulator(Func<Epic, bool>? _Filter = null)
    {
        List<Epic> epics = _Filter is null
            ? m_Epics
            : m_Epics.Where(_Filter).ToList();

        return new Simulator(
            epics,
            m_SprintCapacities,
            m_InitialSprintStart,
            m_iSprintDays,
            m_iMaxSprintCount,
            m_iSprintOffset,
            m_PlannedHours,
            m_EpicSummaries,
            m_PlannedCapacityByEpic);
    }

    #endregion
}