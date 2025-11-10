using EpicPlanner.Core.Checker.Simulation;
using EpicPlanner.Core.Shared.Models;
using EpicPlanner.Core.Shared.Services;
using System;
using System.Collections.Generic;

namespace EpicPlanner.Core.Checker.Services;

public class CheckerPlanningSnapshot : PlanningSnapshotBase
{
    #region Members

    private readonly Dictionary<string, ResourcePlannedHoursBreakdown> m_PlannedHours;
    private readonly IReadOnlyList<SprintEpicSummary> m_EpicSummaries;
    private readonly Dictionary<string, double> m_PlannedCapacityByEpic;

    #endregion

    #region Constructor

    public CheckerPlanningSnapshot(
        List<Epic> _Epics,
        Dictionary<int, Dictionary<string, ResourceCapacity>> _SprintCapacities,
        DateTime _InitialSprintStart,
        int _iSprintDays,
        int _iMaxSprintCount,
        int _iSprintOffset,
        Dictionary<string, ResourcePlannedHoursBreakdown> _PlannedHours,
        IReadOnlyList<SprintEpicSummary> _EpicSummaries,
        IReadOnlyDictionary<string, double> _PlannedCapacityByEpic)
        : base(
            _Epics,
            _SprintCapacities,
            _InitialSprintStart,
            _iSprintDays,
            _iMaxSprintCount,
            _iSprintOffset)
    {
        m_PlannedHours = _PlannedHours;
        m_EpicSummaries = _EpicSummaries;
        m_PlannedCapacityByEpic = _PlannedCapacityByEpic != null
            ? new Dictionary<string, double>(_PlannedCapacityByEpic, StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
    }

    #endregion

    #region Properties

    public IReadOnlyDictionary<string, ResourcePlannedHoursBreakdown> PlannedHours => m_PlannedHours;

    public IReadOnlyList<SprintEpicSummary> EpicSummaries => m_EpicSummaries;

    public IReadOnlyDictionary<string, double> PlannedCapacityByEpic => m_PlannedCapacityByEpic;

    #endregion

    #region Create Simulator

    public CheckerSimulator CreateCheckerSimulator(Func<Epic, bool>? _Filter = null)
    {
        List<Epic> epics = FilterEpics(_Filter);

        return new CheckerSimulator(
            epics,
            SprintCapacities,
            InitialSprintStart,
            SprintDays,
            MaxSprintCount,
            SprintOffset,
            m_PlannedHours,
            m_EpicSummaries,
            m_PlannedCapacityByEpic);
    }

    #endregion
}
