using EpicPlanner.Core.Planner.Simulation;
using EpicPlanner.Core.Shared.Models;
using EpicPlanner.Core.Shared.Services;
using System;
using System.Collections.Generic;

namespace EpicPlanner.Core.Planner.Services;

public class PlannerPlanningSnapshot : PlanningSnapshotBase
{
    #region Constructor

    public PlannerPlanningSnapshot(
        List<Epic> _Epics,
        Dictionary<int, Dictionary<string, ResourceCapacity>> _SprintCapacities,
        DateTime _InitialSprintStart,
        int _iSprintDays,
        int _iMaxSprintCount,
        int _iSprintOffset)
        : base(
            _Epics,
            _SprintCapacities,
            _InitialSprintStart,
            _iSprintDays,
            _iMaxSprintCount,
            _iSprintOffset)
    {
    }

    #endregion

    #region Create Simulator

    public PlannerSimulator CreatePlannerSimulator(Func<Epic, bool>? _Filter = null)
    {
        List<Epic> epics = FilterEpics(_Filter);

        return new PlannerSimulator(
            epics,
            SprintCapacities,
            InitialSprintStart,
            SprintDays,
            MaxSprintCount,
            SprintOffset);
    }

    #endregion
}
