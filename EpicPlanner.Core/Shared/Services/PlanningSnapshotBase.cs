using EpicPlanner.Core.Shared.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EpicPlanner.Core.Shared.Services;

public abstract class PlanningSnapshotBase
{
    #region Members

    private readonly List<Epic> m_Epics;
    private readonly Dictionary<int, Dictionary<string, ResourceCapacity>> m_SprintCapacities;
    private readonly DateTime m_InitialSprintStart;
    private readonly int m_iSprintDays;
    private readonly int m_iMaxSprintCount;
    private readonly int m_iSprintOffset;

    #endregion

    #region Constructor

    protected PlanningSnapshotBase(
        List<Epic> _Epics,
        Dictionary<int, Dictionary<string, ResourceCapacity>> _SprintCapacities,
        DateTime _InitialSprintStart,
        int _iSprintDays,
        int _iMaxSprintCount,
        int _iSprintOffset)
    {
        m_Epics = _Epics;
        m_SprintCapacities = _SprintCapacities;
        m_InitialSprintStart = _InitialSprintStart;
        m_iSprintDays = _iSprintDays;
        m_iMaxSprintCount = _iMaxSprintCount;
        m_iSprintOffset = _iSprintOffset;
    }

    #endregion

    #region Properties

    public IReadOnlyList<Epic> Epics => m_Epics;

    protected Dictionary<int, Dictionary<string, ResourceCapacity>> SprintCapacities => m_SprintCapacities;

    protected DateTime InitialSprintStart => m_InitialSprintStart;

    protected int SprintDays => m_iSprintDays;

    protected int MaxSprintCount => m_iMaxSprintCount;

    protected int SprintOffset => m_iSprintOffset;

    #endregion

    #region Helpers

    protected List<Epic> FilterEpics(Func<Epic, bool>? _Filter)
    {
        return _Filter is null
            ? m_Epics
            : m_Epics.Where(_Filter).ToList();
    }

    #endregion
}
