using System;
using System.Collections.Generic;
using System.Linq;
using EpicPlanner.Core.Shared.Models;
using OfficeOpenXml;
using System.Text.RegularExpressions;

namespace EpicPlanner.Core.Shared.Simulation;

public abstract class SimulatorBase
{
    #region Members

    private readonly List<Epic> m_Epics;
    private readonly Dictionary<int, Dictionary<string, ResourceCapacity>> m_SprintCapacities;
    private readonly DateTime m_InitialSprintDate;
    private readonly int m_iSprintDays;
    private readonly int m_iMaxSprintCount;
    private readonly int m_iSprintOffset;
    private readonly bool m_bOnlyDevelopmentEpics;
    private readonly Dictionary<string, DateTime> m_CompletedMap = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<Allocation> m_Allocations = new();
    private readonly List<(int Sprint, string Resource, double Unused, string Reason)> m_Underutilization = new();
    private readonly Dictionary<string, Epic> m_EpicsByName = new(StringComparer.OrdinalIgnoreCase);

    #endregion

    #region Constructor

    protected SimulatorBase(
        List<Epic> _Epics,
        Dictionary<int, Dictionary<string, ResourceCapacity>> _SprintCapacities,
        DateTime _InitialSprintDate,
        int _iSprintDays,
        int _iMaxSprintCount,
        int _iSprintOffset,
        bool _bOnlyDevelopmentEpics = false)
    {
        m_Epics = _Epics;
        m_SprintCapacities = _SprintCapacities;
        m_InitialSprintDate = _InitialSprintDate;
        m_iSprintDays = _iSprintDays;
        m_iMaxSprintCount = _iMaxSprintCount;
        m_iSprintOffset = _iSprintOffset;
        m_bOnlyDevelopmentEpics = _bOnlyDevelopmentEpics;

        foreach (var epic in _Epics.Where(e => e.Remaining <= 0))
        {
            m_CompletedMap[epic.Name] = epic.EndDate ?? _InitialSprintDate.AddDays(-1);
        }

        foreach (var epic in _Epics)
        {
            if (!m_EpicsByName.ContainsKey(epic.Name))
            {
                m_EpicsByName[epic.Name] = epic;
            }
        }
    }

    #endregion

    #region Properties

    public IReadOnlyList<Epic> Epics => m_Epics;
    public DateTime InitialSprintDate => m_InitialSprintDate;
    public int SprintLengthDays => m_iSprintDays;
    public int MaxSprintCount => m_iMaxSprintCount;
    protected bool OnlyDevelopmentEpics => m_bOnlyDevelopmentEpics;
    protected int SprintOffset => m_iSprintOffset;
    protected IReadOnlyDictionary<int, Dictionary<string, ResourceCapacity>> SprintCapacities => m_SprintCapacities;
    protected IReadOnlyList<Allocation> AllocationHistory => m_Allocations;
    protected IReadOnlyList<(int Sprint, string Resource, double Unused, string Reason)> UnderutilizationEntries => m_Underutilization;

    #endregion

    #region Run

    public void Run()
    {
        for (int sprint = 0; sprint < m_iMaxSprintCount; sprint++)
        {
            if (m_Epics.All(e => e.Remaining <= 1e-6))
            {
                break;
            }

            var sprintStart = m_InitialSprintDate.AddDays(sprint * m_iSprintDays);
            var sprintEnd = sprintStart.AddDays(m_iSprintDays - 1);

            var resourceRemaining = m_SprintCapacities[sprint]
                .ToDictionary(
                    kv => kv.Key,
                    kv => new ResourceCapacity
                    {
                        Development = kv.Value.Development,
                        Maintenance = kv.Value.Maintenance,
                        Analysis = kv.Value.Analysis
                    },
                    StringComparer.OrdinalIgnoreCase);

            var activeDev = m_Epics
                .Where(e => e.Remaining > 1e-6 && e.IsInDevelopment && DependencySatisfied(e, sprintStart))
                .ToList();
            var activeOthers = m_bOnlyDevelopmentEpics
                ? new List<Epic>()
                : m_Epics
                    .Where(e => e.Remaining > 1e-6 && !e.IsInDevelopment && e.IsOtherAllowed && DependencySatisfied(e, sprintStart))
                    .ToList();

            foreach (var activeSet in new[] { activeDev, activeOthers })
            {
                var requests = new Dictionary<string, List<(Epic epic, double desired, double pct)>>(StringComparer.OrdinalIgnoreCase);
                foreach (var epic in activeSet)
                {
                    foreach (var wish in epic.Wishes)
                    {
                        if (!resourceRemaining.ContainsKey(wish.Resource) || wish.Percentage <= 0)
                        {
                            continue;
                        }

                        double desired = resourceRemaining[wish.Resource].Development * wish.Percentage;
                        if (!requests.TryGetValue(wish.Resource, out var list))
                        {
                            list = requests[wish.Resource] = new List<(Epic, double, double)>();
                        }
                        list.Add((epic, desired, wish.Percentage));
                    }
                }

                foreach (var kv in requests)
                {
                    string resourceName = kv.Key;
                    var reqs = kv.Value;
                    if (reqs.Count == 0)
                    {
                        continue;
                    }

                    double available = resourceRemaining[resourceName].Development;
                    var grouped = reqs
                        .GroupBy(x => x.epic.Priority)
                        .OrderByDescending(g => g.Key)
                        .ToList();

                    foreach (var group in grouped)
                    {
                        if (available <= 1e-6)
                        {
                            break;
                        }

                        var epicRequests = group.ToList();
                        if (epicRequests.Count == 1)
                        {
                            var (epic, _, _) = epicRequests[0];
                            double alloc = Math.Min(available, epic.Remaining);
                            if (alloc > 1e-9)
                            {
                                CommitAllocation(epic, sprint, resourceName, alloc, sprintStart);
                                available -= alloc;
                                resourceRemaining[resourceName].Development -= alloc;
                                if (epic.StartDate == null)
                                {
                                    epic.StartDate = sprintStart;
                                }
                                if (epic.Remaining <= 1e-6)
                                {
                                    epic.Remaining = 0;
                                    epic.EndDate = sprintStart;
                                    m_CompletedMap[epic.Name] = epic.EndDate.Value;
                                }
                            }
                            continue;
                        }

                        double totalDesired = epicRequests.Sum(r => r.desired);
                        foreach (var (epic, desired, _) in epicRequests)
                        {
                            if (available <= 1e-6)
                            {
                                break;
                            }

                            double share = totalDesired > 1e-9 ? desired / totalDesired : 1.0 / epicRequests.Count;
                            double alloc = Math.Min(available * share, epic.Remaining);
                            if (alloc <= 1e-9)
                            {
                                continue;
                            }

                            CommitAllocation(epic, sprint, resourceName, alloc, sprintStart);
                            available -= alloc;
                            resourceRemaining[resourceName].Development -= alloc;
                            if (epic.StartDate == null)
                            {
                                epic.StartDate = sprintStart;
                            }
                            if (epic.Remaining <= 1e-6)
                            {
                                epic.Remaining = 0;
                                epic.EndDate = sprintStart;
                                m_CompletedMap[epic.Name] = epic.EndDate.Value;
                            }
                        }
                    }
                }

                foreach (var resourceName in resourceRemaining.Keys.ToList())
                {
                    double leftover = resourceRemaining[resourceName].Development;
                    if (leftover <= 1e-6)
                    {
                        continue;
                    }

                    var candidates = activeSet
                        .Where(e => e.Wishes.Any(w => w.Resource.Equals(resourceName, StringComparison.OrdinalIgnoreCase)) && e.Remaining > 1e-6)
                        .ToList();

                    foreach (var epic in candidates)
                    {
                        if (leftover <= 1e-6)
                        {
                            break;
                        }

                        double alloc = Math.Min(leftover, epic.Remaining);
                        if (alloc <= 1e-9)
                        {
                            continue;
                        }

                        CommitAllocation(epic, sprint, resourceName, alloc, sprintStart);
                        leftover -= alloc;
                        resourceRemaining[resourceName].Development -= alloc;
                        if (epic.StartDate == null)
                        {
                            epic.StartDate = sprintStart;
                        }
                        if (epic.Remaining <= 1e-6)
                        {
                            epic.Remaining = 0;
                            epic.EndDate = sprintStart;
                            m_CompletedMap[epic.Name] = epic.EndDate.Value;
                        }
                    }
                }
            }

            foreach (var kv in resourceRemaining)
            {
                if (kv.Value.Development > 1e-6)
                {
                    bool resourceHadAssignedEpics = m_Epics.Any(ep => ep.Wishes.Any(w => w.Resource.Equals(kv.Key, StringComparison.OrdinalIgnoreCase)));
                    string reason = resourceHadAssignedEpics ? "no remaining hours on assigned epics" : "no assigned epics";
                    m_Underutilization.Add((sprint, kv.Key, Math.Round(kv.Value.Development, 2), reason));
                }
            }
        }
    }

    private bool DependencySatisfied(Epic _Epic, DateTime _SprintStartDate)
    {
        if (_Epic.EndAnalysis.HasValue && _SprintStartDate.Date < _Epic.EndAnalysis.Value.Date)
        {
            return false;
        }

        if (_Epic.Dependencies.Count == 0)
        {
            return true;
        }

        foreach (var dependency in _Epic.Dependencies)
        {
            if (!m_CompletedMap.TryGetValue(dependency, out var depEnd))
            {
                if (m_bOnlyDevelopmentEpics &&
                    m_EpicsByName.TryGetValue(dependency, out Epic? dependencyEpic) &&
                    !dependencyEpic.IsInDevelopment)
                {
                    continue;
                }

                return false;
            }

            if (depEnd.Date >= _SprintStartDate.Date)
            {
                return false;
            }
        }

        return true;
    }

    private void CommitAllocation(Epic _Epic, int _iSprint, string _strResource, double _dHours, DateTime _SprintStartDate)
    {
        _Epic.Remaining -= _dHours;
        var allocation = new Allocation(_Epic.Name, _iSprint, _strResource, _dHours, _SprintStartDate);
        _Epic.History.Add(allocation);
        m_Allocations.Add(allocation);
    }

    #endregion

    #region Helpers

    protected static void WriteTable(ExcelWorksheet _Worksheet, IReadOnlyList<string> _Headers)
    {
        for (int i = 0; i < _Headers.Count; i++)
        {
            _Worksheet.Cells[1, i + 1].Value = _Headers[i];
            _Worksheet.Cells[1, i + 1].Style.Font.Bold = true;
        }
    }

    protected static (int Year, int Num) ExtractEpicKey(string _strEpic)
    {
        if (string.IsNullOrWhiteSpace(_strEpic))
        {
            return (9999, 9999);
        }

        var match = Regex.Match(_strEpic, @"(\d{4})[-_ ]+(\d{1,3})");
        if (match.Success)
        {
            return (int.Parse(match.Groups[1].Value), int.Parse(match.Groups[2].Value));
        }

        var matchYear = Regex.Match(_strEpic, @"(\d{4})");
        if (matchYear.Success)
        {
            return (int.Parse(matchYear.Groups[1].Value), 0);
        }

        return (9999, 9999);
    }

    protected DateTime TryParseDate(string _strDate) => DateTime.TryParse(_strDate, out var d) ? d : DateTime.MaxValue;

    protected DateTime SprintStartDate(int _iSprintIndex) => m_InitialSprintDate.AddDays(_iSprintIndex * m_iSprintDays);

    protected float SprintPosition(DateTime _Date, bool _bIsEnd)
    {
        int sprintDays = m_iSprintDays <= 0 ? 1 : m_iSprintDays;
        double totalDays = (_Date.Date - m_InitialSprintDate.Date).TotalDays;
        if (totalDays < 0)
        {
            totalDays = 0;
        }

        if (_bIsEnd)
        {
            totalDays += 1.0;
        }

        return (float)(totalDays / sprintDays);
    }

    #endregion
}
