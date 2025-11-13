using EpicPlanner.Core.Configuration;
using EpicPlanner.Core.Planner.Services;
using EpicPlanner.Core.Planner.Simulation;
using EpicPlanner.Core.Shared.Models;
using EpicPlanner.Core.Shared.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EpicPlanner.Core.Planner;

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

    public async Task RunAsync(EnumPlanningMode _enumMode)
    {
        bool includePlannedHours = _enumMode == EnumPlanningMode.Standard;

        PlannerPlanningSnapshot snapshot = await m_DataProvider.LoadPlannerSnapshotAsync(_bIncludePlannedHours: includePlannedHours);

        if (_enumMode == EnumPlanningMode.Analysis)
        {
            HashSet<string> analysisScope = BuildAnalysisScope(snapshot.Epics);
            PlannerSimulator simulator = snapshot.CreatePlannerSimulator(epic => analysisScope.Contains(epic.Name));
            simulator.Run();
            AlignAnalysisEpicsToEndDates(
                simulator.Epics,
                simulator.InitialSprintDate,
                simulator.SprintLengthDays,
                simulator.MaxSprintCount);

            simulator.ExportGanttSprintBased(
                m_AppConfiguration.FileConfiguration.OutputPngFilePath,
                EnumPlanningMode.Analysis);
            return;
        }

        bool onlyDevelopmentEpics = m_AppConfiguration.PlannerConfiguration.OnlyDevelopmentEpics;
        PlannerSimulator standardSimulator = snapshot.CreatePlannerSimulator(
            _OnlyDevelopmentEpics: onlyDevelopmentEpics);
        standardSimulator.Run();
        standardSimulator.ExportPlanningExcel(m_AppConfiguration.FileConfiguration.OutputFilePath);
        standardSimulator.ExportGanttSprintBased(
            m_AppConfiguration.FileConfiguration.OutputPngFilePath,
            EnumPlanningMode.Standard,
            onlyDevelopmentEpics);
    }

    #endregion

    #region Helpers

    private static HashSet<string> BuildAnalysisScope(IReadOnlyList<Epic> _Epics)
    {
        Dictionary<string, Epic> byName = _Epics
            .GroupBy(e => e.Name, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .ToDictionary(e => e.Name, StringComparer.OrdinalIgnoreCase);

        HashSet<string> included = new(StringComparer.OrdinalIgnoreCase);
        Queue<Epic> queue = new();

        foreach (Epic epic in byName.Values)
        {
            if (IsAnalysisState(epic))
            {
                queue.Enqueue(epic);
            }
        }

        while (queue.Count > 0)
        {
            Epic current = queue.Dequeue();
            if (!included.Add(current.Name))
            {
                continue;
            }

            foreach (string dependencyName in current.Dependencies)
            {
                if (!byName.TryGetValue(dependencyName, out Epic? dependency))
                {
                    continue;
                }

                if (!included.Contains(dependency.Name) &&
                    (IsAnalysisState(dependency) || IsDevelopmentState(dependency)))
                {
                    queue.Enqueue(dependency);
                }
            }
        }

        return included;
    }

    private static void AlignAnalysisEpicsToEndDates(
        IReadOnlyList<Epic> _Epics,
        DateTime _DateInitialSprintStart,
        int _iSprintLengthDays,
        int _iMaxSprintCount)
    {
        DateTime timelineStart = _DateInitialSprintStart.Date;
        int sprintLength = Math.Max(1, _iSprintLengthDays);
        int sprintCount = Math.Max(1, _iMaxSprintCount);
        DateTime timelineEnd = timelineStart.AddDays(sprintLength * sprintCount - 1);

        DateTime? latestAnalysisEndDate = null;
        foreach (Epic epic in _Epics)
        {
            if (IsAnalysisState(epic) && epic.EndAnalysis.HasValue)
            {
                DateTime candidate = epic.EndAnalysis.Value.Date;
                if (latestAnalysisEndDate is null || candidate > latestAnalysisEndDate.Value)
                {
                    latestAnalysisEndDate = candidate;
                }
            }
        }

        foreach (Epic epic in _Epics)
        {
            if (IsAnalysisState(epic))
            {
                bool isPending = epic.State.Contains("pending", StringComparison.OrdinalIgnoreCase);

                if (epic.EndAnalysis.HasValue)
                {
                    DateTime end = epic.EndAnalysis.Value.Date;
                    if (end < timelineStart)
                    {
                        end = timelineStart;
                    }

                    DateTime start = isPending
                        ? SprintStartForDate(end, timelineStart, sprintLength)
                        : timelineStart;

                    epic.StartDate = start;
                    epic.EndDate = end;
                }
                else
                {
                    DateTime cappedEnd = latestAnalysisEndDate ?? timelineEnd;
                    if (cappedEnd < timelineStart)
                    {
                        cappedEnd = timelineStart;
                    }

                    epic.StartDate = timelineStart;
                    epic.EndDate = cappedEnd;
                }

                continue;
            }

            // Non-analysis epics (development dependencies, etc.) are kept in the
            // simulation for scheduling purposes but must not appear in the analysis
            // Gantt. Clearing their dates prevents them from being rendered.
            epic.StartDate = null;
            epic.EndDate = null;
        }
    }

    private static DateTime SprintStartForDate(DateTime _Date, DateTime _DateInitialSprintStart, int _iSprintLengthDays)
    {
        if (_iSprintLengthDays <= 0)
        {
            return _DateInitialSprintStart;
        }

        double totalDays = (_Date.Date - _DateInitialSprintStart.Date).TotalDays;
        int sprintIndex = (int)Math.Floor(totalDays / _iSprintLengthDays);
        if (sprintIndex < 0)
        {
            sprintIndex = 0;
        }

        return _DateInitialSprintStart.Date.AddDays(sprintIndex * _iSprintLengthDays);
    }

    private static bool IsAnalysisState(Epic _Epic) =>
        _Epic != null &&
        !string.IsNullOrWhiteSpace(_Epic.State) &&
        _Epic.State.Contains("analysis", StringComparison.OrdinalIgnoreCase);

    private static bool IsDevelopmentState(Epic _Epic) =>
        _Epic != null &&
        !string.IsNullOrWhiteSpace(_Epic.State) &&
        _Epic.State.Contains("develop", StringComparison.OrdinalIgnoreCase);

    #endregion
}
