using System;
using System.Collections.Generic;
using System.Linq;

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

        if (_Mode == PlanningMode.Analysis)
        {
            HashSet<string> analysisScope = BuildAnalysisScope(snapshot.Epics);
            Simulator simulator = snapshot.CreateSimulator(epic => analysisScope.Contains(epic.Name));
            simulator.Run();
            AlignAnalysisEpicsToEndDates(simulator.Epics);

            simulator.ExportGanttSprintBased(
                m_AppConfiguration.FileConfiguration.OutputPngFilePath,
                PlanningMode.Analysis);
            return;
        }

        Simulator standardSimulator = snapshot.CreateSimulator();
        standardSimulator.Run();
        standardSimulator.ExportPlanningExcel(m_AppConfiguration.FileConfiguration.OutputFilePath);
        standardSimulator.ExportGanttSprintBased(
            m_AppConfiguration.FileConfiguration.OutputPngFilePath,
            PlanningMode.Standard);
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

    private static void AlignAnalysisEpicsToEndDates(IReadOnlyList<Epic> _Epics)
    {
        foreach (Epic epic in _Epics)
        {
            if (IsAnalysisState(epic))
            {
                if (epic.EndAnalysis.HasValue)
                {
                    DateTime end = epic.EndAnalysis.Value.Date;
                    epic.StartDate = end;
                    epic.EndDate = end;
                }
                else
                {
                    epic.StartDate = null;
                    epic.EndDate = null;
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
