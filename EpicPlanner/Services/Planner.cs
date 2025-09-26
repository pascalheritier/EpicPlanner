using Microsoft.Extensions.Logging;
using OfficeOpenXml;

namespace EpicPlanner;

internal class Planner
{
    #region Members

    private readonly DateTime m_InitialSprintStart;
    private readonly int m_iSprintDays;
    private readonly int m_iSprintCapacityDays;
    private readonly AppConfiguration m_AppConfiguration;

    #endregion

    #region Constructor

    public Planner(AppConfiguration _AppConfiguration, ILoggerFactory _LoggerFactory)
    {
        m_AppConfiguration = _AppConfiguration ?? throw new ArgumentNullException(nameof(_AppConfiguration));
        this.m_InitialSprintStart = _AppConfiguration.PlannerConfiguration.InitialSprintStart;
        this.m_iSprintDays = _AppConfiguration.PlannerConfiguration.SprintDays;
        this.m_iSprintCapacityDays = _AppConfiguration.PlannerConfiguration.SprintCapacityDays;
    }

    #endregion

    #region Planning logic

    public async Task RunAsync()
    {
        using var package = new ExcelPackage(new FileInfo(m_AppConfiguration.FileConfiguration.InputFilePath));
        var wsEpics = package.Workbook.Worksheets[m_AppConfiguration.FileConfiguration.InputEpicsSheetName];
        var wsRes = package.Workbook.Worksheets[m_AppConfiguration.FileConfiguration.InputResourcesSheetName];

        var resources = LoadResources(wsRes);
        var epics = LoadEpics(wsEpics, resources.Keys.ToList());

        // Adjust resources for absences (for each sprint)
        var absFetcher = new AbsenceFetcher(
            m_AppConfiguration.RedmineConfiguration.ServerUrl,
            m_AppConfiguration.RedmineConfiguration.ApiKey);
        Dictionary<int, Dictionary<string, ResourceCapacity>> adjustedCapacities = await AdjustCapacitiesForAbsencesAsync(
            resources,
            absFetcher,
            m_AppConfiguration.PlannerConfiguration.Holidays);

        var simulator = new Simulator(epics, adjustedCapacities, m_InitialSprintStart, m_iSprintDays, m_AppConfiguration.PlannerConfiguration.MaxSprintCount);
        simulator.Run();

        simulator.ExportExcel(m_AppConfiguration.FileConfiguration.OutputFilePath);
        simulator.ExportGanttSprintBased(m_AppConfiguration.FileConfiguration.OutputPngFilePath);
    }

    private Dictionary<string, ResourceCapacity> LoadResources(ExcelWorksheet _ResourceWorksheet)
    {
        int rows = _ResourceWorksheet.Dimension.End.Row;
        var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        for (int c = 1; c <= _ResourceWorksheet.Dimension.End.Column; c++)
        {
            var h = _ResourceWorksheet.Cells[1, c].GetValue<string>()?.Trim();
            if (!string.IsNullOrWhiteSpace(h)) headers[h] = c;
        }

        int nameCol = headers.ContainsKey("Ingénieur") ? headers["Ingénieur"] :
                       headers.ContainsKey("Engineer") ? headers["Engineer"] : 1;

        int devCol = headers.ContainsKey("Heures Dév. Epic") ? headers["Heures Dév. Epic"] : 2;
        int maintCol = headers.ContainsKey("Heures maintenance") ? headers["Heures maintenance"] : 0;
        int analCol = headers.ContainsKey("Heures Analyse Epic") ? headers["Heures Analyse Epic"] : 0;

        var dict = new Dictionary<string, ResourceCapacity>(StringComparer.OrdinalIgnoreCase);
        for (int r = 2; r <= rows; r++)
        {
            string name = _ResourceWorksheet.Cells[r, nameCol].GetValue<string>()?.Trim();
            if (string.IsNullOrWhiteSpace(name)) continue;

            double dev = _ResourceWorksheet.Cells[r, devCol].GetValue<double>();
            double maint = maintCol > 0 ? _ResourceWorksheet.Cells[r, maintCol].GetValue<double>() : 0;
            double anal = analCol > 0 ? _ResourceWorksheet.Cells[r, analCol].GetValue<double>() : 0;

            dict[name] = new ResourceCapacity
            {
                Development = dev,
                Maintenance = maint,
                Analysis = anal
            };
        }
        return dict;
    }

    private async Task<Dictionary<int, Dictionary<string, ResourceCapacity>>> AdjustCapacitiesForAbsencesAsync(
        Dictionary<string, ResourceCapacity> _BaseCapacities,
        AbsenceFetcher _AbsenceFetcher,
        IEnumerable<DateTime> _Holidays)
    {
        Dictionary<int, Dictionary<string, ResourceCapacity>> adjustedCapacities = new();
        Dictionary<string, List<(DateTime, DateTime)>> absencesPerResource = await _AbsenceFetcher.GetResourcesAbsencesAsync();

        for (int sprint = 0; sprint < m_AppConfiguration.PlannerConfiguration.MaxSprintCount; sprint++)
        {
            var sprintStart = m_InitialSprintStart.AddDays(sprint * m_iSprintDays).Date;
            var sprintEnd = sprintStart.AddDays(m_iSprintDays - 1).Date;

            // Compute how many working days exist in this sprint
            int workingDaysInSprint = BusinessCalendar.CountWorkingDays(sprintStart, sprintEnd, _Holidays);

            var sprintCapacities = new Dictionary<string, ResourceCapacity>(StringComparer.OrdinalIgnoreCase);

            foreach (var user in _BaseCapacities.Keys)
            {
                // Start from base sprint hours, scaled to real working days
                ResourceCapacity userSprintCapacity = new(_BaseCapacities[user]);
                double scale = (double)workingDaysInSprint / m_iSprintCapacityDays;
                userSprintCapacity.AdapteCapacityToScale(scale);

                if (absencesPerResource.TryGetValue(user, out var absList))
                {
                    foreach (var (start, end) in absList)
                    {
                        int absentWD = BusinessCalendar.CountWorkingDaysOverlap(start, end, sprintStart, sprintEnd, _Holidays);
                        if (absentWD > 0 && workingDaysInSprint > 0)
                        {
                            userSprintCapacity.AdaptCapacityToAbsences(workingDaysInSprint, absentWD);
                        }
                    }
                }
                userSprintCapacity.RoundUpCapacity();
                sprintCapacities[user] = userSprintCapacity;
            }
            adjustedCapacities[sprint] = sprintCapacities;
        }

        return adjustedCapacities;
    }


    private List<Epic> LoadEpics(ExcelWorksheet _EpicWorksheet, List<string> _ResourceNames)
    {
        // Detect headers
        var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int c = 1; c <= _EpicWorksheet.Dimension.End.Column; c++)
        {
            var h = _EpicWorksheet.Cells[1, c].GetValue<string>()?.Trim();
            if (!string.IsNullOrWhiteSpace(h)) headers[h] = c;
        }

        int epicCol = headers.ContainsKey("Epic name") ? headers["Epic name"] : 1;
        int stateCol = headers.ContainsKey("State") ? headers["State"] : 2;
        int remainingCol = headers.FirstOrDefault(kv => kv.Key.Contains("Remaining", StringComparison.OrdinalIgnoreCase)).Value;
        int roughCol = headers.FirstOrDefault(kv => kv.Key.Contains("Rough", StringComparison.OrdinalIgnoreCase)).Value;
        int assignedCol = headers.FirstOrDefault(kv => kv.Key.Contains("Assigned to", StringComparison.OrdinalIgnoreCase)).Value;
        int willAssignCol = headers.FirstOrDefault(kv => kv.Key.Contains("Will be assigned", StringComparison.OrdinalIgnoreCase) || kv.Key.Contains("Will be assigne", StringComparison.OrdinalIgnoreCase)).Value;
        int depCol = headers.FirstOrDefault(kv => kv.Key.Contains("Epic dependency", StringComparison.OrdinalIgnoreCase) || kv.Key.Contains("Dependency", StringComparison.OrdinalIgnoreCase)).Value;
        int endAnalysisCol = headers.FirstOrDefault(kv => kv.Key.Contains("End of analysis", StringComparison.OrdinalIgnoreCase)).Value;

        int rows = _EpicWorksheet.Dimension.End.Row;
        var epics = new List<Epic>();

        for (int r = 2; r <= rows; r++)
        {
            string epicName = _EpicWorksheet.Cells[r, epicCol].GetValue<string>()?.Trim();
            if (string.IsNullOrWhiteSpace(epicName)) continue;

            string state = _EpicWorksheet.Cells[r, stateCol].GetValue<string>() ?? "";
            double charge = 0.0;
            double remVal = (remainingCol > 0) ? _EpicWorksheet.Cells[r, remainingCol].GetValue<double>() : 0.0;
            double roughVal = (roughCol > 0) ? _EpicWorksheet.Cells[r, roughCol].GetValue<double>() : 0.0;

            if (remVal > 0)
                charge = remVal;
            else if (roughVal > 0)
                charge = roughVal;

            string assigned = assignedCol > 0 ? (_EpicWorksheet.Cells[r, assignedCol].GetValue<string>() ?? "") : "";
            string willAssign = willAssignCol > 0 ? (_EpicWorksheet.Cells[r, willAssignCol].GetValue<string>() ?? "") : "";
            string depRaw = depCol > 0 ? (_EpicWorksheet.Cells[r, depCol].GetValue<string>() ?? "") : "";
            string endAnalysisStr = endAnalysisCol > 0 ? (_EpicWorksheet.Cells[r, endAnalysisCol].GetValue<string>() ?? "") : "";

            DateTime? endAnalysis = null;
            if (DateTime.TryParse(endAnalysisStr, out var parsed))
                endAnalysis = parsed;

            var epic = new Epic(epicName, state, charge, endAnalysis);
            epic.ParseAssignments(assigned, willAssign, _ResourceNames);
            epic.ParseDependencies(depRaw);

            // FIX #2: handle 0h epics
            if (epic.Charge <= 0)
            {
                epic.Remaining = 0;
                epic.StartDate = endAnalysis ?? this.m_InitialSprintStart;
                epic.EndDate = endAnalysis ?? this.m_InitialSprintStart;
            }

            epics.Add(epic);
        }

        // Normalize dependency names to actual epic names (case-insensitive contains/equals)
        var epicNames = epics.Select(e => e.Name).ToList();
        foreach (var e in epics)
        {
            for (int i = 0; i < e.Dependencies.Count; i++)
            {
                var d = e.Dependencies[i];
                var match = epicNames.FirstOrDefault(x =>
                    x.Equals(d, StringComparison.OrdinalIgnoreCase) ||
                    x.IndexOf(d, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    d.IndexOf(x, StringComparison.OrdinalIgnoreCase) >= 0);
                if (!string.IsNullOrWhiteSpace(match))
                    e.Dependencies[i] = match;
            }
        }

        return epics;
    }

    #endregion
}