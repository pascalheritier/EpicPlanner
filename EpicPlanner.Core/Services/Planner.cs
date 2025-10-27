using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace EpicPlanner.Core;

public class PlanningDataProvider
{
    private readonly AppConfiguration m_AppConfiguration;
    private readonly DateTime m_InitialSprintStart;
    private readonly int m_SprintDays;
    private readonly int m_SprintCapacityDays;

    public PlanningDataProvider(AppConfiguration appConfiguration)
    {
        m_AppConfiguration = appConfiguration ?? throw new ArgumentNullException(nameof(appConfiguration));
        m_InitialSprintStart = appConfiguration.PlannerConfiguration.InitialSprintStartDate;
        m_SprintDays = appConfiguration.PlannerConfiguration.SprintDays;
        m_SprintCapacityDays = appConfiguration.PlannerConfiguration.SprintCapacityDays;
    }

    public async Task<PlanningSnapshot> LoadAsync(bool includePlannedHours)
    {
        using ExcelPackage package = new(new FileInfo(m_AppConfiguration.FileConfiguration.InputFilePath));
        ExcelWorksheet wsEpics = package.Workbook.Worksheets[m_AppConfiguration.FileConfiguration.InputEpicsSheetName]
            ?? throw new NullReferenceException($"Worksheet '{m_AppConfiguration.FileConfiguration.InputEpicsSheetName}' could not be found.");
        ExcelWorksheet wsRes = package.Workbook.Worksheets[m_AppConfiguration.FileConfiguration.InputResourcesSheetName]
            ?? throw new NullReferenceException($"Worksheet '{m_AppConfiguration.FileConfiguration.InputResourcesSheetName}' could not be found.");

        Dictionary<string, ResourceCapacity> resources = LoadResources(wsRes);
        List<Epic> epics = LoadEpics(wsEpics, resources.Keys.ToList());

        RedmineDataFetcher redmineDataFetcher = new(
            m_AppConfiguration.RedmineConfiguration.ServerUrl,
            m_AppConfiguration.RedmineConfiguration.ApiKey);

        Dictionary<string, List<(DateTime, DateTime)>> absencesPerResource = await redmineDataFetcher.GetResourcesAbsencesAsync();
        Dictionary<int, Dictionary<string, ResourceCapacity>> adjustedCapacities = AdjustCapacitiesForAbsences(
            resources,
            absencesPerResource,
            m_AppConfiguration.PlannerConfiguration.Holidays);

        Dictionary<string, double> plannedHours = new(StringComparer.OrdinalIgnoreCase);
        if (includePlannedHours)
        {
            plannedHours = await redmineDataFetcher.GetPlannedHoursForSprintAsync(
                m_AppConfiguration.PlannerConfiguration.InitialSprintNumber,
                m_InitialSprintStart,
                m_InitialSprintStart.AddDays(m_SprintDays - 1));
        }

        return new PlanningSnapshot(
            epics,
            adjustedCapacities,
            m_InitialSprintStart,
            m_SprintDays,
            m_AppConfiguration.PlannerConfiguration.MaxSprintCount,
            m_AppConfiguration.PlannerConfiguration.InitialSprintNumber,
            plannedHours);
    }

    private Dictionary<string, ResourceCapacity> LoadResources(ExcelWorksheet resourceWorksheet)
    {
        int rows = resourceWorksheet.Dimension.End.Row;
        var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        for (int c = 1; c <= resourceWorksheet.Dimension.End.Column; c++)
        {
            var h = resourceWorksheet.Cells[2, c].GetValue<string>()?.Trim();
            if (!string.IsNullOrWhiteSpace(h)) headers[h] = c;
        }

        int nameCol = headers.ContainsKey("Ingénieur") ? headers["Ingénieur"] :
                       headers.ContainsKey("Engineer") ? headers["Engineer"] : 1;

        int devCol = headers.ContainsKey("Heures Réalisation Epic") ? headers["Heures Réalisation Epic"] : 2;
        int maintCol = headers.ContainsKey("Heures maintenance") ? headers["Heures maintenance"] : 0;
        int analCol = headers.ContainsKey("Heures Analyse Epic") ? headers["Heures Analyse Epic"] : 0;

        var dict = new Dictionary<string, ResourceCapacity>(StringComparer.OrdinalIgnoreCase);
        for (int row = 3; row <= rows; row++)
        {
            string name = resourceWorksheet.Cells[row, nameCol].GetValue<string>()?.Trim();
            if (string.IsNullOrWhiteSpace(name)) continue;

            double dev = resourceWorksheet.Cells[row, devCol].GetValue<double>();
            double maint = maintCol > 0 ? resourceWorksheet.Cells[row, maintCol].GetValue<double>() : 0;
            double anal = analCol > 0 ? resourceWorksheet.Cells[row, analCol].GetValue<double>() : 0;

            dict[name] = new ResourceCapacity
            {
                Development = dev,
                Maintenance = maint,
                Analysis = anal
            };
        }
        return dict;
    }

    private Dictionary<int, Dictionary<string, ResourceCapacity>> AdjustCapacitiesForAbsences(
        Dictionary<string, ResourceCapacity> baseCapacities,
        Dictionary<string, List<(DateTime, DateTime)>> absencesPerResource,
        IEnumerable<DateTime> holidays)
    {
        Dictionary<int, Dictionary<string, ResourceCapacity>> adjustedCapacities = new();
        for (int sprint = 0; sprint < m_AppConfiguration.PlannerConfiguration.MaxSprintCount; sprint++)
        {
            var sprintStart = m_InitialSprintStart.AddDays(sprint * m_SprintDays).Date;
            var sprintEnd = sprintStart.AddDays(m_SprintDays - 1).Date;

            int workingDaysInSprint = BusinessCalendar.CountWorkingDays(sprintStart, sprintEnd, holidays);

            var sprintCapacities = new Dictionary<string, ResourceCapacity>(StringComparer.OrdinalIgnoreCase);

            foreach (var user in baseCapacities.Keys)
            {
                ResourceCapacity userSprintCapacity = new(baseCapacities[user]);
                double scale = (double)workingDaysInSprint / m_SprintCapacityDays;
                userSprintCapacity.AdapteCapacityToScale(scale);

                if (absencesPerResource.TryGetValue(user, out var absList))
                {
                    foreach (var (start, end) in absList)
                    {
                        int absentWorkingDays = BusinessCalendar.CountWorkingDaysOverlap(start, end, sprintStart, sprintEnd, holidays);
                        if (absentWorkingDays > 0 && workingDaysInSprint > 0)
                        {
                            userSprintCapacity.AdaptCapacityToAbsences(workingDaysInSprint, absentWorkingDays);
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

    private List<Epic> LoadEpics(ExcelWorksheet epicWorksheet, List<string> resourceNames)
    {
        var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int c = 1; c <= epicWorksheet.Dimension.End.Column; c++)
        {
            var h = epicWorksheet.Cells[1, c].GetValue<string>()?.Trim();
            if (!string.IsNullOrWhiteSpace(h)) headers[h] = c;
        }

        int epicCol = headers.ContainsKey("Epic name") ? headers["Epic name"] : 1;
        int stateCol = headers.ContainsKey("State") ? headers["State"] : 2;
        int remainingCol = headers.FirstOrDefault(kv => kv.Key.Contains("Remaining", StringComparison.OrdinalIgnoreCase)).Value;
        int roughCol = headers.FirstOrDefault(kv => kv.Key.Contains("Rough", StringComparison.OrdinalIgnoreCase)).Value;
        int assignedCol = headers.FirstOrDefault(kv => kv.Key.Contains("Assigned to", StringComparison.OrdinalIgnoreCase)).Value;
        int willAssignCol = headers.FirstOrDefault(kv => kv.Key.Contains("Will be assigned", StringComparison.OrdinalIgnoreCase)
            || kv.Key.Contains("Will be assigne", StringComparison.OrdinalIgnoreCase)).Value;
        int priorityCol = headers.ContainsKey("Priority") ? headers["Priority"] : 11;
        int depCol = headers.FirstOrDefault(kv => kv.Key.Contains("Epic dependency", StringComparison.OrdinalIgnoreCase)
            || kv.Key.Contains("Dependency", StringComparison.OrdinalIgnoreCase)).Value;
        int endAnalysisCol = headers.FirstOrDefault(kv => kv.Key.Contains("End of analysis", StringComparison.OrdinalIgnoreCase)).Value;

        int rows = epicWorksheet.Dimension.End.Row;
        var epics = new List<Epic>();

        for (int row = 2; row <= rows; row++)
        {
            string epicName = epicWorksheet.Cells[row, epicCol].GetValue<string>()?.Trim();
            if (string.IsNullOrWhiteSpace(epicName)) continue;

            string state = epicWorksheet.Cells[row, stateCol].GetValue<string>() ?? string.Empty;
            double charge = 0.0;
            double remVal = remainingCol > 0 ? epicWorksheet.Cells[row, remainingCol].GetValue<double>() : 0.0;
            double roughVal = roughCol > 0 ? epicWorksheet.Cells[row, roughCol].GetValue<double>() : 0.0;

            if (remVal > 0)
                charge = remVal;
            else if (roughVal > 0)
                charge = roughVal;

            string assigned = assignedCol > 0 ? (epicWorksheet.Cells[row, assignedCol].GetValue<string>() ?? string.Empty) : string.Empty;
            string willAssign = willAssignCol > 0 ? (epicWorksheet.Cells[row, willAssignCol].GetValue<string>() ?? string.Empty) : string.Empty;
            string depRaw = depCol > 0 ? (epicWorksheet.Cells[row, depCol].GetValue<string>() ?? string.Empty) : string.Empty;
            string endAnalysisStr = endAnalysisCol > 0 ? (epicWorksheet.Cells[row, endAnalysisCol].GetValue<string>() ?? string.Empty) : string.Empty;
            string priorityStr = priorityCol > 0 ? (epicWorksheet.Cells[row, priorityCol].Text?.Trim() ?? "Normal") : "Normal";
            EpicPriority priority = priorityStr.ToLower() switch
            {
                "urgent" => EpicPriority.Urgent,
                "high" => EpicPriority.High,
                _ => EpicPriority.Normal
            };
            DateTime? endAnalysis = null;
            if (DateTime.TryParse(endAnalysisStr, out var parsed))
                endAnalysis = parsed;

            var epic = new Epic(epicName, state, charge, endAnalysis)
            {
                Priority = priority
            };
            epic.ParseAssignments(assigned, willAssign, resourceNames);
            epic.ParseDependencies(depRaw);

            if (epic.Charge <= 0)
            {
                epic.Remaining = 0;
                epic.StartDate = endAnalysis ?? m_InitialSprintStart;
                epic.EndDate = endAnalysis ?? m_InitialSprintStart;
            }

            epics.Add(epic);
        }

        var epicNames = epics.Select(e => e.Name).ToList();
        foreach (var epic in epics)
        {
            for (int i = 0; i < epic.Dependencies.Count; i++)
            {
                var dependency = epic.Dependencies[i];
                var match = epicNames.FirstOrDefault(x =>
                    x.Equals(dependency, StringComparison.OrdinalIgnoreCase) ||
                    x.IndexOf(dependency, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    dependency.IndexOf(x, StringComparison.OrdinalIgnoreCase) >= 0);
                if (!string.IsNullOrWhiteSpace(match))
                    epic.Dependencies[i] = match;
            }
        }

        return epics;
    }
}

public class PlanningRunner
{
    private readonly PlanningDataProvider m_DataProvider;
    private readonly AppConfiguration m_AppConfiguration;

    public PlanningRunner(PlanningDataProvider dataProvider, AppConfiguration appConfiguration)
    {
        m_DataProvider = dataProvider;
        m_AppConfiguration = appConfiguration;
    }

    public async Task RunAsync()
    {
        PlanningSnapshot snapshot = await m_DataProvider.LoadAsync(includePlannedHours: false);
        Simulator simulator = snapshot.CreateSimulator();
        simulator.Run();
        simulator.ExportExcel(m_AppConfiguration.FileConfiguration.OutputFilePath);
        simulator.ExportGanttSprintBased(m_AppConfiguration.FileConfiguration.OutputPngFilePath);
    }
}

public class CheckingRunner
{
    private readonly PlanningDataProvider m_DataProvider;
    private readonly AppConfiguration m_AppConfiguration;

    public CheckingRunner(PlanningDataProvider dataProvider, AppConfiguration appConfiguration)
    {
        m_DataProvider = dataProvider;
        m_AppConfiguration = appConfiguration;
    }

    public async Task RunAsync()
    {
        PlanningSnapshot snapshot = await m_DataProvider.LoadAsync(includePlannedHours: true);
        Simulator simulator = snapshot.CreateSimulator();
        simulator.Run();

        string basePath = m_AppConfiguration.FileConfiguration.OutputFilePath;
        string directory = Path.GetDirectoryName(basePath) ?? string.Empty;
        string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(basePath);
        string extension = Path.GetExtension(basePath);
        if (string.IsNullOrWhiteSpace(extension))
        {
            extension = ".xlsx";
        }

        string comparisonPath = Path.Combine(directory, $"{fileNameWithoutExtension}_Comparison{extension}");
        simulator.ExportComparisonReport(comparisonPath);
    }
}

public class PlanningSnapshot
{
    private readonly List<Epic> m_Epics;
    private readonly Dictionary<int, Dictionary<string, ResourceCapacity>> m_SprintCapacities;
    private readonly DateTime m_InitialSprintStart;
    private readonly int m_SprintDays;
    private readonly int m_MaxSprintCount;
    private readonly int m_SprintOffset;
    private readonly Dictionary<string, double> m_PlannedHours;

    public PlanningSnapshot(
        List<Epic> epics,
        Dictionary<int, Dictionary<string, ResourceCapacity>> sprintCapacities,
        DateTime initialSprintStart,
        int sprintDays,
        int maxSprintCount,
        int sprintOffset,
        Dictionary<string, double> plannedHours)
    {
        m_Epics = epics;
        m_SprintCapacities = sprintCapacities;
        m_InitialSprintStart = initialSprintStart;
        m_SprintDays = sprintDays;
        m_MaxSprintCount = maxSprintCount;
        m_SprintOffset = sprintOffset;
        m_PlannedHours = plannedHours;
    }

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
}
