using EpicPlanner.Core.Checker.Services;
using EpicPlanner.Core.Configuration;
using EpicPlanner.Core.Planner.Services;
using EpicPlanner.Core.Shared.Models;
using OfficeOpenXml;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace EpicPlanner.Core.Shared.Services;

public class PlanningDataProvider
{
    #region Members

    private readonly AppConfiguration m_AppConfiguration;
    private readonly DateTime m_InitialSprintStart;
    private readonly int m_iSprintDays;
    private readonly int m_iSprintCapacityDays;

    #endregion

    #region Constructor

    public PlanningDataProvider(AppConfiguration _AppConfiguration)
    {
        m_AppConfiguration = _AppConfiguration ?? throw new ArgumentNullException(nameof(_AppConfiguration));
        m_InitialSprintStart = _AppConfiguration.PlannerConfiguration.InitialSprintStartDate;
        m_iSprintDays = _AppConfiguration.PlannerConfiguration.SprintDays;
        m_iSprintCapacityDays = _AppConfiguration.PlannerConfiguration.SprintCapacityDays;
    }

    #endregion

    #region Load data

    public async Task<PlannerPlanningSnapshot> LoadPlannerSnapshotAsync(bool _bIncludePlannedHours)
    {
        PlanningSnapshotComponents components = await LoadComponentsAsync(_bIncludePlannedHours);

        return new PlannerPlanningSnapshot(
            components.Epics,
            components.AdjustedCapacities,
            m_InitialSprintStart,
            m_iSprintDays,
            m_AppConfiguration.PlannerConfiguration.MaxSprintCount,
            m_AppConfiguration.PlannerConfiguration.InitialSprintNumber);
    }

    public async Task<CheckerPlanningSnapshot> LoadCheckerSnapshotAsync()
    {
        PlanningSnapshotComponents components = await LoadComponentsAsync(_bIncludePlannedHours: true);

        return new CheckerPlanningSnapshot(
            components.Epics,
            components.AdjustedCapacities,
            m_InitialSprintStart,
            m_iSprintDays,
            m_AppConfiguration.PlannerConfiguration.MaxSprintCount,
            m_AppConfiguration.PlannerConfiguration.InitialSprintNumber,
            components.PlannedHours,
            components.EpicSummaries,
            components.PlannedCapacityByEpic);
    }

    private async Task<PlanningSnapshotComponents> LoadComponentsAsync(bool _bIncludePlannedHours)
    {
        using ExcelPackage package = new(new FileInfo(m_AppConfiguration.FileConfiguration.InputFilePath));
        ExcelWorksheet wsEpics = package.Workbook.Worksheets[m_AppConfiguration.FileConfiguration.InputEpicsSheetName]
            ?? throw new NullReferenceException($"Worksheet '{m_AppConfiguration.FileConfiguration.InputEpicsSheetName}' could not be found.");
        ExcelWorksheet wsRes = package.Workbook.Worksheets[m_AppConfiguration.FileConfiguration.InputResourcesSheetName]
            ?? throw new NullReferenceException($"Worksheet '{m_AppConfiguration.FileConfiguration.InputResourcesSheetName}' could not be found.");

        Dictionary<string, double> plannedCapacityByEpic = LoadPlannedCapacityLookup(
            package,
            m_AppConfiguration.FileConfiguration.InputFilePath,
            m_AppConfiguration.FileConfiguration.PlannedCapacityFilePath,
            m_AppConfiguration.PlannerConfiguration.InitialSprintNumber);

        Dictionary<string, ResourceCapacity> resources = LoadResources(wsRes);
        List<Epic> epics = LoadEpics(wsEpics, resources.Keys.ToList());

        RedmineDataFetcher redmineDataFetcher = new(
            m_AppConfiguration.RedmineConfiguration.ServerUrl,
            m_AppConfiguration.RedmineConfiguration.ApiKey);

        // Adjust resources for absences (for each sprint)
        Dictionary<string, List<(DateTime, DateTime)>> absencesPerResource = await redmineDataFetcher.GetResourcesAbsencesAsync();
        Dictionary<int, Dictionary<string, ResourceCapacity>> adjustedCapacities = AdjustCapacitiesForAbsences(
            resources,
            absencesPerResource,
            m_AppConfiguration.PlannerConfiguration.Holidays);

        Dictionary<string, double> plannedHours = new(StringComparer.OrdinalIgnoreCase);
        List<SprintEpicSummary> epicSummaries = new();
        HashSet<string> plannedEpicNames = epics
            .Where(e => e.IsInDevelopment)
            .Select(e => e.Name)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (_bIncludePlannedHours)
        {
            // Get currently planned hours from Redmine
            plannedHours = await redmineDataFetcher.GetPlannedHoursForSprintAsync(
                m_AppConfiguration.PlannerConfiguration.InitialSprintNumber,
                m_InitialSprintStart,
                m_InitialSprintStart.AddDays(m_iSprintDays - 1),
                plannedEpicNames);

            epicSummaries = await redmineDataFetcher.GetEpicSprintSummariesAsync(
                m_AppConfiguration.PlannerConfiguration.InitialSprintNumber,
                m_InitialSprintStart,
                m_InitialSprintStart.AddDays(m_iSprintDays - 1),
                plannedEpicNames,
                plannedCapacityByEpic.Count > 0 ? plannedCapacityByEpic : null);
        }

        return new PlanningSnapshotComponents(
            epics,
            adjustedCapacities,
            plannedHours,
            epicSummaries,
            plannedCapacityByEpic);
    }

    private sealed record PlanningSnapshotComponents(
        List<Epic> Epics,
        Dictionary<int, Dictionary<string, ResourceCapacity>> AdjustedCapacities,
        Dictionary<string, double> PlannedHours,
        IReadOnlyList<SprintEpicSummary> EpicSummaries,
        Dictionary<string, double> PlannedCapacityByEpic);

    private static Dictionary<string, double> LoadPlannedCapacityLookup(
        ExcelPackage _PrimaryPackage,
        string _strPrimaryFilePath,
        string? _strOverrideFilePath,
        int _iInitialSprintNumber)
    {
        ExcelWorksheet? fallbackWorksheet = _PrimaryPackage.Workbook.Worksheets["AllocationsByEpicPerSprint"];

        if (string.IsNullOrWhiteSpace(_strOverrideFilePath))
        {
            return ReadPlannedCapacityByEpic(fallbackWorksheet, _iInitialSprintNumber);
        }

        try
        {
            string resolvedOverridePath = Path.GetFullPath(_strOverrideFilePath);
            string resolvedPrimaryPath = Path.GetFullPath(_strPrimaryFilePath);

            if (string.Equals(resolvedOverridePath, resolvedPrimaryPath, StringComparison.OrdinalIgnoreCase))
            {
                return ReadPlannedCapacityByEpic(fallbackWorksheet, _iInitialSprintNumber);
            }

            if (!File.Exists(resolvedOverridePath))
            {
                Console.WriteLine($"Warning: Planned capacity file '{resolvedOverridePath}' was not found. Falling back to the primary workbook.");
                return ReadPlannedCapacityByEpic(fallbackWorksheet, _iInitialSprintNumber);
            }

            using ExcelPackage overridePackage = new(new FileInfo(resolvedOverridePath));
            ExcelWorksheet? overrideWorksheet = overridePackage.Workbook.Worksheets["AllocationsByEpicPerSprint"];

            if (overrideWorksheet == null || overrideWorksheet.Dimension == null)
            {
                Console.WriteLine($"Warning: Worksheet 'AllocationsByEpicPerSprint' was not found in '{resolvedOverridePath}'.");
                return ReadPlannedCapacityByEpic(fallbackWorksheet, _iInitialSprintNumber);
            }

            return ReadPlannedCapacityByEpic(overrideWorksheet, _iInitialSprintNumber);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Warning: Failed to load planned capacity file '{_strOverrideFilePath}': {ex.Message}");
            return ReadPlannedCapacityByEpic(fallbackWorksheet, _iInitialSprintNumber);
        }
    }

    private static Dictionary<string, double> ReadPlannedCapacityByEpic(ExcelWorksheet? _Worksheet, int _iInitialSprintNumber)
    {
        var result = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        if (_Worksheet?.Dimension == null)
            return result;

        var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int col = 1; col <= _Worksheet.Dimension.End.Column; col++)
        {
            string? header = _Worksheet.Cells[1, col].GetValue<string>()?.Trim();
            if (!string.IsNullOrWhiteSpace(header))
                headers[header] = col;
        }

        if (!headers.TryGetValue("Epic", out int epicCol) ||
            !headers.TryGetValue("Sprint", out int sprintCol))
        {
            return result;
        }

        if (!headers.TryGetValue("Total_Hours", out int totalHoursCol))
        {
            // Accept also Total Hours without underscore for robustness.
            headers.TryGetValue("Total Hours", out totalHoursCol);
        }

        if (totalHoursCol <= 0)
            return result;

        for (int row = 2; row <= _Worksheet.Dimension.End.Row; row++)
        {
            string? epicName = _Worksheet.Cells[row, epicCol].GetValue<string>()?.Trim();
            if (string.IsNullOrWhiteSpace(epicName))
                continue;

            if (!TryReadSprintNumber(_Worksheet.Cells[row, sprintCol].Value, out int sprintNumber) ||
                sprintNumber != _iInitialSprintNumber)
            {
                continue;
            }

            double totalHours = ReadNumericValue(_Worksheet.Cells[row, totalHoursCol].Value);
            if (totalHours < 0)
                totalHours = 0;

            if (result.TryGetValue(epicName, out double existing))
                result[epicName] = existing + totalHours;
            else
                result[epicName] = totalHours;
        }

        return result;

        static bool TryReadSprintNumber(object? _Value, out int _iSprintNumber)
        {
            switch (_Value)
            {
                case int i:
                    _iSprintNumber = i;
                    return true;
                case long l:
                    _iSprintNumber = (int)l;
                    return true;
                case double d:
                    _iSprintNumber = (int)Math.Round(d);
                    return true;
                case decimal m:
                    _iSprintNumber = (int)Math.Round((double)m);
                    return true;
                case string s when int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed):
                    _iSprintNumber = parsed;
                    return true;
                default:
                    _iSprintNumber = 0;
                    return false;
            }
        }

        static double ReadNumericValue(object? _Value)
        {
            return _Value switch
            {
                null => 0.0,
                double d => d,
                int i => i,
                long l => l,
                decimal m => (double)m,
                float f => f,
                string s when double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsed) => parsed,
                _ => 0.0
            };
        }
    }

    private Dictionary<string, ResourceCapacity> LoadResources(ExcelWorksheet _ResourceWorksheet)
    {
        int rows = _ResourceWorksheet.Dimension.End.Row;
        var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        for (int c = 1; c <= _ResourceWorksheet.Dimension.End.Column; c++)
        {
            var h = _ResourceWorksheet.Cells[2, c].GetValue<string>()?.Trim();
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
            string name = _ResourceWorksheet.Cells[row, nameCol].GetValue<string>()?.Trim();
            if (string.IsNullOrWhiteSpace(name)) continue;

            double dev = _ResourceWorksheet.Cells[row, devCol].GetValue<double>();
            double maint = maintCol > 0 ? _ResourceWorksheet.Cells[row, maintCol].GetValue<double>() : 0;
            double anal = analCol > 0 ? _ResourceWorksheet.Cells[row, analCol].GetValue<double>() : 0;

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
        Dictionary<string, ResourceCapacity> _BaseCapacities,
        Dictionary<string, List<(DateTime, DateTime)>> _AbsencesPerResource,
        IEnumerable<DateTime> _Holidays)
    {
        Dictionary<int, Dictionary<string, ResourceCapacity>> adjustedCapacities = new();
        for (int sprint = 0; sprint < m_AppConfiguration.PlannerConfiguration.MaxSprintCount; sprint++)
        {
            var sprintStart = m_InitialSprintStart.AddDays(sprint * m_iSprintDays).Date;
            var sprintEnd = sprintStart.AddDays(m_iSprintDays - 1).Date;

            int workingDaysInSprint = BusinessCalendar.CountWorkingDays(sprintStart, sprintEnd, _Holidays);

            var sprintCapacities = new Dictionary<string, ResourceCapacity>(StringComparer.OrdinalIgnoreCase);

            foreach (var user in _BaseCapacities.Keys)
            {
                ResourceCapacity userSprintCapacity = new(_BaseCapacities[user]);
                double scale = (double)workingDaysInSprint / m_iSprintCapacityDays;
                userSprintCapacity.AdapteCapacityToScale(scale);

                if (_AbsencesPerResource.TryGetValue(user, out var absList))
                {
                    foreach (var (start, end) in absList)
                    {
                        int absentWorkingDays = BusinessCalendar.CountWorkingDaysOverlap(start, end, sprintStart, sprintEnd, _Holidays);
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

    private List<Epic> LoadEpics(ExcelWorksheet _EpicWorksheet, List<string> _ResourceNames)
    {
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
        int willAssignCol = headers.FirstOrDefault(kv => kv.Key.Contains("Will be assigned", StringComparison.OrdinalIgnoreCase)
            || kv.Key.Contains("Will be assigne", StringComparison.OrdinalIgnoreCase)).Value;
        int priorityCol = headers.ContainsKey("Priority") ? headers["Priority"] : 11;
        int depCol = headers.FirstOrDefault(kv => kv.Key.Contains("Epic dependency", StringComparison.OrdinalIgnoreCase)
            || kv.Key.Contains("Dependency", StringComparison.OrdinalIgnoreCase)).Value;
        int endAnalysisCol = headers.FirstOrDefault(kv => kv.Key.Contains("End of analysis", StringComparison.OrdinalIgnoreCase)).Value;

        int rows = _EpicWorksheet.Dimension.End.Row;
        var epics = new List<Epic>();

        for (int row = 2; row <= rows; row++)
        {
            string epicName = _EpicWorksheet.Cells[row, epicCol].GetValue<string>()?.Trim();
            if (string.IsNullOrWhiteSpace(epicName)) continue;

            string state = _EpicWorksheet.Cells[row, stateCol].GetValue<string>() ?? string.Empty;
            double charge = 0.0;
            double remVal = remainingCol > 0 ? _EpicWorksheet.Cells[row, remainingCol].GetValue<double>() : 0.0;
            double roughVal = roughCol > 0 ? _EpicWorksheet.Cells[row, roughCol].GetValue<double>() : 0.0;

            if (remVal > 0)
                charge = remVal;
            else if (roughVal > 0)
                charge = roughVal;

            string assigned = assignedCol > 0 ? (_EpicWorksheet.Cells[row, assignedCol].GetValue<string>() ?? string.Empty) : string.Empty;
            string willAssign = willAssignCol > 0 ? (_EpicWorksheet.Cells[row, willAssignCol].GetValue<string>() ?? string.Empty) : string.Empty;
            string depRaw = depCol > 0 ? (_EpicWorksheet.Cells[row, depCol].GetValue<string>() ?? string.Empty) : string.Empty;
            string endAnalysisStr = endAnalysisCol > 0 ? (_EpicWorksheet.Cells[row, endAnalysisCol].GetValue<string>() ?? string.Empty) : string.Empty;
            string priorityStr = priorityCol > 0 ? (_EpicWorksheet.Cells[row, priorityCol].Text?.Trim() ?? "Normal") : "Normal";
            EnumEpicPriority priority = priorityStr.ToLower() switch
            {
                "urgent" => EnumEpicPriority.Urgent,
                "high" => EnumEpicPriority.High,
                _ => EnumEpicPriority.Normal
            };
            DateTime? endAnalysis = null;
            if (DateTime.TryParse(endAnalysisStr, out var parsed))
                endAnalysis = parsed;

            var epic = new Epic(epicName, state, charge, endAnalysis)
            {
                Priority = priority
            };
            epic.ParseAssignments(assigned, willAssign, _ResourceNames);
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

    #endregion
}
