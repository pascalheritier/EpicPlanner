using EpicPlanner.Core.Configuration;
using OfficeOpenXml;
using System.Text.RegularExpressions;

namespace EpicPlanner.Core.Checker.Reports;

/// <summary>
/// Reads current epic state from Planification_des_Epics.xlsx and historical sprint
/// data (FinalSchedule + AllocationsByEpicPerSprint) by auto-discovering sprint files
/// within <see cref="EpicAnalysisReportConfiguration.InputFolderPath"/>.
/// </summary>
public class EpicAnalysisDataLoader
{
    #region Inner types

    private sealed record CurrentEpicRow(
        string Id,
        string ShortName,
        string State,
        string Manager,
        string Analyst,
        string AssignedTo,
        double? OriginalEstimate,
        double? RemainingEstimate,
        double? RoughEstimate,
        string Dependencies,
        string EndAnalysisRaw);

    private sealed record SprintFileData(
        int SprintNumber,
        DateTime SprintStart,
        IReadOnlyDictionary<string, double> InitialRemaining,
        IReadOnlyDictionary<string, double> Allocated);

    #endregion

    #region Members

    private readonly AppConfiguration m_Config;
    private static readonly StringComparer s_IdComparer = StringComparer.OrdinalIgnoreCase;

    // Matches directory names like "Sprint 76", "Sprint 83"
    private static readonly Regex s_SprintDirRegex =
        new(@"^Sprint\s+(\d+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    // Matches sprint planning xlsx (not Planification, not retrospective, not comparison/Comparison)
    private static readonly Regex s_SprintPlanFileRegex =
        new(@"Sprint[_\s]?\w*\d+", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    #endregion

    #region Constructor

    public EpicAnalysisDataLoader(AppConfiguration config)
    {
        m_Config = config;
    }

    #endregion

    #region Public API

    public Task<EpicAnalysisReportModel> LoadAsync()
    {
        EpicAnalysisReportConfiguration reportConfig = m_Config.EpicAnalysisReportConfiguration;

        List<SprintFileData> history = DiscoverSprintFiles(reportConfig.InputFolderPath);

        string epicFilePath = ResolveCurrentEpicFile(reportConfig.InputFolderPath, m_Config.FileConfiguration.InputFilePath);
        List<CurrentEpicRow> currentEpics = ReadCurrentEpics(epicFilePath);

        int n = history.Count;
        List<string> sprintLabels = history.Select(h => $"S{h.SprintNumber}").ToList();
        List<string> sprintDates  = history.Select(h => FormatSprintDate(h.SprintStart)).ToList();

        HashSet<string> historicalIds = new(
            history.SelectMany(h => h.InitialRemaining.Keys.Concat(h.Allocated.Keys)),
            s_IdComparer);

        List<EpicAnalysisEntry> mainEntries     = new();
        List<PipelineEpicEntry> pipelineEntries = new();

        foreach (CurrentEpicRow epic in currentEpics)
        {
            string stateNorm = (epic.State ?? string.Empty).ToLowerInvariant();

            bool isDone          = stateNorm.Contains("done") || stateNorm.Contains("abandon");
            bool isInDev         = stateNorm.Contains("develop") && !stateNorm.Contains("pending");
            bool isPendingDev    = stateNorm.Contains("pending") && stateNorm.Contains("develop");
            bool isInAnalysis    = stateNorm.Contains("analysis") || stateNorm.Contains("analyse");
            bool hadHistory      = historicalIds.Contains(epic.Id);

            bool goesToMain     = isDone || isInDev || isPendingDev || (hadHistory && !isInAnalysis);
            bool goesToPipeline = !goesToMain && (isInAnalysis || (isPendingDev && !hadHistory));

            if (goesToMain)
                mainEntries.Add(BuildMainEntry(epic, history, sprintLabels, n, isDone));
            else if (goesToPipeline)
                pipelineEntries.Add(BuildPipelineEntry(epic));
        }

        mainEntries.Sort((a, b) => RiskOrder(a.Risk).CompareTo(RiskOrder(b.Risk)));

        string lastSprint = sprintLabels.LastOrDefault() ?? string.Empty;

        return Task.FromResult(new EpicAnalysisReportModel
        {
            Epics = mainEntries,
            Pipeline = pipelineEntries,
            SprintLabels = sprintLabels,
            SprintDates = sprintDates,
            GeneratedAt = DateTime.Now,
            CurrentSprintLabel = lastSprint,
            CurrentSprintDateRange = history.LastOrDefault() is { } last
                ? last.SprintStart.ToString("dd MMM yyyy") : string.Empty
        });
    }

    #endregion

    #region Sprint file discovery

    /// <summary>
    /// Scans <paramref name="rootFolder"/> for "Sprint XX" subdirectories, picks the latest
    /// version file in each, and returns them sorted by sprint number.
    /// </summary>
    private static List<SprintFileData> DiscoverSprintFiles(string rootFolder)
    {
        if (string.IsNullOrWhiteSpace(rootFolder) || !Directory.Exists(rootFolder))
            return new List<SprintFileData>();

        var result = new List<SprintFileData>();

        foreach (string sprintDir in Directory.GetDirectories(rootFolder))
        {
            string dirName = Path.GetFileName(sprintDir);
            Match m = s_SprintDirRegex.Match(dirName);
            if (!m.Success) continue;

            int sprintNumber = int.Parse(m.Groups[1].Value);
            string? planFile = FindLatestSprintPlanFile(sprintDir);
            if (planFile is null) continue;

            SprintFileData? data = TryReadSprintFile(sprintNumber, planFile);
            if (data is not null)
                result.Add(data);
        }

        return result.OrderBy(d => d.SprintNumber).ToList();
    }

    /// <summary>
    /// Inside a sprint directory, finds the latest versioned sprint planning Excel file.
    /// Prefers vN subdirectories (highest N), then falls back to root.
    /// Excludes Planification_des_Epics, retrospective and comparison files.
    /// </summary>
    private static string? FindLatestSprintPlanFile(string sprintDir)
    {
        // Collect (version, filePath) candidates
        var candidates = new List<(int Version, string Path)>();

        // Check versioned subdirectories (v1, v2, v3 …)
        foreach (string subDir in Directory.GetDirectories(sprintDir))
        {
            string sub = Path.GetFileName(subDir);
            if (!Regex.IsMatch(sub, @"^v\d+$", RegexOptions.IgnoreCase)) continue;
            int ver = int.Parse(sub[1..]);

            string? file = PickSprintPlanFileInDirectory(subDir);
            if (file is not null)
                candidates.Add((ver, file));
        }

        if (candidates.Count > 0)
            return candidates.MaxBy(c => c.Version).Path;

        // Fallback: root of sprint directory
        return PickSprintPlanFileInDirectory(sprintDir);
    }

    private static string? PickSprintPlanFileInDirectory(string dir)
    {
        return Directory
            .GetFiles(dir, "*.xlsx")
            .FirstOrDefault(IsSprintPlanFile);
    }

    private static bool IsSprintPlanFile(string filePath)
    {
        string name = Path.GetFileNameWithoutExtension(filePath);
        if (name.Contains("Planification", StringComparison.OrdinalIgnoreCase)) return false;
        if (name.Contains("retrospective", StringComparison.OrdinalIgnoreCase)) return false;
        if (name.Contains("comparison", StringComparison.OrdinalIgnoreCase)) return false;
        if (name.Contains("Comparison", StringComparison.OrdinalIgnoreCase)) return false;
        return s_SprintPlanFileRegex.IsMatch(name);
    }

    #endregion

    #region Read sprint file

    private static SprintFileData? TryReadSprintFile(int sprintNumber, string filePath)
    {
        try
        {
            using var package = new ExcelPackage(new FileInfo(filePath));

            var initialRemaining = new Dictionary<string, double>(s_IdComparer);
            var allocated        = new Dictionary<string, double>(s_IdComparer);
            DateTime sprintStart = DateTime.MinValue;

            ReadFinalSchedule(package, initialRemaining);
            ReadAllocationsByEpicPerSprint(package, sprintNumber, allocated, ref sprintStart);

            // If we couldn't get the date from allocations, try to infer it
            if (sprintStart == DateTime.MinValue)
                sprintStart = InferSprintStartFromFinalSchedule(package);

            return new SprintFileData(sprintNumber, sprintStart, initialRemaining, allocated);
        }
        catch
        {
            return null;
        }
    }

    private static void ReadFinalSchedule(ExcelPackage package, Dictionary<string, double> target)
    {
        ExcelWorksheet? ws = package.Workbook.Worksheets["FinalSchedule"];
        if (ws?.Dimension is null) return;

        int colEpic = 1, colCharge = 3;
        int maxCol = ws.Dimension.End.Column;
        for (int c = 1; c <= maxCol; c++)
        {
            string? h = ws.Cells[1, c].GetValue<string>()?.Trim();
            if (h is null) continue;
            if (h.StartsWith("Epic", StringComparison.OrdinalIgnoreCase)) colEpic = c;
            else if (h.Equals("Initial_Charge_h", StringComparison.OrdinalIgnoreCase)) colCharge = c;
        }

        int maxRow = ws.Dimension.End.Row;
        for (int r = 2; r <= maxRow; r++)
        {
            string? epicName = ws.Cells[r, colEpic].GetValue<string>()?.Trim();
            if (string.IsNullOrWhiteSpace(epicName)) continue;

            double? charge = ws.Cells[r, colCharge].GetValue<double?>()
                          ?? TryParseDouble(ws.Cells[r, colCharge].Text);
            if (charge is null) continue;

            string id = ExtractEpicId(epicName);
            if (!string.IsNullOrWhiteSpace(id))
                target[id] = charge.Value;
        }
    }

    private static void ReadAllocationsByEpicPerSprint(
        ExcelPackage package,
        int targetSprintNumber,
        Dictionary<string, double> target,
        ref DateTime sprintStart)
    {
        ExcelWorksheet? ws = package.Workbook.Worksheets["AllocationsByEpicPerSprint"];
        if (ws?.Dimension is null) return;

        int colEpic = 1, colSprint = 2, colStart = 3, colHours = 4;
        int maxCol = ws.Dimension.End.Column;
        for (int c = 1; c <= maxCol; c++)
        {
            string? h = ws.Cells[1, c].GetValue<string>()?.Trim();
            if (h is null) continue;
            if (h.StartsWith("Epic", StringComparison.OrdinalIgnoreCase)) colEpic = c;
            else if (h.Equals("Sprint", StringComparison.OrdinalIgnoreCase)) colSprint = c;
            else if (h.StartsWith("Sprint_start", StringComparison.OrdinalIgnoreCase) ||
                     h.StartsWith("Sprint start", StringComparison.OrdinalIgnoreCase)) colStart = c;
            else if (h.StartsWith("Total", StringComparison.OrdinalIgnoreCase)) colHours = c;
        }

        int maxRow = ws.Dimension.End.Row;
        for (int r = 2; r <= maxRow; r++)
        {
            int? sprint = ws.Cells[r, colSprint].GetValue<int?>()
                       ?? (int?)TryParseDouble(ws.Cells[r, colSprint].Text);
            if (sprint != targetSprintNumber) continue;

            // Capture sprint start date from first matching row
            if (sprintStart == DateTime.MinValue)
            {
                DateTime? d = ws.Cells[r, colStart].GetValue<DateTime?>();
                if (d.HasValue) sprintStart = d.Value;
                else if (DateTime.TryParse(ws.Cells[r, colStart].Text, out DateTime parsed))
                    sprintStart = parsed;
            }

            string? epicName = ws.Cells[r, colEpic].GetValue<string>()?.Trim();
            if (string.IsNullOrWhiteSpace(epicName)) continue;

            double? hours = ws.Cells[r, colHours].GetValue<double?>()
                         ?? TryParseDouble(ws.Cells[r, colHours].Text);
            if (hours is null) continue;

            string id = ExtractEpicId(epicName);
            if (!string.IsNullOrWhiteSpace(id))
                target[id] = target.TryGetValue(id, out double existing) ? existing + hours.Value : hours.Value;
        }
    }

    private static DateTime InferSprintStartFromFinalSchedule(ExcelPackage package)
    {
        ExcelWorksheet? ws = package.Workbook.Worksheets["FinalSchedule"];
        if (ws?.Dimension is null) return DateTime.MinValue;

        int colStart = -1, maxCol = ws.Dimension.End.Column;
        for (int c = 1; c <= maxCol; c++)
        {
            string? h = ws.Cells[1, c].GetValue<string>()?.Trim();
            if (h is null) continue;
            if (h.StartsWith("Start", StringComparison.OrdinalIgnoreCase)) { colStart = c; break; }
        }
        if (colStart < 0) return DateTime.MinValue;

        for (int r = 2; r <= ws.Dimension.End.Row; r++)
        {
            DateTime? d = ws.Cells[r, colStart].GetValue<DateTime?>();
            if (d.HasValue && d.Value > DateTime.MinValue) return d.Value;
        }
        return DateTime.MinValue;
    }

    #endregion

    #region Read current epics

    /// <summary>
    /// Resolves the path to the current Planification_des_Epics.xlsx.
    /// Prefers the most recent one found in the sprint folder, falls back to FileConfiguration.InputFilePath.
    /// </summary>
    private static string ResolveCurrentEpicFile(string inputFolder, string fallbackPath)
    {
        if (!string.IsNullOrWhiteSpace(inputFolder) && Directory.Exists(inputFolder))
        {
            // Find the latest "Planification_des_Epics.xlsx" in any Sprint XX / vN subdirectory
            var candidates = Directory
                .GetFiles(inputFolder, "Planification_des_Epics.xlsx", SearchOption.AllDirectories)
                .Select(f => (File: f, SprintNum: ExtractSprintNumFromPath(f)))
                .Where(c => c.SprintNum > 0)
                .OrderByDescending(c => c.SprintNum)
                .ThenByDescending(c => ExtractVersionFromPath(c.File))
                .ToList();

            if (candidates.Count > 0) return candidates[0].File;
        }

        return fallbackPath;
    }

    private static int ExtractSprintNumFromPath(string path)
    {
        Match m = Regex.Match(path, @"[/\\]Sprint\s+(\d+)[/\\]", RegexOptions.IgnoreCase);
        return m.Success ? int.Parse(m.Groups[1].Value) : 0;
    }

    private static int ExtractVersionFromPath(string path)
    {
        Match m = Regex.Match(path, @"[/\\]v(\d+)[/\\]", RegexOptions.IgnoreCase);
        return m.Success ? int.Parse(m.Groups[1].Value) : 0;
    }

    private static List<CurrentEpicRow> ReadCurrentEpics(string filePath)
    {
        var result = new List<CurrentEpicRow>();
        using var package = new ExcelPackage(new FileInfo(filePath));

        ExcelWorksheet? ws = package.Workbook.Worksheets["Planification des Epics"];
        if (ws?.Dimension is null) return result;

        var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        int maxCol = ws.Dimension.End.Column;
        for (int c = 1; c <= maxCol; c++)
        {
            string? h = ws.Cells[1, c].GetValue<string>()?.Trim();
            if (!string.IsNullOrWhiteSpace(h)) headers[h] = c;
        }

        int colName     = GetCol(headers, "Epic name",            1);
        int colManager  = GetCol(headers, "Epic Manager",         2);
        int colAnalyst  = GetCol(headers, "Epic Analyst",         3);
        int colState    = GetCol(headers, "State",                4);
        int colEndAna   = GetCol(headers, "End of analysis",      5);
        int colOrigEst  = GetCol(headers, "Original estimate [h]",6);
        int colRemain   = GetCol(headers, "Remaining  [h]",       7);
        int colAssigned = GetCol(headers, "Assigned to",          8);
        int colRough    = GetCol(headers, "Rough estimate  [h]",  9);
        int colDeps     = GetCol(headers, "Epic dependency",      12);

        int maxRow = ws.Dimension.End.Row;
        for (int r = 2; r <= maxRow; r++)
        {
            string? fullName = ws.Cells[r, colName].GetValue<string>()?.Trim();
            if (string.IsNullOrWhiteSpace(fullName)) continue;

            result.Add(new CurrentEpicRow(
                Id:               ExtractEpicId(fullName),
                ShortName:        ExtractShortName(fullName),
                State:            ws.Cells[r, colState].GetValue<string>()?.Trim() ?? string.Empty,
                Manager:          ws.Cells[r, colManager].GetValue<string>()?.Trim() ?? string.Empty,
                Analyst:          ws.Cells[r, colAnalyst].GetValue<string>()?.Trim() ?? string.Empty,
                AssignedTo:       ws.Cells[r, colAssigned].GetValue<string>()?.Trim() ?? string.Empty,
                OriginalEstimate: ws.Cells[r, colOrigEst].GetValue<double?>()  ?? TryParseDouble(ws.Cells[r, colOrigEst].Text),
                RemainingEstimate:ws.Cells[r, colRemain].GetValue<double?>()   ?? TryParseDouble(ws.Cells[r, colRemain].Text),
                RoughEstimate:    ws.Cells[r, colRough].GetValue<double?>()    ?? TryParseDouble(ws.Cells[r, colRough].Text),
                Dependencies:     ws.Cells[r, colDeps].GetValue<string>()?.Trim() ?? string.Empty,
                EndAnalysisRaw:   ws.Cells[r, colEndAna].Text?.Trim() ?? string.Empty));
        }

        return result;
    }

    #endregion

    #region Build report entries

    private static EpicAnalysisEntry BuildMainEntry(
        CurrentEpicRow epic,
        List<SprintFileData> history,
        List<string> sprintLabels,
        int n,
        bool isDone)
    {
        double[] allocation = new double[n];
        double?[] remaining = new double?[n + 1];

        for (int i = 0; i < n; i++)
        {
            SprintFileData sprint = history[i];
            allocation[i] = sprint.Allocated.TryGetValue(epic.Id, out double alloc)
                ? Math.Round(alloc, 2) : 0.0;
            remaining[i]  = sprint.InitialRemaining.TryGetValue(epic.Id, out double rem)
                ? rem : null;
        }
        remaining[n] = epic.RemainingEstimate ?? 0.0;

        double? orig = epic.OriginalEstimate > 0 ? epic.OriginalEstimate : null;
        double  cur  = epic.RemainingEstimate ?? 0.0;

        (string risk, string riskSince, string riskDesc) = isDone
            ? ("done", "—", string.Empty)
            : EpicRiskAssessor.Assess(epic.Id, DetermineStateJs(epic.State), orig, cur, allocation, remaining, sprintLabels);

        return new EpicAnalysisEntry
        {
            Id               = epic.Id,
            Name             = epic.ShortName,
            Manager          = epic.Manager,
            Assigned         = epic.AssignedTo,
            State            = DetermineStateJs(epic.State),
            Risk             = risk,
            OriginalEstimate = orig,
            CurrentRemaining = cur,
            RiskSince        = riskSince,
            StateLabel       = epic.State,
            RiskDesc         = riskDesc,
            Allocation       = allocation,
            Remaining        = remaining
        };
    }

    private static PipelineEpicEntry BuildPipelineEntry(CurrentEpicRow epic)
    {
        string notes = string.Empty;
        if (!string.IsNullOrWhiteSpace(epic.EndAnalysisRaw) && epic.EndAnalysisRaw != "?")
        {
            notes = $"Fin analyse : {epic.EndAnalysisRaw}.";
            if (DateTime.TryParse(epic.EndAnalysisRaw, out DateTime end) && end < DateTime.Today)
                notes += " ⚠️ Date dépassée.";
        }
        if (!string.IsNullOrWhiteSpace(epic.Dependencies))
            notes += (notes.Length > 0 ? " " : "") + $"Dépend de : {epic.Dependencies}.";

        return new PipelineEpicEntry
        {
            Id           = epic.Id,
            Name         = epic.ShortName,
            Manager      = epic.Manager,
            Analyst      = epic.Analyst,
            State        = epic.State,
            RoughEstimate = epic.RoughEstimate > 0 ? epic.RoughEstimate : null,
            Dependencies = string.IsNullOrWhiteSpace(epic.Dependencies) ? "—" : epic.Dependencies,
            Notes        = notes.Trim()
        };
    }

    #endregion

    #region Helpers

    private static string ExtractEpicId(string fullName)
    {
        if (string.IsNullOrWhiteSpace(fullName)) return string.Empty;
        int colonIdx = fullName.IndexOf(':', StringComparison.Ordinal);
        string raw = colonIdx > 0 ? fullName[..colonIdx].Trim() : fullName.Split(' ')[0].Trim();
        return raw.Trim('\u00a0', ' ', '\u200b');
    }

    private static string ExtractShortName(string fullName)
    {
        int colonIdx = fullName.IndexOf(':', StringComparison.Ordinal);
        return colonIdx > 0 ? fullName[(colonIdx + 1)..].Trim() : fullName;
    }

    private static string DetermineStateJs(string? rawState)
    {
        string s = (rawState ?? string.Empty).ToLowerInvariant();
        if (s.Contains("done") || s.Contains("abandon")) return "done";
        if (s.Contains("develop") && !s.Contains("pending")) return "in_dev";
        return "pending";
    }

    private static string FormatSprintDate(DateTime d) =>
        d == DateTime.MinValue ? "?" :
        d.ToString("MMM", System.Globalization.CultureInfo.GetCultureInfo("fr-FR")) + "'" + d.ToString("yy");

    private static double? TryParseDouble(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        string n = text.Replace(',', '.').Trim();
        return double.TryParse(n, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out double v) ? v : null;
    }

    private static int GetCol(Dictionary<string, int> headers, string key, int fallback) =>
        headers.TryGetValue(key, out int col) ? col : fallback;

    private static int RiskOrder(string r) => r switch
    {
        "critical" => 0,
        "watch"    => 1,
        "ok"       => 2,
        _          => 3
    };

    #endregion
}
