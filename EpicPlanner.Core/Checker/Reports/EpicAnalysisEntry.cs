namespace EpicPlanner.Core.Checker.Reports;

/// <summary>
/// Represents a single epic with its historical data for the analysis report.
/// </summary>
public class EpicAnalysisEntry
{
    public string   Id               { get; set; } = string.Empty;
    public string   Name             { get; set; } = string.Empty;
    public string   Manager          { get; set; } = string.Empty;
    public string   Assigned         { get; set; } = string.Empty;

    /// <summary>"in_dev" | "pending" | "done"</summary>
    public string   State            { get; set; } = string.Empty;

    /// <summary>"critical" | "watch" | "ok" | "done"</summary>
    public string   Risk             { get; set; } = string.Empty;

    public double?  OriginalEstimate { get; set; }
    public double   CurrentRemaining { get; set; }
    public string   RiskSince        { get; set; } = string.Empty;
    public string   StateLabel       { get; set; } = string.Empty;
    public string   RiskDesc         { get; set; } = string.Empty;

    /// <summary>Hours allocated per sprint (one value per historical sprint).</summary>
    public double[] Allocation       { get; set; } = Array.Empty<double>();

    /// <summary>
    /// Remaining hours at the START of each sprint (null = not in plan that sprint),
    /// plus one extra entry at the end = current remaining.
    /// </summary>
    public double?[] Remaining       { get; set; } = Array.Empty<double?>();
}

/// <summary>
/// Represents an epic in the analysis pipeline (not yet in development).
/// </summary>
public class PipelineEpicEntry
{
    public string  Id            { get; set; } = string.Empty;
    public string  Name          { get; set; } = string.Empty;
    public string  Manager       { get; set; } = string.Empty;
    public string  Analyst       { get; set; } = string.Empty;
    public string  State         { get; set; } = string.Empty;
    public double? RoughEstimate { get; set; }
    public string  Dependencies  { get; set; } = string.Empty;
    public string  Notes         { get; set; } = string.Empty;
}

/// <summary>
/// Root model returned by <see cref="EpicAnalysisDataLoader"/> and consumed by
/// <see cref="EpicAnalysisHtmlGenerator"/>.
/// </summary>
public class EpicAnalysisReportModel
{
    public List<EpicAnalysisEntry> Epics    { get; set; } = new();
    public List<PipelineEpicEntry> Pipeline { get; set; } = new();

    /// <summary>One label per historical sprint, e.g. ["S76","S77",…]</summary>
    public List<string> SprintLabels        { get; set; } = new();

    /// <summary>One human-readable date per sprint, e.g. ["oct'25","nov'25",…]</summary>
    public List<string> SprintDates         { get; set; } = new();

    public DateTime GeneratedAt             { get; set; }
    public string   CurrentSprintLabel      { get; set; } = string.Empty;
    public string   CurrentSprintDateRange  { get; set; } = string.Empty;
}
