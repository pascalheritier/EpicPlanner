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

    /// <summary>
    /// Hours actually consumed in each sprint, derived from the decrease in remaining
    /// (remaining[i] − remaining[i+1], clamped to 0 for re-estimations).
    /// null where the computation is not possible (missing remaining on either side).
    /// For the current sprint, overridden with live Redmine time-entry data when available.
    /// </summary>
    public double?[] Consumed        { get; set; } = Array.Empty<double?>();
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
/// Per-epic consumption data for the current sprint, sourced from Redmine time entries.
/// </summary>
public class EpicConsumptionEntry
{
    public string EpicId   { get; set; } = string.Empty;
    public string EpicName { get; set; } = string.Empty;
    public double Planned  { get; set; }
    public double Consumed { get; set; }
    public double Remaining{ get; set; }

    /// <summary>Consumed + Remaining − Planned. Positive = overhead beyond plan.</summary>
    public double Overhead  => Math.Round(Consumed + Remaining - Planned, 1);

    /// <summary>Consumed / Planned × 100, clamped. 0 if nothing planned.</summary>
    public double UsageRatePct => Planned > 0.01
        ? Math.Round(Consumed / Planned * 100, 1)
        : (Consumed > 0 ? 999 : 0);
}


/// <summary>
/// Per-epic detail for one developer in one sprint.
/// </summary>
public class DeveloperEpicDetail
{
    public string EpicId   { get; set; } = string.Empty;
    public string EpicName { get; set; } = string.Empty;
    /// <summary>Hours charged on this epic (Redmine time entries).</summary>
    public double Consumed { get; set; }
    /// <summary>Hours allocated to this epic (Excel, pro-rated by assignee count).</summary>
    public double Planned  { get; set; }
}

/// <summary>
/// Consumed hours per sprint for one developer, sourced from Redmine time entries.
/// </summary>
public class DeveloperSprintStats
{
    /// <summary>Canonical Redmine user name.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Total hours charged per sprint (one value per historical sprint).</summary>
    public double[] Consumed { get; set; } = Array.Empty<double>();

    /// <summary>Per-sprint epic breakdown. Index matches SprintLabels.</summary>
    public List<DeveloperEpicDetail>[] EpicDetails { get; set; } = Array.Empty<List<DeveloperEpicDetail>>();
}

/// <summary>
/// Root model returned by <see cref="EpicAnalysisDataLoader"/> and consumed by
/// <see cref="EpicAnalysisHtmlGenerator"/>.
/// </summary>
public class EpicAnalysisReportModel
{
    public List<EpicAnalysisEntry> Epics    { get; set; } = new();
    public List<PipelineEpicEntry> Pipeline { get; set; } = new();

    /// <summary>Current-sprint consumption per epic (from Excel). Empty if no sprint files found.</summary>
    public List<EpicConsumptionEntry> EpicConsumptions      { get; set; } = new();

    /// <summary>Consumed hours per developer per sprint (from Redmine time entries). Empty if Redmine unavailable.</summary>
    public List<DeveloperSprintStats> DeveloperStats         { get; set; } = new();

    /// <summary>Label of the sprint for which Redmine consumption data was fetched (e.g. "S84").</summary>
    public string EpicConsumptionSprintLabel { get; set; } = string.Empty;

    /// <summary>One label per historical sprint, e.g. ["S76","S77",…]</summary>
    public List<string> SprintLabels        { get; set; } = new();

    /// <summary>One human-readable date per sprint, e.g. ["oct'25","nov'25",…]</summary>
    public List<string> SprintDates         { get; set; } = new();

    /// <summary>Full start date per sprint in dd/MM/yyyy format.</summary>
    public List<string> SprintStartDatesFull { get; set; } = new();

    /// <summary>Full end date per sprint in dd/MM/yyyy format.</summary>
    public List<string> SprintEndDatesFull   { get; set; } = new();

    public DateTime GeneratedAt             { get; set; }
    public string   CurrentSprintLabel      { get; set; } = string.Empty;
    public string   CurrentSprintDateRange  { get; set; } = string.Empty;
}
