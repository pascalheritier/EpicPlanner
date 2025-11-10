namespace EpicPlanner.Core.Configuration;

public class AppConfiguration
{
    public FileConfiguration FileConfiguration { get; set; } = new();
    public RedmineConfiguration RedmineConfiguration { get; set; } = new();
    public PlannerConfiguration PlannerConfiguration { get; set; } = new();
}

public class FileConfiguration
{
    public string InputFilePath { get; set; } = string.Empty;
    public string? PlannedCapacityFilePath { get; set; } = null;
    public string InputResourcesSheetName { get; set; } = string.Empty;
    public string InputEpicsSheetName { get; set; } = string.Empty;
    public string OutputFilePath { get; set; } = string.Empty;
    public string OutputPngFilePath { get; set; } = string.Empty;
}

public class PlannerConfiguration
{
    public DateTime InitialSprintStartDate { get; set; }
    public int InitialSprintNumber { get; set; }
    public int SprintDays { get; set; }
    public int SprintCapacityDays { get; set; }
    public int MaxSprintCount { get; set; }
    public List<DateTime> Holidays { get; set; } = new();
    public bool IncludeNonInDevelopmentEpicsInRealisationGantt { get; set; } = true;
}

public class RedmineConfiguration
{
    public string ServerUrl { get; set; } = string.Empty;
    public string ApiKey { get; set; } = string.Empty;
    public string TargetUserId { get; set; } = string.Empty;
}
