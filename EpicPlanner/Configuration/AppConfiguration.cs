namespace EpicPlanner;

internal class AppConfiguration
{
    public FileConfiguration FileConfiguration { get; set; }
    public RedmineConfiguration RedmineConfiguration { get; set; }
    public PlannerConfiguration PlannerConfiguration { get; set; }
}

internal class FileConfiguration
{
    public string InputFilePath { get; set; }
    public string InputResourcesSheetName { get; set; }
    public string InputEpicsSheetName { get; set; }
    public string OutputFilePath { get; set; }
    public string OutputPngFilePath { get; set; }
}

internal class PlannerConfiguration
{
    public DateTime InitialSprintStartDate { get; set; }
    public int InitialSprintNumber { get; set; }
    public int SprintDays { get; set; }
    public int SprintCapacityDays { get; set; }
    public int MaxSprintCount { get; set; }
    public List<DateTime> Holidays { get; set; }
}

internal class RedmineConfiguration
{
    public string ServerUrl { get; set; }
    public string ApiKey { get; set; }
    public string TargetUserId { get; set; }
}