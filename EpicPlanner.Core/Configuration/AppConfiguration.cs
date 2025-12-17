namespace EpicPlanner.Core.Configuration;

public class AppConfiguration
{
    public FileConfiguration FileConfiguration { get; set; } = new();
    public RedmineConfiguration RedmineConfiguration { get; set; } = new();
    public PlannerConfiguration PlannerConfiguration { get; set; } = new();
    public StrategicPlanningConfiguration StrategicPlanningConfiguration { get; set; } = new();
}

public class FileConfiguration
{
    public string InputFilePath { get; set; } = string.Empty;
    public string? PlannedCapacityFilePath { get; set; } = null;
    public string InputResourcesSheetName { get; set; } = string.Empty;
    public string InputEpicsSheetName { get; set; } = string.Empty;
    public string? StrategicOutputFilePath { get; set; }
    public string? StrategicOutputPngFilePath { get; set; }
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
    public bool OnlyDevelopmentEpics { get; set; } = false;
}

public class RedmineConfiguration
{
    public string ServerUrl { get; set; } = string.Empty;
    public string ApiKey { get; set; } = string.Empty;
    public string TargetUserId { get; set; } = string.Empty;
}

public class StrategicPlanningConfiguration
{
    public string EpicsSheetName { get; set; } = "Thèmes";
    public string ResourcesSheetName { get; set; } = "Capacité par Sprint";
    public string EpicNameColumn { get; set; } = "Epic";
    public string VersionColumn { get; set; } = "Version Athena";
    public string TargetVersionName { get; set; } = "Version Athena";
    public string TrueEstimateColumn { get; set; } = "True estimate [h]";
    public string RoughEstimateColumn { get; set; } = "Rough estimate [h]";
    public string EpicCompetenceColumn { get; set; } = "Compétences";
    public string ResourceCompetenceColumn { get; set; } = "Compétence";
    public string OrderColumn { get; set; } = "Ordre";
    public double AbsenceWeeksPerYear { get; set; } = 0;
}
