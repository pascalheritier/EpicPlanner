namespace EpicPlanner;

internal class AppConfiguration
{
    public RedmineConfiguration RedmineConfiguration { get; set; }

    public PlannerConfiguration PlannerConfiguration { get; set; }
}

internal class PlannerConfiguration
{
    public DateTime Sprint0Start { get; set; }
    public int SprintDays { get; set; }
}

internal class RedmineConfiguration
{
    public string ServerUrl { get; set; }
    public string ApiKey { get; set; }
    public string TargetUserId { get; set; }
}