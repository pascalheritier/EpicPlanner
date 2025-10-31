namespace EpicPlanner.Core.Shared.Models;

public class Allocation
{
    public Allocation(string epic, int sprint, string resource, double hours, DateTime sprintStart)
    {
        Epic = epic;
        Sprint = sprint;
        Resource = resource;
        Hours = hours;
        SprintStart = sprintStart;
    }

    public string Epic { get; }
    public int Sprint { get; }
    public string Resource { get; }
    public double Hours { get; }
    public DateTime SprintStart { get; }
}
