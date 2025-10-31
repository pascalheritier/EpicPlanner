namespace EpicPlanner.Core.Shared.Models;

public class Wish
{
    public Wish(string resource, double percentage)
    {
        Resource = resource;
        Percentage = percentage;
    }

    public string Resource { get; }
    public double Percentage { get; }
}
