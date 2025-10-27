namespace EpicPlanner.Core;

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
