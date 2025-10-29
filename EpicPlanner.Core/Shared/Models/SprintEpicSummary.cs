namespace EpicPlanner.Core.Shared.Models;

public class SprintEpicSummary
{
    public string Epic { get; init; } = string.Empty;

    public double PlannedCapacity { get; set; }

    public double Consumed { get; set; }

    public double Remaining { get; set; }

    public double ExpectedRemaining
    {
        get
        {
            double expected = (Consumed + Remaining) - PlannedCapacity;
            return expected > 0 ? expected : 0.0;
        }
    }

    public double OverheadHours => (Consumed + Remaining) - PlannedCapacity;

    public double OverheadRatio => PlannedCapacity > 1e-9 ? OverheadHours / PlannedCapacity : 0.0;
}
