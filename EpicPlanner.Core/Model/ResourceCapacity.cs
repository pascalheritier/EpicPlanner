namespace EpicPlanner.Core;

public class ResourceCapacity
{
    public ResourceCapacity()
    {
    }

    public ResourceCapacity(ResourceCapacity originalCapacity)
    {
        Development = originalCapacity.Development;
        Maintenance = originalCapacity.Maintenance;
        Analysis = originalCapacity.Analysis;
    }

    public double Development { get; set; }
    public double Maintenance { get; set; }
    public double Analysis { get; set; }

    public void AdapteCapacityToScale(double scale)
    {
        Development *= scale;
        Maintenance *= scale;
        Analysis *= scale;
    }

    public void AdaptCapacityToAbsences(double workingDaysInSprint, double absentWeekDays)
    {
        Development = AdaptCapacityToAbsences(Development, workingDaysInSprint, absentWeekDays);
        Maintenance = AdaptCapacityToAbsences(Maintenance, workingDaysInSprint, absentWeekDays);
        Analysis = AdaptCapacityToAbsences(Analysis, workingDaysInSprint, absentWeekDays);
    }

    private static double AdaptCapacityToAbsences(double inputValue, double workingDaysInSprint, double absentWeekDays)
    {
        double dailyCap = inputValue / workingDaysInSprint;
        inputValue -= dailyCap * absentWeekDays;
        return inputValue;
    }

    public void RoundUpCapacity()
    {
        Development = RoundUpCapacity(Development);
        Maintenance = RoundUpCapacity(Maintenance);
        Analysis = RoundUpCapacity(Analysis);
    }

    private static double RoundUpCapacity(double inputValue)
    {
        if (inputValue < 0)
            inputValue = 0;
        return Math.Round(inputValue, 2);
    }
}
