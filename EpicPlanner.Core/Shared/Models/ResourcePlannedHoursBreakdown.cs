namespace EpicPlanner.Core.Shared.Models;

public class ResourcePlannedHoursBreakdown
{
    public double EpicHours { get; private set; }

    public double OutsideEpicHours { get; private set; }

    public double MaintenanceHours => OutsideEpicHours;

    public double TotalHours => EpicHours + OutsideEpicHours;

    public void AddEpicHours(double _Hours)
    {
        if (_Hours <= 0)
        {
            return;
        }

        EpicHours += _Hours;
    }

    public void AddOutsideEpicHours(double _Hours)
    {
        if (_Hours <= 0)
        {
            return;
        }

        OutsideEpicHours += _Hours;
    }

    public void AddHours(double _Hours, bool _bIsEpic)
    {
        if (_bIsEpic)
        {
            AddEpicHours(_Hours);
        }
        else
        {
            AddOutsideEpicHours(_Hours);
        }
    }
}
