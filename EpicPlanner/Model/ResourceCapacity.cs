namespace EpicPlanner;

public class ResourceCapacity
{
    #region Constructor

    public ResourceCapacity()
    {
    }

    public ResourceCapacity(ResourceCapacity _OriginalCapacity)
        : this()
    {
        Development = _OriginalCapacity.Development;
        Maintenance = _OriginalCapacity.Maintenance;
        Analysis = _OriginalCapacity.Analysis;
    }

    #endregion

    #region Description

    public double Development { get; set; }
    public double Maintenance { get; set; }
    public double Analysis { get; set; }

    #endregion

    #region Behaviour

    public void AdapteCapacityToScale(double _dScale)
    {
        Development *= _dScale;
        Maintenance *= _dScale;
        Analysis *= _dScale;
    }

    public void AdaptCapacityToAbsences(double _dWorkingDaysInSprint, double _dAbsentWeekDays)
    {
        Development = this.AdaptCapacityToAbsences(Development, _dWorkingDaysInSprint, _dAbsentWeekDays);
        Maintenance = this.AdaptCapacityToAbsences(Maintenance, _dWorkingDaysInSprint, _dAbsentWeekDays);
        Analysis = this.AdaptCapacityToAbsences(Analysis, _dWorkingDaysInSprint, _dAbsentWeekDays);
    }

    private double AdaptCapacityToAbsences(double _dInputValue, double _dWorkingDaysInSprint, double _dAbsentWeekDays)
    {
        double dailyCap = _dInputValue / _dWorkingDaysInSprint;
        _dInputValue -= dailyCap * _dAbsentWeekDays;
        return _dInputValue;
    }

    public void RoundUpCapacity()
    {
        Development = RoundUpCapacity(Development);
        Maintenance = RoundUpCapacity(Maintenance);
        Analysis = RoundUpCapacity(Analysis);
    }

    private double RoundUpCapacity(double _dInputValue)
    {
        if (_dInputValue < 0)
            _dInputValue = 0;
        return Math.Round(_dInputValue, 2);
    }

    #endregion
}