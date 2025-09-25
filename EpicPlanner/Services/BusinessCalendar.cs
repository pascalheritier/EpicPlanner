namespace EpicPlanner;

public static class BusinessCalendar
{
    #region Date helpers

    public static bool IsWeekend(DateTime _InputDate) =>
        _InputDate.DayOfWeek == DayOfWeek.Saturday || _InputDate.DayOfWeek == DayOfWeek.Sunday;

    public static bool IsHoliday(DateTime _InputDate, IEnumerable<DateTime> _Holidays) =>
        _Holidays != null && _Holidays.Contains(_InputDate.Date);

    public static int CountWorkingDays(
        DateTime _StartDate,
        DateTime _EndDate,
        IEnumerable<DateTime> _Holidays)
    {
        if (_EndDate < _StartDate) return 0;
        int count = 0;
        for (var d = _StartDate.Date; d <= _EndDate.Date; d = d.AddDays(1))
        {
            if (!IsWeekend(d) && !IsHoliday(d, _Holidays)) count++;
        }
        return count;
    }

    public static int CountWorkingDaysOverlap(
        DateTime _AbsenceStartDate,
        DateTime _AbsenceEndDate,
        DateTime _SprintStartDate,
        DateTime _SprintEndDate,
        IEnumerable<DateTime> _Holidays)
    {
        var start = (_AbsenceStartDate > _SprintStartDate) ? _AbsenceStartDate.Date : _SprintStartDate.Date;
        var end = (_AbsenceEndDate < _SprintEndDate) ? _AbsenceEndDate.Date : _SprintEndDate.Date;
        if (end < start) return 0;
        return CountWorkingDays(start, end, _Holidays);
    }

    #endregion
}