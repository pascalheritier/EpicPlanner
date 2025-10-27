namespace EpicPlanner.Core;

public static class BusinessCalendar
{
    public static bool IsWeekend(DateTime _InputDate) =>
        _InputDate.DayOfWeek == DayOfWeek.Saturday || _InputDate.DayOfWeek == DayOfWeek.Sunday;

    public static bool IsHoliday(DateTime inputDate, IEnumerable<DateTime> holidays) =>
        holidays != null && holidays.Contains(inputDate.Date);

    public static int CountWorkingDays(
        DateTime _StartDate,
        DateTime _EndDate,
        IEnumerable<DateTime> _Holidays)
    {
        if (_EndDate < _StartDate) return 0;
        int count = 0;
        for (var date = _StartDate.Date; date <= _EndDate.Date; date = date.AddDays(1))
        {
            if (!IsWeekend(date) && !IsHoliday(date, _Holidays)) count++;
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
}
