public static class BusinessCalendar
{
    public static bool IsWeekend(DateTime d) =>
        d.DayOfWeek == DayOfWeek.Saturday || d.DayOfWeek == DayOfWeek.Sunday;

    public static bool IsHoliday(DateTime d, IEnumerable<DateTime> holidays) =>
        holidays != null && holidays.Contains(d.Date);

    public static int CountWorkingDays(DateTime start, DateTime end, IEnumerable<DateTime> holidays)
    {
        if (end < start) return 0;
        int count = 0;
        for (var d = start.Date; d <= end.Date; d = d.AddDays(1))
        {
            if (!IsWeekend(d) && !IsHoliday(d, holidays)) count++;
        }
        return count;
    }

    public static int CountWorkingDaysOverlap(DateTime aStart, DateTime aEnd, DateTime sStart, DateTime sEnd, IEnumerable<DateTime> holidays)
    {
        var start = (aStart > sStart) ? aStart.Date : sStart.Date;
        var end = (aEnd < sEnd) ? aEnd.Date : sEnd.Date;
        if (end < start) return 0;
        return CountWorkingDays(start, end, holidays);
    }
}
