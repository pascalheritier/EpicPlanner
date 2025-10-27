namespace EpicPlanner.Core;

public static class BusinessCalendar
{
    public static bool IsWeekend(DateTime inputDate) =>
        inputDate.DayOfWeek == DayOfWeek.Saturday || inputDate.DayOfWeek == DayOfWeek.Sunday;

    public static bool IsHoliday(DateTime inputDate, IEnumerable<DateTime> holidays) =>
        holidays != null && holidays.Contains(inputDate.Date);

    public static int CountWorkingDays(
        DateTime startDate,
        DateTime endDate,
        IEnumerable<DateTime> holidays)
    {
        if (endDate < startDate) return 0;
        int count = 0;
        for (var date = startDate.Date; date <= endDate.Date; date = date.AddDays(1))
        {
            if (!IsWeekend(date) && !IsHoliday(date, holidays)) count++;
        }
        return count;
    }

    public static int CountWorkingDaysOverlap(
        DateTime absenceStartDate,
        DateTime absenceEndDate,
        DateTime sprintStartDate,
        DateTime sprintEndDate,
        IEnumerable<DateTime> holidays)
    {
        var start = (absenceStartDate > sprintStartDate) ? absenceStartDate.Date : sprintStartDate.Date;
        var end = (absenceEndDate < sprintEndDate) ? absenceEndDate.Date : sprintEndDate.Date;
        if (end < start) return 0;
        return CountWorkingDays(start, end, holidays);
    }
}
