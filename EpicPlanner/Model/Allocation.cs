namespace EpicPlanner
{
    internal class Allocation
    {
        public string Epic { get; }
        public int Sprint { get; }
        public string Resource { get; }
        public double Hours { get; }
        public DateTime SprintStart { get; }

        public Allocation(string epic, int sprint, string resource, double hours, DateTime start)
        {
            Epic = epic;
            Sprint = sprint;
            Resource = resource;
            Hours = hours;
            SprintStart = start;
        }
    }
}
