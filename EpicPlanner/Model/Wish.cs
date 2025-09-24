namespace EpicPlanner
{
    internal class Wish
    {
        public string Resource { get; }
        public double Pct { get; }
        public Wish(string resource, double pct)
        {
            Resource = resource;
            Pct = pct;
        }
    }
}
