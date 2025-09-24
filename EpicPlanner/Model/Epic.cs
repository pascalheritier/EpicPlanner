using System.Globalization;
using System.Text.RegularExpressions;

namespace EpicPlanner
{
    internal class Epic
    {
        public string Name { get; }
        public string State { get; }
        public double Charge { get; }
        public double Remaining { get; set; }
        public DateTime? EndAnalysis { get; }
        public List<Wish> Wishes { get; } = new();
        public List<string> Dependencies { get; } = new();

        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }

        public List<Allocation> History { get; } = new();

        public Epic(string name, string state, double charge, DateTime? endAnalysis)
        {
            Name = (name ?? "").Trim();
            State = (state ?? "").Trim().ToLowerInvariant();
            Charge = charge;
            Remaining = charge;
            EndAnalysis = endAnalysis;
        }

        public bool IsInDevelopment => State.Contains("develop") && !State.Contains("pending");
        public bool IsOtherAllowed => (State.Contains("analysis") || State.Contains("pending"));

        public void ParseAssignments(string assigned, string willAssign, List<string> resourceNames)
        {
            foreach (var raw in (assigned + "," + willAssign).Split(',', ';'))
            {
                string s = raw.Trim();
                if (string.IsNullOrWhiteSpace(s)) continue;

                // Parse optional trailing percentage
                var m = Regex.Match(s, @"(\d{1,3})\s*%$");
                double pct = 1.0;
                string name = s;
                if (m.Success)
                {
                    pct = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture) / 100.0;
                    name = s.Substring(0, m.Index).Trim();
                }

                // Fuzzy match a resource (exact or contains)
                string matched = resourceNames.FirstOrDefault(r =>
                    r.Equals(name, StringComparison.OrdinalIgnoreCase) ||
                    r.Contains(name, StringComparison.OrdinalIgnoreCase));

                if (!string.IsNullOrWhiteSpace(matched))
                    Wishes.Add(new Wish(matched, pct));
            }
        }

        public void ParseDependencies(string depRaw)
        {
            if (string.IsNullOrWhiteSpace(depRaw)) return;
            foreach (var d in depRaw.Split(',', ';'))
            {
                string dep = d.Trim();
                if (!string.IsNullOrWhiteSpace(dep))
                    Dependencies.Add(dep);
            }
        }
    }
}
