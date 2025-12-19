using System.Globalization;
using System.Text.RegularExpressions;

namespace EpicPlanner.Core.Shared.Models;

public class Epic
{
    public Epic(string name, string state, double charge, DateTime? endAnalysis)
    {
        Name = (name ?? string.Empty).Trim();
        State = (state ?? string.Empty).Trim().ToLowerInvariant();
        Charge = charge;
        Remaining = charge;
        EndAnalysis = endAnalysis;
    }

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
    public EnumEpicPriority Priority { get; set; } = EnumEpicPriority.Normal;
    public string Group { get; set; } = string.Empty;

    public bool IsInDevelopment => State.Contains("develop") && !State.Contains("pending");
    public bool IsOtherAllowed => State.Contains("analysis") || State.Contains("pending");

    public void ParseAssignments(string assigned, string willAssign, List<string> resourceNames)
    {
        foreach (var raw in (assigned + "," + willAssign).Split(',', ';'))
        {
            string candidate = raw.Trim();
            if (string.IsNullOrWhiteSpace(candidate)) continue;

            var match = Regex.Match(candidate, @"(\d{1,3})\s*%$");
            double pct = 1.0;
            string name = candidate;
            if (match.Success)
            {
                pct = int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture) / 100.0;
                name = candidate[..match.Index].Trim();
            }

            string? matched = resourceNames.FirstOrDefault(r =>
                r.Equals(name, StringComparison.OrdinalIgnoreCase) ||
                r.Contains(name, StringComparison.OrdinalIgnoreCase));

            if (!string.IsNullOrWhiteSpace(matched))
                Wishes.Add(new Wish(matched, pct));
        }
    }

    public void ParseDependencies(string dependencyRaw)
    {
        if (string.IsNullOrWhiteSpace(dependencyRaw)) return;
        foreach (var dependency in dependencyRaw.Split(',', ';'))
        {
            string dep = dependency.Trim();
            if (!string.IsNullOrWhiteSpace(dep))
                Dependencies.Add(dep);
        }
    }
}
