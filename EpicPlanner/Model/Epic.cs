using System.Globalization;
using System.Text.RegularExpressions;

namespace EpicPlanner;

internal class Epic
{
    #region Constructor

    public Epic(string _strName, string _strState, double _dCharge, DateTime? _EndAnalysis)
    {
        Name = (_strName ?? "").Trim();
        State = (_strState ?? "").Trim().ToLowerInvariant();
        Charge = _dCharge;
        Remaining = _dCharge;
        EndAnalysis = _EndAnalysis;
    }

    #endregion

    #region Description

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

    #endregion

    #region Epic logic

    public bool IsInDevelopment => State.Contains("develop") && !State.Contains("pending");
    public bool IsOtherAllowed => (State.Contains("analysis") || State.Contains("pending"));

    public void ParseAssignments(string _strAssigned, string _strWillAssign, List<string> _ResourceNames)
    {
        foreach (var raw in (_strAssigned + "," + _strWillAssign).Split(',', ';'))
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
            string matched = _ResourceNames.FirstOrDefault(r =>
                r.Equals(name, StringComparison.OrdinalIgnoreCase) ||
                r.Contains(name, StringComparison.OrdinalIgnoreCase));

            if (!string.IsNullOrWhiteSpace(matched))
                Wishes.Add(new Wish(matched, pct));
        }
    }

    public void ParseDependencies(string _strDependencyRaw)
    {
        if (string.IsNullOrWhiteSpace(_strDependencyRaw)) return;
        foreach (var d in _strDependencyRaw.Split(',', ';'))
        {
            string dep = d.Trim();
            if (!string.IsNullOrWhiteSpace(dep))
                Dependencies.Add(dep);
        }
    }

    #endregion
}