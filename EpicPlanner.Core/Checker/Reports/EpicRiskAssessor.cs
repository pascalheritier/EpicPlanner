namespace EpicPlanner.Core.Checker.Reports;

/// <summary>
/// Determines the risk level of an epic based on its allocation and remaining history.
/// </summary>
public static class EpicRiskAssessor
{
    /// <summary>
    /// Returns (risk, riskSince, riskDesc) where risk is "critical" | "watch" | "ok".
    /// </summary>
    public static (string Risk, string RiskSince, string RiskDesc) Assess(
        string       epicId,
        string       stateJs,        // "in_dev" | "pending"
        double?      originalEstimate,
        double       currentRemaining,
        double[]     allocation,
        double?[]    remaining,
        List<string> sprintLabels)
    {
        int n = allocation.Length;
        if (n == 0) return ("ok", "—", string.Empty);

        // ── Critical: never allocated ────────────────────────────────────────────
        bool neverAllocated = allocation.All(a => a == 0.0);
        if (neverAllocated && stateJs != "done")
            return ("critical", sprintLabels.FirstOrDefault() ?? "—", "Jamais alloué dans le plan");

        // ── Critical: dropped from plan (≥3 trailing zeros) ─────────────────────
        int trailingZeros = 0;
        for (int i = n - 1; i >= 0; i--)
        {
            if (allocation[i] == 0.0) trailingZeros++;
            else break;
        }
        if (trailingZeros >= 3)
        {
            string since = trailingZeros < n
                ? sprintLabels[n - trailingZeros]
                : (sprintLabels.FirstOrDefault() ?? "—");
            return ("critical", since, $"Retiré du plan depuis {trailingZeros} sprints");
        }

        // ── Critical: zero consumption despite allocation ────────────────────────
        double totalAllocated = allocation.Sum();
        double firstRemaining = remaining.FirstOrDefault(r => r.HasValue) ?? currentRemaining;
        bool   noConsumption  = totalAllocated > 0
                                && currentRemaining > 0
                                && currentRemaining >= firstRemaining - 1.0;
        if (noConsumption && stateJs == "in_dev")
            return ("critical", sprintLabels.FirstOrDefault() ?? "—", "Aucune consommation malgré les allocations");

        // ── Critical: pace > 20 sprints at current rate ──────────────────────────
        double consumed = firstRemaining - currentRemaining;
        if (consumed > 0 && currentRemaining > 0)
        {
            double sprintsWithAlloc = allocation.Count(a => a > 0);
            double pace = sprintsWithAlloc > 0
                ? (currentRemaining / (consumed / sprintsWithAlloc))
                : double.MaxValue;
            if (pace > 20)
                return ("critical", sprintLabels.LastOrDefault() ?? "—",
                    $"Rythme de {pace:F0} sprints estimés à ce rythme");
        }

        // ── Watch: last allocation < 60% of previous ────────────────────────────
        if (n >= 2)
        {
            double last = allocation[n - 1];
            double prev = allocation[n - 2];
            if (prev > 0 && last > 0 && last < prev * 0.6)
                return ("watch", sprintLabels[n - 1],
                    $"Allocation en baisse ({last:F0}h vs {prev:F0}h sprint précédent)");
        }

        // ── Watch: re-estimation upward ──────────────────────────────────────────
        if (originalEstimate.HasValue && originalEstimate.Value > 0
            && currentRemaining > originalEstimate.Value * 1.1)
        {
            return ("watch", sprintLabels.LastOrDefault() ?? "—",
                $"Ré-estimation à la hausse ({currentRemaining:F0}h vs {originalEstimate.Value:F0}h initial)");
        }

        // ── Watch: late start (> 60% of sprint history without allocation) ───────
        int firstAllocSprint = -1;
        for (int i = 0; i < n; i++)
        {
            if (allocation[i] > 0) { firstAllocSprint = i; break; }
        }
        if (firstAllocSprint > 0 && (double)firstAllocSprint / n > 0.6 && stateJs == "in_dev")
            return ("watch", sprintLabels[firstAllocSprint],
                "Démarrage tardif dans le plan");

        return ("ok", "—", string.Empty);
    }
}
