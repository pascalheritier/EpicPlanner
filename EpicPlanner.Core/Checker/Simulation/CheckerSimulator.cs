using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using EpicPlanner.Core.Checker;
using EpicPlanner.Core.Shared.Models;
using EpicPlanner.Core.Shared.Simulation;
using OfficeOpenXml;

namespace EpicPlanner.Core.Checker.Simulation;

public class CheckerSimulator : SimulatorBase
{
    private readonly Dictionary<string, ResourcePlannedHoursBreakdown> m_PlannedHours;
    private readonly IReadOnlyList<SprintEpicSummary> m_EpicSummaries;
    private readonly Dictionary<string, double> m_PlannedCapacityByEpic;

    public CheckerSimulator(
        List<Epic> _Epics,
        Dictionary<int, Dictionary<string, ResourceCapacity>> _SprintCapacities,
        DateTime _InitialSprintDate,
        int _iSprintDays,
        int _iMaxSprintCount,
        int _iSprintOffset,
        Dictionary<string, ResourcePlannedHoursBreakdown>? _PlannedHours = null,
        IReadOnlyList<SprintEpicSummary>? _EpicSummaries = null,
        IReadOnlyDictionary<string, double>? _PlannedCapacityByEpic = null)
        : base(
            _Epics,
            _SprintCapacities,
            _InitialSprintDate,
            _iSprintDays,
            _iMaxSprintCount,
            _iSprintOffset)
    {
        m_PlannedHours = _PlannedHours is not null
            ? new Dictionary<string, ResourcePlannedHoursBreakdown>(_PlannedHours, StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, ResourcePlannedHoursBreakdown>(StringComparer.OrdinalIgnoreCase);
        m_EpicSummaries = _EpicSummaries ?? Array.Empty<SprintEpicSummary>();
        m_PlannedCapacityByEpic = _PlannedCapacityByEpic is not null
            ? new Dictionary<string, double>(_PlannedCapacityByEpic, StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
    }

    private IReadOnlyDictionary<string, ResourcePlannedHoursBreakdown> PlannedHours => m_PlannedHours;
    private IReadOnlyList<SprintEpicSummary> EpicSummaries => m_EpicSummaries;
    private IReadOnlyDictionary<string, double> PlannedCapacityByEpic => m_PlannedCapacityByEpic;

    public void ExportCheckerReport(string _strOutputExcelPath, EnumCheckerMode _enumMode)
    {
        using ExcelPackage package = new();

        switch (_enumMode)
        {
            case EnumCheckerMode.Comparison:
                WriteComparisonWorksheet(package);
                break;
            case EnumCheckerMode.EpicStates:
                WriteEpicWorksheet(package);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(_enumMode), _enumMode, "Unsupported checker mode.");
        }

        package.SaveAs(new FileInfo(_strOutputExcelPath));
    }

    private void WriteComparisonWorksheet(ExcelPackage _Package)
    {
        var worksheet = _Package.Workbook.Worksheets.Add($"Sprint{SprintOffset}Comparison");
        WriteTable(worksheet, new[]
        {
            "Resource",
            "Capacity_h",
            "Planned_epic_h",
            "Planned_non_epic_h",
            "Planned_total_h",
            "Diff_h"
        });

        int row = 2;
        var sprintCapacities = SprintCapacities[0];
        foreach (var resourceName in sprintCapacities.Keys.OrderBy(k => k, StringComparer.OrdinalIgnoreCase))
        {
            double capacity = sprintCapacities[resourceName].Development;
            ResourcePlannedHoursBreakdown? plannedBreakdown = FindPlannedHoursForResource(resourceName);
            double plannedEpic = plannedBreakdown?.EpicHours ?? 0.0;
            double plannedOutsideEpic = plannedBreakdown?.OutsideEpicHours ?? 0.0;
            double plannedTotal = plannedBreakdown?.TotalHours ?? 0.0;

            worksheet.Cells[row, 1].Value = resourceName;
            worksheet.Cells[row, 2].Value = capacity;
            worksheet.Cells[row, 3].Value = plannedEpic;
            worksheet.Cells[row, 4].Value = plannedOutsideEpic;
            worksheet.Cells[row, 5].Value = plannedTotal;
            worksheet.Cells[row, 6].Formula = $"=B{row}-E{row}";

            row++;
        }

        worksheet.Cells.AutoFitColumns();

        if (row > 2)
        {
            int lastRow = row - 1;
            var diffRange = worksheet.Cells[2, 6, lastRow, 6];

            var condUnder = diffRange.ConditionalFormatting.AddExpression();
            condUnder.Formula = "AND(F2<0,ABS(F2)/B2>0.05)";
            condUnder.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            condUnder.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightCoral);
            condUnder.Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);

            var condOver = diffRange.ConditionalFormatting.AddExpression();
            condOver.Formula = "AND(F2>0,F2/B2>0.15)";
            condOver.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            condOver.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightYellow);
            condOver.Style.Font.Color.SetColor(System.Drawing.Color.DarkOrange);
        }
        else
        {
            worksheet.Cells[2, 1].Value = "No planned hours available.";
        }
    }

    private ResourcePlannedHoursBreakdown? FindPlannedHoursForResource(string _ResourceName)
    {
        if (string.IsNullOrWhiteSpace(_ResourceName))
        {
            return null;
        }

        if (PlannedHours.TryGetValue(_ResourceName, out ResourcePlannedHoursBreakdown? directMatch))
        {
            return directMatch;
        }

        string normalizedResource = NormalizeResourceName(_ResourceName);
        if (string.IsNullOrEmpty(normalizedResource))
        {
            return null;
        }

        List<string> resourceTokens = GetNameTokens(_ResourceName);

        ResourcePlannedHoursBreakdown? bestMatch = null;
        int bestScore = int.MaxValue;

        foreach (var kvp in PlannedHours)
        {
            string candidateName = kvp.Key;
            if (string.IsNullOrWhiteSpace(candidateName))
            {
                continue;
            }

            if (candidateName.Contains(_ResourceName, StringComparison.OrdinalIgnoreCase) ||
                _ResourceName.Contains(candidateName, StringComparison.OrdinalIgnoreCase))
            {
                int score = Math.Abs(candidateName.Length - _ResourceName.Length);
                if (score < bestScore)
                {
                    bestScore = score;
                    bestMatch = kvp.Value;
                }
                continue;
            }

            string normalizedCandidate = NormalizeResourceName(candidateName);
            if (!string.IsNullOrEmpty(normalizedCandidate) &&
                (normalizedCandidate.Contains(normalizedResource) ||
                 normalizedResource.Contains(normalizedCandidate)))
            {
                int score = 1000 + Math.Abs(normalizedCandidate.Length - normalizedResource.Length);
                if (score < bestScore)
                {
                    bestScore = score;
                    bestMatch = kvp.Value;
                    continue;
                }
            }

            if (resourceTokens.Count > 0)
            {
                List<string> candidateTokens = GetNameTokens(candidateName);
                if (TokensMatch(resourceTokens, candidateTokens))
                {
                    int score = 2000 + Math.Abs(candidateTokens.Count - resourceTokens.Count);
                    if (score < bestScore)
                    {
                        bestScore = score;
                        bestMatch = kvp.Value;
                    }
                }
            }
        }

        if (bestMatch != null &&
            !PlannedHours.ContainsKey(_ResourceName))
        {
            // Preserve the association for subsequent lookups within the same session.
            m_PlannedHours[_ResourceName] = bestMatch;
        }

        return bestMatch;
    }

    private static string NormalizeResourceName(string _Name)
    {
        var builder = new System.Text.StringBuilder(_Name.Length);
        foreach (char c in _Name)
        {
            if (char.IsLetterOrDigit(c))
            {
                builder.Append(char.ToLowerInvariant(c));
            }
        }

        return builder.ToString();
    }

    private static List<string> GetNameTokens(string _Name)
    {
        var tokens = new List<string>();
        var builder = new System.Text.StringBuilder();

        foreach (char c in _Name)
        {
            if (char.IsLetterOrDigit(c))
            {
                builder.Append(char.ToLowerInvariant(c));
            }
            else if (builder.Length > 0)
            {
                tokens.Add(builder.ToString());
                builder.Clear();
            }
        }

        if (builder.Length > 0)
        {
            tokens.Add(builder.ToString());
        }

        return tokens;
    }

    private static bool TokensMatch(IReadOnlyList<string> _ResourceTokens, IReadOnlyList<string> _CandidateTokens)
    {
        if (_ResourceTokens.Count == 0 || _CandidateTokens.Count == 0)
        {
            return false;
        }

        foreach (string resourceToken in _ResourceTokens)
        {
            bool found = false;
            foreach (string candidateToken in _CandidateTokens)
            {
                if (candidateToken.Contains(resourceToken, StringComparison.OrdinalIgnoreCase) ||
                    resourceToken.Contains(candidateToken, StringComparison.OrdinalIgnoreCase))
                {
                    found = true;
                    break;
                }
            }

            if (!found)
            {
                return false;
            }
        }

        return true;
    }

    private void WriteEpicWorksheet(ExcelPackage _Package)
    {
        var epicSheet = _Package.Workbook.Worksheets.Add($"Sprint{SprintOffset}EpicCheck");
        WriteTable(epicSheet, new[]
        {
            "Epic",
            "Initial_remaining_h",
            "Planned_capacity_h",
            "Consumed_h",
            "Actual_remaining_h",
            "Projected_remaining_h",
            "Delta_remaining_h",
            "Planning_reliability_rate",
            "Capacity_usage_rate"
        });

        if (EpicSummaries.Count == 0)
        {
            epicSheet.Cells[2, 1].Value = "No epic sprint data available.";
            epicSheet.Cells.AutoFitColumns();
            return;
        }

        var initialRemainingByEpic = Epics
            .GroupBy(e => e.Name, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First().Charge, StringComparer.OrdinalIgnoreCase);

        int epicRow = 2;
        bool hasPlannedCapacityLookup = PlannedCapacityByEpic.Count > 0;
        HashSet<string> missingPlanningEntries = new(StringComparer.OrdinalIgnoreCase);
        foreach (var summary in EpicSummaries
                     .OrderBy(s => s.Epic, StringComparer.OrdinalIgnoreCase))
        {
            epicSheet.Cells[epicRow, 1].Value = summary.Epic;

            bool hasInitial = initialRemainingByEpic.TryGetValue(summary.Epic, out double initialRemaining);
            double plannedCapacityRaw = summary.PlannedCapacity;
            double consumedRaw = summary.Consumed;
            double actualRemainingRaw = summary.Remaining;
            double plannedCapacity = Math.Round(plannedCapacityRaw, 2);
            double consumed = Math.Round(consumedRaw, 2);
            double actualRemaining = Math.Round(actualRemainingRaw, 2);

            if (hasPlannedCapacityLookup && !PlannedCapacityByEpic.ContainsKey(summary.Epic))
            {
                if (missingPlanningEntries.Add(summary.Epic))
                {
                    Console.WriteLine($"Warning: Planned capacity for epic '{summary.Epic}' was not found in the planning sheet. Using Redmine estimates instead.");
                }
            }

            if (hasInitial)
            {
                epicSheet.Cells[epicRow, 2].Value = Math.Round(initialRemaining, 2);
            }
            epicSheet.Cells[epicRow, 3].Value = plannedCapacity;
            epicSheet.Cells[epicRow, 4].Value = consumed;
            epicSheet.Cells[epicRow, 5].Value = actualRemaining;
            epicSheet.Cells[epicRow, 6].Formula = $"=IF(ISBLANK(B{epicRow}),E{epicRow},MAX(0,B{epicRow}-C{epicRow}))";
            epicSheet.Cells[epicRow, 7].Formula = $"=IF(ISBLANK(B{epicRow}),\"\",B{epicRow}-E{epicRow})";

            string reliabilityFormula =
                $"=IF(ISBLANK(B{epicRow}),\"\",IF(C{epicRow}<=0,MIN(1,MAX(0,1-ABS(F{epicRow}-E{epicRow})/MAX(ABS(F{epicRow}),0.0001))),MIN(1,MAX(0,1-ABS(F{epicRow}-E{epicRow})/C{epicRow}))))";
            string usageFormula =
                $"=IF(C{epicRow}<=0,IF(D{epicRow}<=0,1,2),D{epicRow}/C{epicRow})";

            epicSheet.Cells[epicRow, 8].Formula = reliabilityFormula;
            epicSheet.Cells[epicRow, 9].Formula = usageFormula;

            epicRow++;
        }

        int lastEpicRow = epicRow - 1;
        if (lastEpicRow >= 2)
        {
            var reliabilityRange = epicSheet.Cells[2, 8, lastEpicRow, 8];
            var usageRange = epicSheet.Cells[2, 9, lastEpicRow, 9];

            reliabilityRange.Style.Numberformat.Format = "0.00%";
            usageRange.Style.Numberformat.Format = "0.00%";

            var condReliability = reliabilityRange.ConditionalFormatting.AddExpression();
            condReliability.Formula = "AND($H2<>\"\",$H2<0.8)";
            condReliability.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            condReliability.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightYellow);
            condReliability.Style.Font.Color.SetColor(System.Drawing.Color.DarkOrange);

            var condUsage = usageRange.ConditionalFormatting.AddExpression();
            condUsage.Formula = "AND($I2<>\"\",OR($I2<0.8,$I2>1.2))";
            condUsage.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            condUsage.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightCoral);
            condUsage.Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);
        }

        epicSheet.Cells.AutoFitColumns();
    }
}
