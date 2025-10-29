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
    public CheckerSimulator(
        List<Epic> _Epics,
        Dictionary<int, Dictionary<string, ResourceCapacity>> _SprintCapacities,
        DateTime _InitialSprintDate,
        int _iSprintDays,
        int _iMaxSprintCount,
        int _iSprintOffset,
        Dictionary<string, double>? _PlannedHours = null,
        IReadOnlyList<SprintEpicSummary>? _EpicSummaries = null,
        IReadOnlyDictionary<string, double>? _PlannedCapacityByEpic = null)
        : base(
            _Epics,
            _SprintCapacities,
            _InitialSprintDate,
            _iSprintDays,
            _iMaxSprintCount,
            _iSprintOffset,
            _PlannedHours,
            _EpicSummaries,
            _PlannedCapacityByEpic)
    {
    }

    public void ExportCheckerReport(string _strOutputExcelPath, CheckerMode _enumMode)
    {
        using ExcelPackage package = new();

        switch (_enumMode)
        {
            case CheckerMode.Comparison:
                WriteComparisonWorksheet(package);
                break;
            case CheckerMode.EpicStates:
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
        WriteTable(worksheet, new[] { "Resource", "Capacity_h", "Planned_h", "Diff_h" });

        int row = 2;
        var sprintCapacities = SprintCapacities[0];
        foreach (var resourceName in sprintCapacities.Keys.OrderBy(k => k, StringComparer.OrdinalIgnoreCase))
        {
            double capacity = sprintCapacities[resourceName].Development;
            double planned = PlannedHours.TryGetValue(resourceName, out var plannedHours) ? plannedHours : 0.0;

            worksheet.Cells[row, 1].Value = resourceName;
            worksheet.Cells[row, 2].Value = capacity;
            worksheet.Cells[row, 3].Value = planned;
            worksheet.Cells[row, 4].Formula = $"=B{row}-C{row}";

            row++;
        }

        worksheet.Cells.AutoFitColumns();

        if (row > 2)
        {
            int lastRow = row - 1;
            var diffRange = worksheet.Cells[2, 4, lastRow, 4];

            var condUnder = diffRange.ConditionalFormatting.AddExpression();
            condUnder.Formula = "AND(D2<0,ABS(D2)/B2>0.05)";
            condUnder.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            condUnder.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightCoral);
            condUnder.Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);

            var condOver = diffRange.ConditionalFormatting.AddExpression();
            condOver.Formula = "AND(D2>0,D2/B2>0.15)";
            condOver.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            condOver.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightYellow);
            condOver.Style.Font.Color.SetColor(System.Drawing.Color.DarkOrange);
        }
        else
        {
            worksheet.Cells[2, 1].Value = "No planned hours available.";
        }
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
                $"=IF(ISBLANK(B{epicRow}),\"\",IF(C{epicRow}<=0,IF(D{epicRow}>0,0,\"\"),MIN(1,MAX(0,1-ABS(F{epicRow}-E{epicRow})/C{epicRow}))))";
            string usageFormula =
                $"=IF(C{epicRow}<=0,IF(D{epicRow}>0,2,\"\"),D{epicRow}/C{epicRow})";

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
