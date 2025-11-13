using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using EpicPlanner.Core.Planner;
using EpicPlanner.Core.Shared.Models;
using EpicPlanner.Core.Shared.Simulation;
using OfficeOpenXml;
using SkiaSharp;

namespace EpicPlanner.Core.Planner.Simulation;

public class PlannerSimulator : SimulatorBase
{
    public PlannerSimulator(
        List<Epic> _Epics,
        Dictionary<int, Dictionary<string, ResourceCapacity>> _SprintCapacities,
        DateTime _InitialSprintDate,
        int _iSprintDays,
        int _iMaxSprintCount,
        int _iSprintOffset,
        bool _OnlyDevelopmentEpics = false)
        : base(
            _Epics,
            _SprintCapacities,
            _InitialSprintDate,
            _iSprintDays,
            _iMaxSprintCount,
            _iSprintOffset,
            _OnlyDevelopmentEpics)
    {
    }

    public void ExportPlanningExcel(string _strOutputExcelFilePath)
    {
        using var package = new ExcelPackage();
        var wsFinal = package.Workbook.Worksheets.Add("FinalSchedule");
        var wsA1 = package.Workbook.Worksheets.Add("AllocationsByEpicAndSprint");
        var wsA2 = package.Workbook.Worksheets.Add("AllocationsByEpicPerSprint");
        var wsVer = package.Workbook.Worksheets.Add("Verification");
        var wsPS = package.Workbook.Worksheets.Add("PerSprintSummary");
        var wsUnder = package.Workbook.Worksheets.Add("Underutilization");
        var wsOver = package.Workbook.Worksheets.Add("OverBooking");
        var wsMaint = package.Workbook.Worksheets.Add("MaintenanceCapacities");
        var wsAnal = package.Workbook.Worksheets.Add("AnalysisCapacities");

        WriteTable(wsFinal, new[] { "Epic", "State", "Priority", "Initial_Charge_h", "Allocated_total_h", "Remaining_after_h", "Start_date", "End_date" });
        var finalRows = Epics.Select(e => new
        {
            Epic = e.Name,
            State = e.State,
            Priority = e.Priority.ToString(),
            Initial_Charge_h = e.Charge,
            Allocated_total_h = Math.Round(e.History.Sum(h => h.Hours), 2),
            Remaining_after_h = Math.Round(Math.Max(0, e.Remaining), 2),
            Start_date = e.StartDate?.ToString("yyyy-MM-dd") ?? string.Empty,
            End_date = e.EndDate?.ToString("yyyy-MM-dd") ?? string.Empty
        }).ToList();

        var finalSorted = finalRows
            .Select(r => new
            {
                r.Epic,
                r.State,
                r.Priority,
                r.Initial_Charge_h,
                r.Allocated_total_h,
                r.Remaining_after_h,
                r.Start_date,
                r.End_date,
                Start_dt = TryParseDate(r.Start_date),
                Key = ExtractEpicKey(r.Epic)
            })
            .OrderBy(r => r.Start_dt)
            .ThenBy(r => r.Key.Year)
            .ThenBy(r => r.Key.Num)
            .ToList();

        int row = 2;
        foreach (var r in finalSorted)
        {
            wsFinal.Cells[row, 1].Value = r.Epic;
            wsFinal.Cells[row, 2].Value = r.State;
            wsFinal.Cells[row, 3].Value = r.Priority;
            wsFinal.Cells[row, 4].Value = r.Initial_Charge_h;
            wsFinal.Cells[row, 5].Value = r.Allocated_total_h;
            wsFinal.Cells[row, 6].Value = r.Remaining_after_h;
            wsFinal.Cells[row, 7].Value = r.Start_date;
            wsFinal.Cells[row, 8].Value = r.End_date;
            row++;
        }
        wsFinal.Cells.AutoFitColumns();

        var aggA1 = AllocationHistory
            .GroupBy(a => new { a.Epic, a.Sprint, a.Resource })
            .Select(g => new { g.Key.Epic, g.Key.Sprint, g.Key.Resource, Hours = Math.Round(g.Sum(x => x.Hours), 2) })
            .OrderBy(x => x.Sprint)
            .ThenBy(x => ExtractEpicKey(x.Epic).Year)
            .ThenBy(x => ExtractEpicKey(x.Epic).Num)
            .ThenBy(x => x.Resource, StringComparer.OrdinalIgnoreCase);
        WriteTable(wsA1, new[] { "Epic", "Sprint", "Resource", "Hours" });
        row = 2;
        foreach (var r in aggA1)
        {
            wsA1.Cells[row, 1].Value = r.Epic;
            wsA1.Cells[row, 2].Value = r.Sprint + SprintOffset;
            wsA1.Cells[row, 3].Value = r.Resource;
            wsA1.Cells[row, 4].Value = r.Hours;
            row++;
        }
        wsA1.Cells.AutoFitColumns();

        var aggA2 = AllocationHistory
            .GroupBy(a => new { a.Epic, a.Sprint, a.SprintStart })
            .Select(g => new { g.Key.Epic, g.Key.Sprint, g.Key.SprintStart, Total_Hours = Math.Round(g.Sum(x => x.Hours), 2) })
            .OrderBy(x => x.Sprint)
            .ThenBy(x => ExtractEpicKey(x.Epic).Year)
            .ThenBy(x => ExtractEpicKey(x.Epic).Num);
        WriteTable(wsA2, new[] { "Epic", "Sprint", "Sprint_start", "Total_Hours" });
        row = 2;
        foreach (var r in aggA2)
        {
            wsA2.Cells[row, 1].Value = r.Epic;
            wsA2.Cells[row, 2].Value = r.Sprint + SprintOffset;
            wsA2.Cells[row, 3].Value = r.SprintStart.ToString("yyyy-MM-dd");
            wsA2.Cells[row, 4].Value = r.Total_Hours;
            row++;
        }
        wsA2.Cells.AutoFitColumns();

        WriteTable(wsVer, new[] { "Epic", "Initial_Charge_h", "Allocated_total_h", "Delta_h" });
        row = 2;
        foreach (var e in Epics
                     .OrderBy(x => x.StartDate ?? DateTime.MaxValue)
                     .ThenBy(x => ExtractEpicKey(x.Name).Year)
                     .ThenBy(x => ExtractEpicKey(x.Name).Num))
        {
            double allocated = Math.Round(e.History.Sum(h => h.Hours), 2);
            double delta = Math.Round(e.Charge - allocated, 2);
            wsVer.Cells[row, 1].Value = e.Name;
            wsVer.Cells[row, 2].Value = e.Charge;
            wsVer.Cells[row, 3].Value = allocated;
            wsVer.Cells[row, 4].Value = delta;
            row++;
        }
        wsVer.Cells.AutoFitColumns();

        var sprintIndexes = AllocationHistory.Select(a => a.Sprint).Distinct().OrderBy(s => s).ToList();
        if (sprintIndexes.Count == 0)
        {
            sprintIndexes.Add(0);
        }

        var resources = SprintCapacities.Values
            .SelectMany(dict => dict.Keys)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(k => k, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var psHeaders = new List<string> { "Sprint", "Sprint_start", "Sprint_end" };
        foreach (var resourceName in resources)
        {
            psHeaders.Add($"{resourceName}_allocated_h");
            psHeaders.Add($"{resourceName}_capacity_h");
            psHeaders.Add($"{resourceName}_util_pct");
        }
        WriteTable(wsPS, psHeaders);

        row = 2;
        foreach (int sprintIndex in sprintIndexes)
        {
            DateTime start = SprintStartDate(sprintIndex);
            DateTime end = start.AddDays(SprintLengthDays - 1);
            int col = 1;
            wsPS.Cells[row, col++].Value = sprintIndex + SprintOffset;
            wsPS.Cells[row, col++].Value = start.ToString("yyyy-MM-dd");
            wsPS.Cells[row, col++].Value = end.ToString("yyyy-MM-dd");
            foreach (var resourceName in resources)
            {
                double allocation = Math.Round(AllocationHistory
                    .Where(a => a.Sprint == sprintIndex && a.Resource.Equals(resourceName, StringComparison.OrdinalIgnoreCase))
                    .Sum(a => a.Hours), 2);
                double capacity = SprintCapacities[sprintIndex].TryGetValue(resourceName, out var rc) ? rc.Development : 0.0;
                double pct = capacity > 0 ? Math.Round(allocation / capacity * 100.0, 2) : 0.0;
                wsPS.Cells[row, col++].Value = allocation;
                wsPS.Cells[row, col++].Value = capacity;
                wsPS.Cells[row, col++].Value = pct;
            }
            row++;
        }
        wsPS.Cells.AutoFitColumns();

        WriteTable(wsMaint, new[] { "Sprint", "Resource", "Maintenance_h" });
        row = 2;
        foreach (int sprintIndex in sprintIndexes)
        {
            foreach (var resourceName in resources)
            {
                double maintenance = SprintCapacities[sprintIndex].TryGetValue(resourceName, out var rc) ? rc.Maintenance : 0.0;
                wsMaint.Cells[row, 1].Value = sprintIndex + SprintOffset;
                wsMaint.Cells[row, 2].Value = resourceName;
                wsMaint.Cells[row, 3].Value = maintenance;
                row++;
            }
        }
        wsMaint.Cells.AutoFitColumns();

        WriteTable(wsAnal, new[] { "Sprint", "Resource", "Analysis_h" });
        row = 2;
        foreach (int sprintIndex in sprintIndexes)
        {
            foreach (var resourceName in resources)
            {
                double analysis = SprintCapacities[sprintIndex].TryGetValue(resourceName, out var rc) ? rc.Analysis : 0.0;
                wsAnal.Cells[row, 1].Value = sprintIndex + SprintOffset;
                wsAnal.Cells[row, 2].Value = resourceName;
                wsAnal.Cells[row, 3].Value = analysis;
                row++;
            }
        }
        wsAnal.Cells.AutoFitColumns();

        WriteTable(wsUnder, new[] { "Sprint", "Resource", "Unused_h", "Reason" });
        row = 2;
        foreach (var entry in UnderutilizationEntries
                     .OrderBy(x => x.Sprint)
                     .ThenBy(x => x.Resource, StringComparer.OrdinalIgnoreCase))
        {
            wsUnder.Cells[row, 1].Value = entry.Sprint + SprintOffset;
            wsUnder.Cells[row, 2].Value = entry.Resource;
            wsUnder.Cells[row, 3].Value = entry.Unused;
            wsUnder.Cells[row, 4].Value = entry.Reason;
            row++;
        }
        wsUnder.Cells.AutoFitColumns();

        WriteTable(wsOver, new[] { "Resource", "Total_wish_pct", "Over_100pct", "Details" });
        row = 2;
        var wishesByResource = new Dictionary<string, List<double>>(StringComparer.OrdinalIgnoreCase);
        foreach (var epic in Epics)
        {
            foreach (var wish in epic.Wishes)
            {
                if (!wishesByResource.TryGetValue(wish.Resource, out var list))
                {
                    list = wishesByResource[wish.Resource] = new List<double>();
                }
                list.Add(wish.Percentage);
            }
        }

        foreach (var kv in wishesByResource.OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase))
        {
            double totalPct = kv.Value.Sum();
            string details = string.Join(
                "; ",
                Epics
                    .Where(e => e.Wishes.Any(w => w.Resource.Equals(kv.Key, StringComparison.OrdinalIgnoreCase)))
                    .Select(e =>
                    {
                        var wish = e.Wishes.First(w => w.Resource.Equals(kv.Key, StringComparison.OrdinalIgnoreCase));
                        return $"{e.Name}:{(int)(wish.Percentage * 100)}%";
                    }));
            wsOver.Cells[row, 1].Value = kv.Key;
            wsOver.Cells[row, 2].Value = Math.Round(totalPct * 100, 1);
            wsOver.Cells[row, 3].Value = totalPct > 1.0;
            wsOver.Cells[row, 4].Value = details;
            row++;
        }
        wsOver.Cells.AutoFitColumns();

        package.SaveAs(new FileInfo(_strOutputExcelFilePath));
    }

    public void ExportGanttSprintBased(
        string _strOutputPngPath,
        EnumPlanningMode _enumMode,
        bool _bOnlyDevelopmentEpics = false)
    {
        IEnumerable<Epic> epicSource = Epics
            .Where(e => e.StartDate.HasValue && e.EndDate.HasValue);

        if (_enumMode == EnumPlanningMode.Standard && _bOnlyDevelopmentEpics)
        {
            epicSource = epicSource.Where(e => e.IsInDevelopment);
        }

        var ranges = epicSource
            .Select(e =>
            {
                float start = SprintPosition(e.StartDate.Value, false);
                float end = SprintPosition(e.EndDate.Value, true);

                if (_enumMode == EnumPlanningMode.Standard)
                {
                    start = (float)Math.Floor(start);
                    if (start < 0f)
                    {
                        start = 0f;
                    }

                    end = (float)Math.Ceiling(end);
                    if (end < start + 1f)
                    {
                        end = start + 1f;
                    }
                }

                return new
                {
                    e.Name,
                    e.State,
                    StartPosition = start,
                    EndPosition = end,
                    Key = ExtractEpicKey(e.Name),
                    Hatched = !e.EndAnalysis.HasValue && e.State.Contains("analysis", StringComparison.OrdinalIgnoreCase)
                };
            })
            .OrderBy(r => r.StartPosition)
            .ThenBy(r => r.Key.Year)
            .ThenBy(r => r.Key.Num)
            .ToList();

        int rightLegendPad = 260;
        int topDateAxisPad = 80;
        int topTitlePad = 50;
        int bottomAxisPad = 80;

        int rows = Math.Max(1, ranges.Count);
        int rowHeight = 26;
        int plotHeight = rows * rowHeight + 40;
        int height = topTitlePad + topDateAxisPad + plotHeight + bottomAxisPad;
        int width = 1600 + rightLegendPad;

        using var bitmap = new SKBitmap(width, height);
        using var canvas = new SKCanvas(bitmap);
        canvas.Clear(SKColors.White);

        var gridPaintV = new SKPaint
        {
            Color = new SKColor(200, 200, 200),
            IsAntialias = true,
            StrokeWidth = 1,
            PathEffect = SKPathEffect.CreateDash(new float[] { 4, 4 }, 0)
        };
        var gridPaintH = new SKPaint
        {
            Color = new SKColor(220, 220, 220),
            IsAntialias = true,
            StrokeWidth = 1,
            PathEffect = SKPathEffect.CreateDash(new float[] { 4, 4 }, 0)
        };
        var titlePaint = new SKPaint { Color = SKColors.Black, TextSize = 20, IsAntialias = true };
        var labelPaint = new SKPaint { Color = SKColors.Black, TextSize = 13, IsAntialias = true };
        var axisPaint = new SKPaint { Color = SKColors.Black, TextSize = 12, IsAntialias = true };

        float maxLabelWidth = 0;
        foreach (var epic in Epics.Where(ep => ep.StartDate.HasValue && ep.EndDate.HasValue))
        {
            float widthMeasure = labelPaint.MeasureText(epic.Name);
            if (widthMeasure > maxLabelWidth)
            {
                maxLabelWidth = widthMeasure;
            }
        }

        int leftLabelPad = (int)(maxLabelWidth + 40);

        SKColor LIGHT_GREEN = SKColor.Parse("#77dd77");
        SKColor LIGHT_BLUE = SKColor.Parse("#89CFF0");
        SKColor BORDEAUX = SKColor.Parse("#800020");
        SKColor MAUVE = SKColors.Orchid;

        string title = _enumMode == EnumPlanningMode.Analysis
            ? "Gantt - Sprints (Analysis duration)"
            : "Gantt - Sprints (Realisation duration)";
        canvas.DrawText(title, leftLabelPad, 30, titlePaint);

        float maxPosition = ranges.Count > 0 ? ranges.Max(r => r.EndPosition) : 0f;
        int maxSprint = Math.Max(1, (int)Math.Ceiling(maxPosition));

        int plotLeft = leftLabelPad;
        int plotTop = topTitlePad + topDateAxisPad;
        int plotRight = width - rightLegendPad - 20;
        int plotBottom = height - bottomAxisPad;
        int plotWidth = plotRight - plotLeft;

        float xStep = plotWidth / Math.Max(1f, (float)maxSprint);
        for (int i = 0; i <= rows; i++)
        {
            float y = plotTop + i * rowHeight;
            canvas.DrawLine(plotLeft, y, plotRight, y, gridPaintH);
        }

        for (int s = 0; s <= maxSprint; s++)
        {
            float x = plotLeft + s * xStep;
            canvas.DrawLine(x, plotTop, x, plotBottom, gridPaintV);
        }

        void DrawHatchedRect(SKRect rect, SKColor baseColor)
        {
            SKColor fillColor = new(baseColor.Red, baseColor.Green, baseColor.Blue, 160);
            SKColor lineColor = new(baseColor.Red, baseColor.Green, baseColor.Blue, 220);

            using (var fillPaint = new SKPaint { Color = fillColor, IsAntialias = true, Style = SKPaintStyle.Fill })
            {
                canvas.DrawRect(rect, fillPaint);
            }

            using var hatchPaint = new SKPaint
            {
                Color = lineColor,
                IsAntialias = true,
                Style = SKPaintStyle.Stroke,
                StrokeWidth = 1.5f
            };

            float spacing = 6f;
            canvas.Save();
            canvas.ClipRect(rect);

            float maxOffset = rect.Width + rect.Height;
            for (float offset = -rect.Height; offset <= maxOffset; offset += spacing)
            {
                float startX = rect.Left + offset;
                float endX = startX - rect.Height;
                canvas.DrawLine(startX, rect.Top, endX, rect.Bottom, hatchPaint);
            }

            canvas.Restore();
        }

        for (int i = 0; i < ranges.Count; i++)
        {
            var r = ranges[i];
            float centerY = plotTop + i * rowHeight + rowHeight * 0.5f;
            float startX = plotLeft + r.StartPosition * xStep;
            float endX = plotLeft + r.EndPosition * xStep;
            float barHeight = Math.Min(18f, rowHeight - 6f);

            SKColor fill = r.State.Contains("pending") && r.State.Contains("analysis") ? BORDEAUX :
                           r.State.Contains("pending") && (r.State.Contains("develop") || r.State.Contains("dev")) ? LIGHT_BLUE :
                           r.State.Contains("analysis") && !r.State.Contains("pending") ? MAUVE :
                           r.State.Contains("develop") && !r.State.Contains("pending") ? LIGHT_GREEN :
                           SKColors.LightGray;

            var rect = new SKRect(startX, centerY - barHeight / 2f, endX, centerY + barHeight / 2f);
            if (r.Hatched)
            {
                DrawHatchedRect(rect, fill);
            }
            else
            {
                using var barPaint = new SKPaint { Color = fill, IsAntialias = true, Style = SKPaintStyle.Fill };
                canvas.DrawRect(rect, barPaint);
            }

            labelPaint.TextAlign = SKTextAlign.Right;
            canvas.DrawText(r.Name, plotLeft - 10, centerY + 5, labelPaint);
        }

        for (int s = 0; s <= maxSprint; s++)
        {
            float x = plotLeft + s * xStep;
            string text = (s + SprintOffset).ToString(CultureInfo.InvariantCulture);
            canvas.DrawText(text, x + 2, plotBottom + 20, axisPaint);
        }
        canvas.DrawText("Sprint number", (plotLeft + plotRight) / 2f - 40, plotBottom + 40, axisPaint);

        for (int s = 0; s <= maxSprint; s++)
        {
            float x = plotLeft + s * xStep;
            DateTime dt = InitialSprintDate.AddDays(s * SprintLengthDays);
            string txt = dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

            canvas.Save();
            canvas.RotateDegrees(-45, x, plotTop - 5);
            canvas.DrawText(txt, x, plotTop - 5, axisPaint);
            canvas.Restore();
        }
        canvas.DrawText("Sprint start date", (plotLeft + plotRight) / 2f - 50, topTitlePad + 20, axisPaint);

        float legendX = plotRight + 40;
        float legendY = topTitlePad + 20;
        float box = 16f;

        List<(string Label, SKColor Color, bool Hatched)> legendItems =
            _enumMode == EnumPlanningMode.Analysis
                ? new List<(string, SKColor, bool)>
                {
                    ("In Analysis", MAUVE, false),
                    ("Pending Analysis", BORDEAUX, false),
                    ("Analysis (no end date)", MAUVE, true)
                }
                : !_bOnlyDevelopmentEpics
                    ? new List<(string, SKColor, bool)>
                    {
                        ("In Development", LIGHT_GREEN, false),
                        ("In Analysis", MAUVE, false),
                        ("Pending Analysis", BORDEAUX, false),
                        ("Pending Development", LIGHT_BLUE, false),
                        ("Analysis (no end date)", MAUVE, true)
                    }
                    : new List<(string, SKColor, bool)>
                    {
                        ("In Development", LIGHT_GREEN, false)
                    };

        foreach (var item in legendItems)
        {
            var rect = new SKRect(legendX, legendY, legendX + box, legendY + box);
            if (item.Hatched)
            {
                DrawHatchedRect(rect, item.Color);
            }
            else
            {
                using var paint = new SKPaint { Color = item.Color, IsAntialias = true, Style = SKPaintStyle.Fill };
                canvas.DrawRect(rect, paint);
            }

            labelPaint.TextAlign = SKTextAlign.Left;
            canvas.DrawText(item.Label, legendX + box + 10, legendY + box - 3, labelPaint);
            legendY += 28;
        }

        using var image = SKImage.FromPixels(bitmap.PeekPixels());
        using var data = image.Encode(SKEncodedImageFormat.Png, 90);
        using var stream = File.OpenWrite(_strOutputPngPath);
        data.SaveTo(stream);
    }
}
