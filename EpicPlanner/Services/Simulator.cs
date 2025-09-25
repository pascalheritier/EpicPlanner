using OfficeOpenXml;
using SkiaSharp;
using System.Globalization;
using System.Text.RegularExpressions;

namespace EpicPlanner
{
    internal class Simulator
    {
        private readonly List<Epic> epics;
        private readonly Dictionary<int, Dictionary<string, double>> sprintCapacities;
        private readonly DateTime sprint0;
        private readonly int sprintDays;
        private readonly int maxSprintCount;

        private readonly Dictionary<string, DateTime> completedMap = new(StringComparer.OrdinalIgnoreCase);
        private readonly List<Allocation> allocations = new();
        private readonly List<(int Sprint, string Resource, double Unused, string Reason)> underutil = new();

        public Simulator(List<Epic> epics, Dictionary<int, Dictionary<string, double>> sprintCapacities, DateTime sprint0, int sprintDays, int maxSprintCount)
        {
            this.epics = epics;
            this.sprintCapacities = sprintCapacities;
            this.sprint0 = sprint0;
            this.sprintDays = sprintDays;
            this.maxSprintCount = maxSprintCount;

            // Mark 0h epics as completed in completedMap
            foreach (var e in epics.Where(e => e.Remaining <= 0))
                completedMap[e.Name] = e.EndDate ?? sprint0.AddDays(-1);
        }

        private bool DependencySatisfied(Epic e, DateTime sprintStart)
        {
            // If End of analysis is provided, cannot start before it (strict v4.1 rule).
            if (e.EndAnalysis.HasValue && sprintStart.Date < e.EndAnalysis.Value.Date)
                return false;

            if (e.Dependencies.Count == 0) return true;

            foreach (var d in e.Dependencies)
            {
                if (!completedMap.TryGetValue(d, out var depEnd)) return false;
                // FIX #1: allow start if depEnd < sprintStart
                if (depEnd.Date >= sprintStart.Date) return false;
            }
            return true;
        }

        public void Run()
        {
            // Strict assignment + spillover limited to assigned; priority In Development first, then others
            for (int sprint = 0; sprint < maxSprintCount; sprint++)
            {
                if (epics.All(e => e.Remaining <= 1e-6)) break;

                var sprintStart = sprint0.AddDays(sprint * sprintDays);
                var sprintEnd = sprintStart.AddDays(sprintDays - 1);

                // resource capacity left this sprint
                var resourceRemaining = new Dictionary<string, double>(sprintCapacities[sprint], StringComparer.OrdinalIgnoreCase);

                // Two passes: In Development first, Others second
                var activeDev = epics.Where(e => e.Remaining > 1e-6 && e.IsInDevelopment && DependencySatisfied(e, sprintStart)).ToList();
                var activeOthers = epics.Where(e => e.Remaining > 1e-6 && !e.IsInDevelopment && e.IsOtherAllowed && DependencySatisfied(e, sprintStart)).ToList();

                foreach (var activeSet in new[] { activeDev, activeOthers })
                {
                    // Requests: desired = remaining_capacity_of_resource * pct
                    var requests = new Dictionary<string, List<(Epic epic, double desired, double pct)>>(StringComparer.OrdinalIgnoreCase);
                    foreach (var e in activeSet)
                    {
                        foreach (var w in e.Wishes)
                        {
                            if (!resourceRemaining.ContainsKey(w.Resource) || w.Pct <= 0) continue;
                            double desired = resourceRemaining[w.Resource] * w.Pct;
                            if (!requests.TryGetValue(w.Resource, out var list))
                                list = requests[w.Resource] = new List<(Epic, double, double)>();
                            list.Add((e, desired, w.Pct));
                        }
                    }

                    // Allocate proportionally when oversubscribed
                    foreach (var kv in requests)
                    {
                        string rname = kv.Key;
                        var reqs = kv.Value;
                        if (reqs.Count == 0) continue;

                        double avail = resourceRemaining[rname];
                        double totalDesired = reqs.Sum(x => x.desired);
                        double scale = (totalDesired <= avail || totalDesired <= 1e-9) ? 1.0 : (avail / totalDesired);

                        foreach (var (epic, desired, pct) in reqs)
                        {
                            if (epic.Remaining <= 1e-6) continue;
                            double alloc = Math.Min(desired * scale, epic.Remaining);
                            if (alloc <= 1e-9) continue;

                            CommitAllocation(epic, sprint, rname, alloc, sprintStart);
                            resourceRemaining[rname] -= alloc;
                            if (epic.StartDate == null) epic.StartDate = sprintStart;
                            if (epic.Remaining <= 1e-6)
                            {
                                epic.Remaining = 0;
                                // End date = start of the sprint where the epic was finished
                                epic.EndDate = sprintStart;
                                completedMap[epic.Name] = epic.EndDate.Value;
                            }
                        }
                    }

                    // Spillover (only to epics where the resource is assigned)
                    foreach (var rname in resourceRemaining.Keys.ToList())
                    {
                        double leftover = resourceRemaining[rname];
                        if (leftover <= 1e-6) continue;

                        var candidates = activeSet.Where(e => e.Wishes.Any(w => w.Resource.Equals(rname, StringComparison.OrdinalIgnoreCase)) && e.Remaining > 1e-6).ToList();
                        if (candidates.Count == 0) continue;

                        var weights = candidates.Select(ep =>
                        {
                            var w = ep.Wishes.FirstOrDefault(xx => xx.Resource.Equals(rname, StringComparison.OrdinalIgnoreCase));
                            return (w != null && w.Pct > 0) ? w.Pct : 1.0;
                        }).ToList();

                        double sumW = weights.Sum();
                        if (sumW <= 1e-9) sumW = candidates.Count;

                        for (int i = 0; i < candidates.Count; i++)
                        {
                            if (leftover <= 1e-6) break;
                            var ep = candidates[i];
                            double share = leftover * (weights[i] / sumW);
                            double alloc = Math.Min(share, ep.Remaining);
                            if (alloc <= 1e-9) continue;

                            CommitAllocation(ep, sprint, rname, alloc, sprintStart);
                            leftover -= alloc;
                            resourceRemaining[rname] -= alloc;
                            if (ep.StartDate == null) ep.StartDate = sprintStart;
                            if (ep.Remaining <= 1e-6)
                            {
                                ep.Remaining = 0;
                                // FIX #3: EndDate = sprintStart
                                ep.EndDate = sprintStart;
                                completedMap[ep.Name] = ep.EndDate.Value;
                            }
                        }
                    }
                }

                // Underutilization (why capacity was left unconsumed)
                foreach (var kv in resourceRemaining)
                {
                    if (kv.Value > 1e-6)
                    {
                        string reason;
                        bool resourceHadAssignedEpics = epics.Any(ep => ep.Wishes.Any(w => w.Resource.Equals(kv.Key, StringComparison.OrdinalIgnoreCase)));
                        reason = resourceHadAssignedEpics ? "no remaining hours on assigned epics" : "no assigned epics";
                        underutil.Add((sprint, kv.Key, Math.Round(kv.Value, 2), reason));
                    }
                }
            }
        }

        private void CommitAllocation(Epic epic, int sprint, string resource, double hours, DateTime sprintStart)
        {
            epic.Remaining -= hours;
            var alloc = new Allocation(epic.Name, sprint, resource, hours, sprintStart);
            epic.History.Add(alloc);
            allocations.Add(alloc);
        }

        #region Exports (Excel + Gantt)
        public void ExportExcel(string path)
        {
            using var p = new ExcelPackage();
            var wsFinal = p.Workbook.Worksheets.Add("FinalSchedule");
            var wsA1 = p.Workbook.Worksheets.Add("AllocationsByEpicAndSprint");
            var wsA2 = p.Workbook.Worksheets.Add("AllocationsByEpicPerSprint");
            var wsVer = p.Workbook.Worksheets.Add("Verification");
            var wsPS = p.Workbook.Worksheets.Add("PerSprintSummary");
            var wsUnder = p.Workbook.Worksheets.Add("Underutilization");
            var wsOver = p.Workbook.Worksheets.Add("OverBooking");

            // FinalSchedule
            var finalRows = epics.Select(e => new
            {
                Epic = e.Name,
                State = e.State,
                Initial_Charge_h = e.Charge,
                Allocated_total_h = Math.Round(e.History.Sum(h => h.Hours), 2),
                Remaining_after_h = Math.Round(Math.Max(0, e.Remaining), 2),
                Start_date = e.StartDate?.ToString("yyyy-MM-dd") ?? "",
                End_date = e.EndDate?.ToString("yyyy-MM-dd") ?? ""
            }).ToList();

            // Sort by Start then numeric epic key (year, num)
            var finalSorted = finalRows
                .Select(r => new
                {
                    r.Epic,
                    r.State,
                    r.Initial_Charge_h,
                    r.Allocated_total_h,
                    r.Remaining_after_h,
                    r.Start_date,
                    r.End_date,
                    Start_dt = TryParseDate(r.Start_date),
                    Key = ExtractEpicKey(r.Epic)
                })
                .OrderBy(r => r.Start_dt)
                .ThenBy(r => r.Key.Year).ThenBy(r => r.Key.Num)
                .ToList();

            WriteTable(wsFinal, new[] { "Epic", "State", "Initial_Charge_h", "Allocated_total_h", "Remaining_after_h", "Start_date", "End_date" });
            int row = 2;
            foreach (var r in finalSorted)
            {
                wsFinal.Cells[row, 1].Value = r.Epic;
                wsFinal.Cells[row, 2].Value = r.State;
                wsFinal.Cells[row, 3].Value = r.Initial_Charge_h;
                wsFinal.Cells[row, 4].Value = r.Allocated_total_h;
                wsFinal.Cells[row, 5].Value = r.Remaining_after_h;
                wsFinal.Cells[row, 6].Value = r.Start_date;
                wsFinal.Cells[row, 7].Value = r.End_date;
                row++;
            }
            wsFinal.Cells.AutoFitColumns();

            // AllocationsByEpicAndSprint
            var aggA1 = allocations
                .GroupBy(a => new { a.Epic, a.Sprint, a.Resource })
                .Select(g => new { g.Key.Epic, g.Key.Sprint, g.Key.Resource, Hours = Math.Round(g.Sum(x => x.Hours), 2) })
                .OrderBy(x => x.Sprint).ThenBy(x => ExtractEpicKey(x.Epic).Year).ThenBy(x => ExtractEpicKey(x.Epic).Num).ThenBy(x => x.Resource, StringComparer.OrdinalIgnoreCase);
            WriteTable(wsA1, new[] { "Epic", "Sprint", "Resource", "Hours" });
            row = 2;
            foreach (var r in aggA1)
            {
                wsA1.Cells[row, 1].Value = r.Epic;
                wsA1.Cells[row, 2].Value = r.Sprint;
                wsA1.Cells[row, 3].Value = r.Resource;
                wsA1.Cells[row, 4].Value = r.Hours;
                row++;
            }
            wsA1.Cells.AutoFitColumns();

            // AllocationsByEpicPerSprint
            var aggA2 = allocations
                .GroupBy(a => new { a.Epic, a.Sprint, a.SprintStart })
                .Select(g => new { g.Key.Epic, g.Key.Sprint, g.Key.SprintStart, Total_Hours = Math.Round(g.Sum(x => x.Hours), 2) })
                .OrderBy(x => x.Sprint).ThenBy(x => ExtractEpicKey(x.Epic).Year).ThenBy(x => ExtractEpicKey(x.Epic).Num);
            WriteTable(wsA2, new[] { "Epic", "Sprint", "Sprint_start", "Total_Hours" });
            row = 2;
            foreach (var r in aggA2)
            {
                wsA2.Cells[row, 1].Value = r.Epic;
                wsA2.Cells[row, 2].Value = r.Sprint;
                wsA2.Cells[row, 3].Value = r.SprintStart.ToString("yyyy-MM-dd");
                wsA2.Cells[row, 4].Value = r.Total_Hours;
                row++;
            }
            wsA2.Cells.AutoFitColumns();

            // Verification
            WriteTable(wsVer, new[] { "Epic", "Initial_Charge_h", "Allocated_total_h", "Delta_h" });
            row = 2;
            foreach (var e in epics.OrderBy(x => x.StartDate ?? DateTime.MaxValue).ThenBy(x => ExtractEpicKey(x.Name).Year).ThenBy(x => ExtractEpicKey(x.Name).Num))
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

            // PerSprintSummary
            var sprints = allocations.Select(a => a.Sprint).Distinct().OrderBy(s => s).ToList();
            if (sprints.Count == 0) sprints.Add(0);
            var resList = sprintCapacities.Values
                .SelectMany(dict => dict.Keys)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(k => k, StringComparer.OrdinalIgnoreCase)
                .ToList();

            // Header
            var psHeaders = new List<string> { "Sprint", "Sprint_start", "Sprint_end" };
            foreach (var rname in resList)
            {
                psHeaders.Add($"{rname}_allocated_h");
                psHeaders.Add($"{rname}_capacity_h");
                psHeaders.Add($"{rname}_util_pct");
            }
            WriteTable(wsPS, psHeaders.ToArray());

            row = 2;
            foreach (int s in sprints)
            {
                DateTime start = SprintStartDate(s);
                DateTime end = start.AddDays(sprintDays - 1);
                int col = 1;
                wsPS.Cells[row, col++].Value = s;
                wsPS.Cells[row, col++].Value = start.ToString("yyyy-MM-dd");
                wsPS.Cells[row, col++].Value = end.ToString("yyyy-MM-dd");
                foreach (var rname in resList)
                {
                    double alloc = Math.Round(allocations.Where(a => a.Sprint == s && a.Resource.Equals(rname, StringComparison.OrdinalIgnoreCase)).Sum(a => a.Hours), 2);
                    double cap = sprintCapacities[s].TryGetValue(rname, out var hh) ? hh : 0.0;
                    double pct = (cap > 0) ? Math.Round(alloc / cap * 100.0, 2) : 0.0;
                    wsPS.Cells[row, col++].Value = alloc;
                    wsPS.Cells[row, col++].Value = cap;
                    wsPS.Cells[row, col++].Value = pct;
                }
                row++;
            }
            wsPS.Cells.AutoFitColumns();

            // Underutilization
            WriteTable(wsUnder, new[] { "Sprint", "Resource", "Unused_h", "Reason" });
            row = 2;
            foreach (var u in underutil.OrderBy(x => x.Sprint).ThenBy(x => x.Resource, StringComparer.OrdinalIgnoreCase))
            {
                wsUnder.Cells[row, 1].Value = u.Sprint;
                wsUnder.Cells[row, 2].Value = u.Resource;
                wsUnder.Cells[row, 3].Value = u.Unused;
                wsUnder.Cells[row, 4].Value = u.Reason;
                row++;
            }
            wsUnder.Cells.AutoFitColumns();

            // OverBooking (wish % sum by resource)
            var wishesByResource = new Dictionary<string, List<double>>(StringComparer.OrdinalIgnoreCase);
            foreach (var e in epics)
            {
                foreach (var w in e.Wishes)
                {
                    if (!wishesByResource.TryGetValue(w.Resource, out var list))
                        list = wishesByResource[w.Resource] = new List<double>();
                    list.Add(w.Pct);
                }
            }
            var wsOverHeaders = new[] { "Resource", "Total_wish_pct", "Over_100pct", "Details" };
            WriteTable(wsOver, wsOverHeaders);
            row = 2;
            foreach (var kv in wishesByResource.OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase))
            {
                double totalPct = kv.Value.Sum();
                string details = string.Join("; ", epics.Where(e => e.Wishes.Any(w => w.Resource.Equals(kv.Key, StringComparison.OrdinalIgnoreCase)))
                    .Select(e => $"{e.Name}:{(int)(e.Wishes.First(w => w.Resource.Equals(kv.Key, StringComparison.OrdinalIgnoreCase)).Pct * 100)}%"));
                wsOver.Cells[row, 1].Value = kv.Key;
                wsOver.Cells[row, 2].Value = Math.Round(totalPct * 100, 1);
                wsOver.Cells[row, 3].Value = totalPct > 1.0;
                wsOver.Cells[row, 4].Value = details;
                row++;
            }
            wsOver.Cells.AutoFitColumns();

            // Save
            p.SaveAs(new FileInfo(path));
        }

        private static void WriteTable(ExcelWorksheet ws, string[] headers)
        {
            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cells[1, i + 1].Value = headers[i];
                ws.Cells[1, i + 1].Style.Font.Bold = true;
            }
        }

        public void ExportGanttSprintBased(string pngPath)
        {
            // Build ranges per epic from allocations
            var ranges = epics
                .Where(e => e.StartDate.HasValue && e.EndDate.HasValue)
                .Select(e =>
                {
                    int s0 = SprintIndex(e.StartDate.Value);
                    int s1 = SprintIndex(e.EndDate.Value);
                    var key = ExtractEpicKey(e.Name);
                    return new { e.Name, e.State, SprintStart = s0, SprintEnd = s1, Key = key };
                })
                .OrderBy(r => r.SprintStart)
                .ThenBy(r => r.Key.Year).ThenBy(r => r.Key.Num)
                .ToList();

            // Visual constants

            //int leftLabelPad = 420;                 // space reserved at left for epic names
            int rightLegendPad = 260;               // legend area
            int topDateAxisPad = 80;                // top dates
            int topTitlePad = 50;
            int bottomAxisPad = 80;

            int rows = Math.Max(1, ranges.Count);
            int rowHeight = 26;                     // compact bars
            int plotHeight = rows * rowHeight + 40;
            int height = topTitlePad + topDateAxisPad + plotHeight + bottomAxisPad;
            int width = 1600 + rightLegendPad;

            using var bmp = new SKBitmap(width, height);
            using var canvas = new SKCanvas(bmp);
            canvas.Clear(SKColors.White);

            // Paints
            var black = new SKPaint { Color = SKColors.Black, IsAntialias = true };
            var grey = new SKPaint { Color = new SKColor(230, 230, 230), IsAntialias = true };
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

            // Measure max epic title width
            float maxLabelWidth = 0;
            foreach (var e in epics.Where(ep => ep.StartDate.HasValue && ep.EndDate.HasValue))
            {
                float w = labelPaint.MeasureText(e.Name);
                if (w > maxLabelWidth) maxLabelWidth = w;
            }

            int leftLabelPad = (int)(maxLabelWidth + 40); // space reserved at left for epic names, dynamic padding with margin

            // Colors
            SKColor LIGHT_GREEN = SKColor.Parse("#77dd77");
            SKColor LIGHT_BLUE = SKColor.Parse("#89CFF0");
            SKColor BORDEAUX = SKColor.Parse("#800020");
            SKColor MAUVE = SKColors.Orchid;
            SKColor BAR_BORDER = new SKColor(0, 0, 0, 0); // no border as requested

            // Title
            canvas.DrawText("Gantt - Sprint-based (v5, strict assignment)", leftLabelPad, 30, titlePaint);

            // Determine sprint range
            int maxSprint = ranges.Count > 0 ? ranges.Max(r => r.SprintEnd) + 1 : 1;

            // Plot area
            var plotLeft = leftLabelPad;
            var plotTop = topTitlePad + topDateAxisPad;
            var plotRight = width - rightLegendPad - 20;
            var plotBottom = height - bottomAxisPad;
            var plotWidth = plotRight - plotLeft;
            var plotHeightPx = plotBottom - plotTop;

            // X-scale: one unit per sprint
            float xStep = plotWidth / Math.Max(1f, (float)maxSprint);
            // Y positions
            // Row i: y = plotTop + i*rowHeight + rowHeight/2
            // Also draw horizontal grid lines per epic row
            for (int i = 0; i <= rows; i++)
            {
                float y = plotTop + i * rowHeight;
                canvas.DrawLine(plotLeft, y, plotRight, y, gridPaintH);
            }

            // Vertical grid at each sprint boundary (integer)
            for (int s = 0; s <= maxSprint; s++)
            {
                float x = plotLeft + s * xStep;
                canvas.DrawLine(x, plotTop, x, plotBottom, gridPaintV);
            }

            // Bars and left labels
            for (int i = 0; i < ranges.Count; i++)
            {
                var r = ranges[i];
                int rowYIndex = i;
                float cy = plotTop + rowYIndex * rowHeight + rowHeight * 0.5f;
                float xs = plotLeft + r.SprintStart * xStep;
                float xe = plotLeft + (r.SprintEnd + 1) * xStep; // inclusive end
                float barH = Math.Min(18f, rowHeight - 6f);

                SKColor fill = r.State.Contains("pending") && r.State.Contains("analysis") ? BORDEAUX :
                               r.State.Contains("pending") && (r.State.Contains("develop") || r.State.Contains("dev")) ? LIGHT_BLUE :
                               r.State.Contains("analysis") && !r.State.Contains("pending") ? MAUVE :
                               r.State.Contains("develop") && !r.State.Contains("pending") ? LIGHT_GREEN :
                               SKColors.LightGray;

                using var barPaint = new SKPaint { Color = fill, IsAntialias = true, Style = SKPaintStyle.Fill };
                var rect = new SKRect(xs, cy - barH / 2f, xe, cy + barH / 2f);
                canvas.DrawRect(rect, barPaint);

                // Epic label to the left (outside)
                labelPaint.TextAlign = SKTextAlign.Right;
                canvas.DrawText(r.Name, plotLeft - 10, cy + 5, labelPaint);
            }

            // Bottom axis: sprint numbers (centered at each sprint band)
            for (int s = 0; s <= maxSprint; s++)
            {
                float x = plotLeft + s * xStep;
                string txt = s.ToString(CultureInfo.InvariantCulture);
                canvas.DrawText(txt, x + 2, plotBottom + 20, axisPaint); // small offset
            }
            canvas.DrawText("Sprint number", (plotLeft + plotRight) / 2f - 40, plotBottom + 40, axisPaint);

            // Top axis: sprint start dates, left-anchored on the vertical line (exactly as requested)
            for (int s = 0; s <= maxSprint; s++)
            {
                float x = plotLeft + s * xStep;
                var dt = sprint0.AddDays(s * sprintDays);
                string txt = dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

                canvas.Save();
                canvas.RotateDegrees(-45, x, plotTop - 5);   // anchor at the grid top
                canvas.DrawText(txt, x, plotTop - 5, axisPaint);
                canvas.Restore();
            }
            canvas.DrawText("Sprint start date", (plotLeft + plotRight) / 2f - 50, topTitlePad + 20, axisPaint);

            // Legend on the right
            float lgX = plotRight + 40;
            float lgY = topTitlePad + 20;
            float box = 16f;

            var legendItems = new List<(string Label, SKColor Color)>
            {
                ("In Development", LIGHT_GREEN),
                ("In Analysis", MAUVE),
                ("Pending Analysis", BORDEAUX),
                ("Pending Development", LIGHT_BLUE)
            };

            foreach (var item in legendItems)
            {
                using var paint = new SKPaint { Color = item.Color, IsAntialias = true, Style = SKPaintStyle.Fill };
                var rect = new SKRect(lgX, lgY, lgX + box, lgY + box);
                canvas.DrawRect(rect, paint);

                labelPaint.TextAlign = SKTextAlign.Left;
                canvas.DrawText(item.Label, lgX + box + 10, lgY + box - 3, labelPaint);

                lgY += 28; // spacing
            }

            // Save PNG
            using var image = SKImage.FromPixels(bmp.PeekPixels());
            using var data = image.Encode(SKEncodedImageFormat.Png, 90);
            using var fs = File.OpenWrite(pngPath);
            data.SaveTo(fs);
        }
        #endregion

        #region Helpers
        private DateTime TryParseDate(string s)
        {
            if (DateTime.TryParse(s, out var d)) return d;
            return DateTime.MaxValue;
        }

        private static (int Year, int Num) ExtractEpicKey(string epic)
        {
            if (string.IsNullOrWhiteSpace(epic)) return (9999, 9999);
            var m = Regex.Match(epic, @"(\d{4})[-_ ]+(\d{1,3})");
            if (m.Success) return (int.Parse(m.Groups[1].Value), int.Parse(m.Groups[2].Value));
            var m2 = Regex.Match(epic, @"(\d{4})");
            if (m2.Success) return (int.Parse(m2.Groups[1].Value), 0);
            return (9999, 9999);
        }

        private int SprintIndex(DateTime dt) => (int)((dt.Date - sprint0.Date).TotalDays / sprintDays);
        private DateTime SprintStartDate(int sprint) => sprint0.AddDays(sprint * sprintDays);
        #endregion
    }
}
