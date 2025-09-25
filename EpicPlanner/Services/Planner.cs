using Microsoft.Extensions.Logging;
using OfficeOpenXml;

namespace EpicPlanner
{
    internal class Planner
    {
        private readonly DateTime sprint0Start;
        private readonly int sprintDays;
        private const int MaxSprints = 300;
        private readonly AppConfiguration m_AppConfiguration;

        public Planner(AppConfiguration appConfiguration, ILoggerFactory loggerFactory)
        {
            m_AppConfiguration = appConfiguration ?? throw new ArgumentNullException(nameof(appConfiguration));
            this.sprint0Start = appConfiguration.PlannerConfiguration.Sprint0Start;
            this.sprintDays = appConfiguration.PlannerConfiguration.SprintDays;
        }

        public async Task RunAsync(string inputPath, string outputExcel, string outputPng)
        {
            using var package = new ExcelPackage(new FileInfo(inputPath));
            var wsEpics = package.Workbook.Worksheets["Planification des Epics"];
            var wsRes = package.Workbook.Worksheets["Resources par Sprint"];

            var resources = LoadResources(wsRes);                 // name -> dev hours per sprint (Heures Dév. Epic)
            var epics = LoadEpics(wsEpics, resources.Keys.ToList());

            // Adjust resources for absences (for each sprint)
            var absFetcher = new AbsenceFetcher(
                m_AppConfiguration.RedmineConfiguration.ServerUrl,
                m_AppConfiguration.RedmineConfiguration.ApiKey);
            Dictionary<int, Dictionary<string, double>> adjustedCapacities = await AdjustCapacitiesForAbsencesAsync(resources, absFetcher);

            var simulator = new Simulator(epics, adjustedCapacities, sprint0Start, sprintDays, m_AppConfiguration.PlannerConfiguration.MaxSprintCount);
            simulator.Run();

            simulator.ExportExcel(outputExcel);
            simulator.ExportGanttSprintBased(outputPng);
        }

        private Dictionary<string, double> LoadResources(ExcelWorksheet ws)
        {
            // We explicitly expect columns: [1] Ingénieur/Engineer, [2] Heures Dév. Epic
            int rows = ws.Dimension.End.Row;

            // Detect headers by names to be resilient
            var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int c = 1; c <= ws.Dimension.End.Column; c++)
            {
                var h = ws.Cells[1, c].GetValue<string>()?.Trim();
                if (!string.IsNullOrWhiteSpace(h)) headers[h] = c;
            }

            int nameCol = headers.ContainsKey("Ingénieur") ? headers["Ingénieur"]
                        : headers.ContainsKey("Engineer") ? headers["Engineer"] : 1;

            int hoursCol = headers.ContainsKey("Heures Dév. Epic")
                         ? headers["Heures Dév. Epic"]
                         : headers.FirstOrDefault(kv => kv.Key.Contains("Heures", StringComparison.OrdinalIgnoreCase) && kv.Key.Contains("Epic", StringComparison.OrdinalIgnoreCase)).Value;

            if (hoursCol == 0) hoursCol = 2; // fallback

            var dict = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            for (int r = 2; r <= rows; r++)
            {
                string name = ws.Cells[r, nameCol].GetValue<string>()?.Trim();
                double hours = ws.Cells[r, hoursCol].GetValue<double>();
                if (!string.IsNullOrWhiteSpace(name))
                    dict[name] = hours;
            }
            return dict;
        }

        private async Task<Dictionary<int, Dictionary<string, double>>> AdjustCapacitiesForAbsencesAsync(
            Dictionary<string, double> baseCapacities,
            AbsenceFetcher absFetcher)
        {
            Dictionary<int, Dictionary<string, double>> dict = new();
            List<(DateTime Start, DateTime End)> absences = await absFetcher.GetEngineersVacationsAsync();

            // Loop over sprints up to MaxSprints
            for (int sprint = 0; sprint < m_AppConfiguration.PlannerConfiguration.MaxSprintCount; sprint++)
            {
                var sprintStart = sprint0Start.AddDays(sprint * sprintDays);
                var sprintEnd = sprintStart.AddDays(sprintDays - 1);

                var cap = new Dictionary<string, double>(baseCapacities, StringComparer.OrdinalIgnoreCase);

                foreach (var user in baseCapacities.Keys)
                {
                    foreach (var (start, end) in absences)
                    {
                        var overlapStart = (start > sprintStart) ? start : sprintStart;
                        var overlapEnd = (end < sprintEnd) ? end : sprintEnd;

                        if (overlapEnd >= overlapStart)
                        {
                            int absentDays = (int)(overlapEnd - overlapStart).TotalDays + 1;
                            double dailyCap = baseCapacities[user] / sprintDays;
                            cap[user] -= dailyCap * absentDays;
                            if (cap[user] < 0) cap[user] = 0;
                        }
                    }
                }

                dict[sprint] = cap;
            }
            return dict;
        }

        private List<Epic> LoadEpics(OfficeOpenXml.ExcelWorksheet ws, List<string> resourceNames)
        {
            // Detect headers
            var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int c = 1; c <= ws.Dimension.End.Column; c++)
            {
                var h = ws.Cells[1, c].GetValue<string>()?.Trim();
                if (!string.IsNullOrWhiteSpace(h)) headers[h] = c;
            }

            int epicCol = headers.ContainsKey("Epic name") ? headers["Epic name"] : 1;
            int stateCol = headers.ContainsKey("State") ? headers["State"] : 2;
            int remainingCol = headers.FirstOrDefault(kv => kv.Key.Contains("Remaining", StringComparison.OrdinalIgnoreCase)).Value;
            int roughCol = headers.FirstOrDefault(kv => kv.Key.Contains("Rough", StringComparison.OrdinalIgnoreCase)).Value;
            int assignedCol = headers.FirstOrDefault(kv => kv.Key.Contains("Assigned to", StringComparison.OrdinalIgnoreCase)).Value;
            int willAssignCol = headers.FirstOrDefault(kv => kv.Key.Contains("Will be assigned", StringComparison.OrdinalIgnoreCase) || kv.Key.Contains("Will be assigne", StringComparison.OrdinalIgnoreCase)).Value;
            int depCol = headers.FirstOrDefault(kv => kv.Key.Contains("Epic dependency", StringComparison.OrdinalIgnoreCase) || kv.Key.Contains("Dependency", StringComparison.OrdinalIgnoreCase)).Value;
            int endAnalysisCol = headers.FirstOrDefault(kv => kv.Key.Contains("End of analysis", StringComparison.OrdinalIgnoreCase)).Value;

            int rows = ws.Dimension.End.Row;
            var epics = new List<Epic>();

            for (int r = 2; r <= rows; r++)
            {
                string epicName = ws.Cells[r, epicCol].GetValue<string>()?.Trim();
                if (string.IsNullOrWhiteSpace(epicName)) continue;

                string state = ws.Cells[r, stateCol].GetValue<string>() ?? "";
                double charge = 0.0;
                double remVal = (remainingCol > 0) ? ws.Cells[r, remainingCol].GetValue<double>() : 0.0;
                double roughVal = (roughCol > 0) ? ws.Cells[r, roughCol].GetValue<double>() : 0.0;

                if (remVal > 0)
                    charge = remVal;
                else if (roughVal > 0)
                    charge = roughVal;

                string assigned = assignedCol > 0 ? (ws.Cells[r, assignedCol].GetValue<string>() ?? "") : "";
                string willAssign = willAssignCol > 0 ? (ws.Cells[r, willAssignCol].GetValue<string>() ?? "") : "";
                string depRaw = depCol > 0 ? (ws.Cells[r, depCol].GetValue<string>() ?? "") : "";
                string endAnalysisStr = endAnalysisCol > 0 ? (ws.Cells[r, endAnalysisCol].GetValue<string>() ?? "") : "";

                DateTime? endAnalysis = null;
                if (DateTime.TryParse(endAnalysisStr, out var parsed))
                    endAnalysis = parsed;

                var epic = new Epic(epicName, state, charge, endAnalysis);
                epic.ParseAssignments(assigned, willAssign, resourceNames);
                epic.ParseDependencies(depRaw);

                // FIX #2: handle 0h epics
                if (epic.Charge <= 0)
                {
                    epic.Remaining = 0;
                    epic.StartDate = endAnalysis ?? this.sprint0Start;
                    epic.EndDate = endAnalysis ?? this.sprint0Start;
                }

                epics.Add(epic);
            }

            // Normalize dependency names to actual epic names (case-insensitive contains/equals)
            var epicNames = epics.Select(e => e.Name).ToList();
            foreach (var e in epics)
            {
                for (int i = 0; i < e.Dependencies.Count; i++)
                {
                    var d = e.Dependencies[i];
                    var match = epicNames.FirstOrDefault(x =>
                        x.Equals(d, StringComparison.OrdinalIgnoreCase) ||
                        x.IndexOf(d, StringComparison.OrdinalIgnoreCase) >= 0 ||
                        d.IndexOf(x, StringComparison.OrdinalIgnoreCase) >= 0);
                    if (!string.IsNullOrWhiteSpace(match))
                        e.Dependencies[i] = match;
                }
            }

            return epics;
        }
    }
}
