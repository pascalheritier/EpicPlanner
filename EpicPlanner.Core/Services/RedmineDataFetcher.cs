using Redmine.Net.Api;
using Redmine.Net.Api.Net;
using Redmine.Net.Api.Types;
using System.Collections.Specialized;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace EpicPlanner.Core;

public class RedmineDataFetcher
{
    #region Members

    private readonly RedmineManager m_RedmineManager;

    #endregion

    #region Constructor

    public RedmineDataFetcher(string _strBaseUrl, string _strApiKey)
    {
        m_RedmineManager = new RedmineManager(new RedmineManagerOptionsBuilder()
            .WithHost(_strBaseUrl)
            .WithApiKeyAuthentication(_strApiKey));
    }

    #endregion

    #region Redmine data fetching

    /// <summary>
    /// Get all accepted vacation absences for all resources
    /// </summary>
    public async Task<Dictionary<string, List<(DateTime Start, DateTime End)>>> GetResourcesAbsencesAsync()
    {
        Dictionary<string, List<(DateTime, DateTime)>> result = new();
        var parameters = new NameValueCollection
        {
            { RedmineKeys.TRACKER_ID, "7" }, // Tracker ID for Absence
            { RedmineKeys.STATUS_ID, "9" } // Accepted state
        };

        IEnumerable<Issue> issues = await GetIssuesAsync(parameters);
        foreach (var issue in issues)
        {
            if (issue.StartDate.HasValue && issue.DueDate.HasValue)
            {
                if (!result.ContainsKey(issue.Author.Name))
                    result[issue.Author.Name] = new();
                result[issue.Author.Name].Add((issue.StartDate.Value, issue.DueDate.Value));
            }
        }
        return result;
    }

    public async Task<Dictionary<string, double>> GetPlannedHoursForSprintAsync(int _iSprintNumber, DateTime _SprintStart, DateTime _SprintEnd)
    {
        var result = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        IEnumerable<Issue> issues = await GetSprintIssuesAsync(_iSprintNumber);
        foreach (var issue in issues)
        {
            if (issue.AssignedTo == null || issue.Subject.Contains("[Suivi]") || issue.Subject.Contains("[Analyse]"))
                continue;

            IssueCustomField? estimation = issue.CustomFields.FirstOrDefault(_Field => _Field.Name == "Reste à faire");
            if (estimation?.Value == null)
                continue;

            string user = issue.AssignedTo.Name;
            double.TryParse(estimation.Values.Select(_V => _V.Info).First(), out double hours);
            if (!result.ContainsKey(user)) result[user] = 0;
            result[user] += hours;
        }

        return result;
    }

    public async Task<List<SprintEpicSummary>> GetEpicSprintSummariesAsync(int _iSprintNumber, DateTime _SprintStart, DateTime _SprintEnd)
    {
        var summaries = new Dictionary<string, SprintEpicSummary>(StringComparer.OrdinalIgnoreCase);
        IEnumerable<Issue> issues = await GetSprintIssuesAsync(_iSprintNumber);

        foreach (var issue in issues)
        {
            if (issue.Subject.Contains("[Suivi]") || issue.Subject.Contains("[Analyse]"))
                continue;

            string epicName = ExtractEpicName(issue);
            if (!summaries.TryGetValue(epicName, out var summary))
            {
                summary = new SprintEpicSummary { Epic = epicName };
                summaries[epicName] = summary;
            }

            double planned = issue.EstimatedHours ?? 0.0;
            double consumed = issue.SpentHours ?? 0.0;
            double remaining = ExtractRemaining(issue);

            if (consumed <= 0 && planned > 0 && remaining >= 0)
            {
                double fallback = planned - remaining;
                if (fallback > consumed)
                    consumed = fallback;
            }

            summary.PlannedCapacity += planned;
            summary.Consumed += consumed;
            summary.Remaining += remaining;
        }

        return summaries.Values
            .OrderBy(s => s.Epic, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private async Task<IEnumerable<Issue>> GetSprintIssuesAsync(int _iSprintNumber)
    {
        var parameters = new NameValueCollection
        {
            { RedmineKeys.TRACKER_ID, "6" }, // Tracker ID for TODO
            { RedmineKeys.FIXED_VERSION_ID, (184 + _iSprintNumber).ToString() } // Sprint version IDs start at 185 for Sprint 1
        };

        return await GetIssuesAsync(parameters);
    }

    private static string ExtractEpicName(Issue _Issue)
    {
        if (_Issue.CustomFields != null)
        {
            var epicField = _Issue.CustomFields.FirstOrDefault(cf =>
                cf.Name.Equals("Epic", StringComparison.OrdinalIgnoreCase) ||
                cf.Name.Equals("Epic name", StringComparison.OrdinalIgnoreCase) ||
                cf.Name.IndexOf("epic", StringComparison.OrdinalIgnoreCase) >= 0);

            if (epicField != null)
            {
                string? value = null;
                if (epicField.Values != null)
                    value = epicField.Values.Select(v => v.Info ?? v.Value).FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));

                if (string.IsNullOrWhiteSpace(value))
                    value = epicField.Value;

                if (!string.IsNullOrWhiteSpace(value))
                    return value.Trim();
            }
        }

        if (!string.IsNullOrWhiteSpace(_Issue.Subject))
        {
            var match = Regex.Match(_Issue.Subject, @"\[(?<epic>[^\]]+)\]");
            if (match.Success)
                return match.Groups["epic"].Value.Trim();
        }

        return "(No Epic)";
    }

    private static double ExtractRemaining(Issue _Issue)
    {
        IssueCustomField? estimation = _Issue.CustomFields?.FirstOrDefault(_Field => _Field.Name == "Reste à faire");
        if (estimation == null)
            return 0.0;

        if (!string.IsNullOrWhiteSpace(estimation.Value) &&
            double.TryParse(estimation.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double direct))
            return direct;

        if (estimation.Values != null)
        {
            foreach (var val in estimation.Values)
            {
                string? raw = val.Info ?? val.Value;
                if (string.IsNullOrWhiteSpace(raw)) continue;
                if (double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsed))
                    return parsed;
            }
        }

        return 0.0;
    }

    private async Task<IEnumerable<Issue>> GetIssuesAsync(NameValueCollection _Parameters)
    {
        try
        {
            RequestOptions requestOptions = new RequestOptions
            {
                QueryString = _Parameters
            };
            IEnumerable<Issue>? foundIssues = await m_RedmineManager.GetAsync<Issue>(requestOptions);
            if (foundIssues is not null)
                return foundIssues;
        }
        catch
        {
            // Swallow exception and return empty sequence to keep planner working when Redmine is unreachable.
        }
        return Enumerable.Empty<Issue>();
    } 

    #endregion
}
