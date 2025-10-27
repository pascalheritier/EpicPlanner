using Redmine.Net.Api;
using Redmine.Net.Api.Net;
using Redmine.Net.Api.Types;
using System.Collections.Specialized;

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
        var parameters = new NameValueCollection
        {
            { RedmineKeys.TRACKER_ID, "6" }, // Tracker ID for TODO
            { RedmineKeys.FIXED_VERSION_ID, (184 + _iSprintNumber).ToString() } // Sprint version IDs start at 185 for Sprint 1
        };

        IEnumerable<Issue> issues = await GetIssuesAsync(parameters);
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
