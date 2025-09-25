using Redmine.Net.Api;
using Redmine.Net.Api.Net;
using Redmine.Net.Api.Types;
using System.Collections.Specialized;

namespace EpicPlanner;

internal class AbsenceFetcher
{
    #region Members

    private readonly RedmineManager m_RedmineManager;

    #endregion

    #region Constructor

    public AbsenceFetcher(string _strBaseUrl, string _strApiKey)
    {
        m_RedmineManager = new RedmineManager(new RedmineManagerOptionsBuilder()
            .WithHost(_strBaseUrl)
            .WithApiKeyAuthentication(_strApiKey));
    }

    #endregion

    #region Absences fetching

    /// <summary>
    /// Get all accepted vacation absences for all resources
    /// </summary>
    public async Task<Dictionary<string, List<(DateTime Start, DateTime End)>>> GetResourcesAbsencesAsync()
    {
        Dictionary<string, List<(DateTime, DateTime)>> result = new();
        var parameters = new NameValueCollection
            {
                { RedmineKeys.TRACKER_ID, "7" },
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

    private async Task<IEnumerable<Issue>> GetIssuesAsync(NameValueCollection _Parameters)
    {
        try
        {
            RequestOptions requestOptions = new RequestOptions
            {
                QueryString = _Parameters
            };
            IEnumerable<Issue> foundIssues = await m_RedmineManager.GetAsync<Issue>(requestOptions);
            if (foundIssues is not null)
                return foundIssues;
        }
        catch (Exception e)
        {
            // silent failure, found issue is null
        }
        return Enumerable.Empty<Issue>();
    } 

    #endregion
}