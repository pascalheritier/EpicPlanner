using Redmine.Net.Api;
using Redmine.Net.Api.Async;
using Redmine.Net.Api.Net;
using Redmine.Net.Api.Types;
using System.Collections.Specialized;

namespace EpicPlanner
{
    internal class AbsenceFetcher
    {
        private RedmineManager _manager;

        public AbsenceFetcher(string baseUrl, string apiKey)
        {
            _manager = new RedmineManager(new RedmineManagerOptionsBuilder()
                .WithHost(baseUrl)
                .WithApiKeyAuthentication(apiKey));
        }

        public int GetUserId(string userName)
        {
            var parameters = new NameValueCollection
            {
                { RedmineKeys.LOGIN, userName }
            };
            throw new NotImplementedException();
        }

        /// <summary>
        /// Get all accepted vacation absences for a given user
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

        private async Task<IEnumerable<Issue>> GetIssuesAsync(NameValueCollection parameters)
        {
            try
            {
                RequestOptions requestOptions = new RequestOptions
                {
                    QueryString = parameters
                };
                IEnumerable<Issue> foundIssues = await _manager.GetAsync<Issue>(requestOptions);
                if (foundIssues is not null)
                    return foundIssues;
            }
            catch (System.Exception e)
            {
                // silent failure, found issue is null
            }
            return Enumerable.Empty<Issue>();
        }
    }
}
