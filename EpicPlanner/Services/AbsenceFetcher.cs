using Redmine.Net.Api;
using Redmine.Net.Api.Types;
using System.Collections.Specialized;

namespace EpicPlanner
{
    internal class AbsenceFetcher
    {
        private RedmineManager _manager;

        public AbsenceFetcher(string baseUrl, string apiKey)
        {
            _manager = new RedmineManager(baseUrl, apiKey);
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
        public List<(DateTime Start, DateTime End)> GetVacationsForUserAsync(string userName)
        {
            var result = new List<(DateTime, DateTime)>();
            var parameters = new NameValueCollection
            {
                { RedmineKeys.SUBJECT, "Vacances" },
                { RedmineKeys.ASSIGNED_TO_ID, GetUserId(userName).ToString() },
                { RedmineKeys.STATUS_ID, "Accepted" }
            };

            if (!TryGetIssues(parameters, out IEnumerable<Issue> issues))
                return result;

            foreach (var issue in issues)
            {
                if (issue.StartDate.HasValue && issue.DueDate.HasValue)
                {
                    result.Add((issue.StartDate.Value, issue.DueDate.Value));
                }
            }
            return result;
        }

        private bool TryGetIssues(NameValueCollection parameters, out IEnumerable<Issue>? foundIssues)
        {
            foundIssues = null;
            try
            {
                foundIssues = _manager.GetObjects<Issue>(parameters);
                if (foundIssues is not null)
                    return true;
            }
            catch (System.Exception e)
            {
                // silent failure, found issue is null
            }
            return false;
        }
    }
}
