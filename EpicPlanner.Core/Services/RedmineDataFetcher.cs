using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Redmine.Net.Api;
using Redmine.Net.Api.Net;
using Redmine.Net.Api.Types;

namespace EpicPlanner.Core;

public class RedmineDataFetcher
{
    #region Members

    private const int EpicCustomFieldId = 57;

    private readonly RedmineManager m_RedmineManager;
    private readonly HttpClient m_HttpClient;
    private readonly SemaphoreSlim m_EpicEnumerationLock = new(1, 1);
    private Dictionary<string, string>? m_EpicEnumerationCache;

    private static readonly Regex s_EnumerationRowRegex = new(
        @"custom_field_enumeration_(?<id>\d+)[^>]*>(?<content>.*?)</tr>",
        RegexOptions.IgnoreCase | RegexOptions.Singleline);

    private static readonly Regex s_EnumerationNameCellRegex = new(
        @"<td[^>]*class\s*=\s*""name""[^>]*>(?<name>.*?)</td>",
        RegexOptions.IgnoreCase | RegexOptions.Singleline);

    private static readonly Regex s_HtmlTagRegex = new("<[^>]+>", RegexOptions.Singleline);

    #endregion

    #region Constructor

    public RedmineDataFetcher(string _strBaseUrl, string _strApiKey)
    {
        if (string.IsNullOrWhiteSpace(_strBaseUrl))
            throw new ArgumentException("Redmine base URL must be provided.", nameof(_strBaseUrl));

        m_RedmineManager = new RedmineManager(new RedmineManagerOptionsBuilder()
            .WithHost(_strBaseUrl)
            .WithApiKeyAuthentication(_strApiKey));

        string normalizedBaseUrl = _strBaseUrl.EndsWith("/") ? _strBaseUrl : _strBaseUrl + "/";

        m_HttpClient = new HttpClient
        {
            BaseAddress = new Uri(normalizedBaseUrl)
        };
        m_HttpClient.DefaultRequestHeaders.Add("X-Redmine-API-Key", _strApiKey);
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
            if (estimation?.Values == null)
                continue;

            string? rawHours = estimation.Values.Select(_V => _V.Info).FirstOrDefault(_V => !string.IsNullOrWhiteSpace(_V));
            if (string.IsNullOrWhiteSpace(rawHours))
                continue;

            string user = issue.AssignedTo.Name;
            double.TryParse(rawHours, NumberStyles.Any, CultureInfo.InvariantCulture, out double hours);
            if (!result.ContainsKey(user)) result[user] = 0;
            result[user] += hours;
        }

        return result;
    }

    public async Task<List<SprintEpicSummary>> GetEpicSprintSummariesAsync(int _iSprintNumber, DateTime _SprintStart, DateTime _SprintEnd)
    {
        var summaries = new Dictionary<string, SprintEpicSummary>(StringComparer.OrdinalIgnoreCase);
        IEnumerable<Issue> issues = await GetSprintIssuesAsync(_iSprintNumber);

        var parentCache = new Dictionary<int, Issue>();

        foreach (var issue in issues)
        {
            if (issue.Subject.Contains("[Suivi]") || issue.Subject.Contains("[Analyse]"))
                continue;

            string epicName = await ResolveEpicNameAsync(issue, parentCache);
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

    private async Task<string> ResolveEpicNameAsync(
        Issue _Issue,
        Dictionary<int, Issue> _ParentCache)
    {
        string? direct = NormalizeEpicName(await ExtractEpicFromCustomFieldAsync(_Issue));
        if (!string.IsNullOrWhiteSpace(direct))
            return direct;

        Issue? parent = await GetParentIssueAsync(_Issue, _ParentCache);
        if (parent != null)
        {
            string? parentValue = NormalizeEpicName(await ExtractEpicFromCustomFieldAsync(parent));
            if (!string.IsNullOrWhiteSpace(parentValue))
                return parentValue;
        }

        string? fromSubject = NormalizeEpicName(ExtractEpicFromSubject(_Issue));
        if (!string.IsNullOrWhiteSpace(fromSubject))
            return fromSubject;

        if (parent != null)
        {
            string? parentSubject = NormalizeEpicName(ExtractEpicFromSubject(parent));
            if (!string.IsNullOrWhiteSpace(parentSubject))
                return parentSubject;
        }

        return "(No Epic)";
    }

    private string? NormalizeEpicName(string? _RawEpicValue)
    {
        if (string.IsNullOrWhiteSpace(_RawEpicValue))
            return null;

        return _RawEpicValue.Trim();
    }

    private async Task<string?> ExtractEpicFromCustomFieldAsync(Issue _Issue)
    {
        if (_Issue.CustomFields == null)
            return null;

        IssueCustomField? epicField = _Issue.CustomFields
            .FirstOrDefault(cf => cf.Id == EpicCustomFieldId);

        epicField ??= _Issue.CustomFields.FirstOrDefault(cf =>
            cf.Name.Equals("Epic", StringComparison.OrdinalIgnoreCase) ||
            cf.Name.Equals("Epic name", StringComparison.OrdinalIgnoreCase) ||
            cf.Name.IndexOf("epic", StringComparison.OrdinalIgnoreCase) >= 0);

        if (epicField == null)
            return null;

        if (epicField.Values == null)
            return null;

        string? rawValue = epicField.Values
            .Select(v => v.Info)
            .FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));

        if (string.IsNullOrWhiteSpace(rawValue))
            return null;

        string trimmedValue = rawValue.Trim();

        Dictionary<string, string> lookup = await GetEpicEnumerationLookupAsync();
        if (lookup.TryGetValue(trimmedValue, out string? mappedValue))
            return string.IsNullOrWhiteSpace(mappedValue) ? null : mappedValue.Trim();

        return trimmedValue;
    }

    private static string? ExtractEpicFromSubject(Issue _Issue)
    {
        if (string.IsNullOrWhiteSpace(_Issue.Subject))
            return null;

        var match = Regex.Match(_Issue.Subject, @"\[(?<epic>[^\]]+)\]");
        if (!match.Success)
            return null;

        string value = match.Groups["epic"].Value.Trim();
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }

    private async Task<Issue?> GetParentIssueAsync(Issue _Issue, Dictionary<int, Issue> _ParentCache)
    {
        int? parentId = _Issue.ParentIssue?.Id;
        if (!parentId.HasValue)
            return null;

        if (_ParentCache.TryGetValue(parentId.Value, out Issue? cached))
            return cached;

        Issue? parent = await GetIssueByIdAsync(parentId.Value);
        if (parent != null)
            _ParentCache[parentId.Value] = parent;

        return parent;
    }

    private async Task<Issue?> GetIssueByIdAsync(int _IssueId)
    {
        var parameters = new NameValueCollection
        {
            { RedmineKeys.ISSUE_ID, _IssueId.ToString(CultureInfo.InvariantCulture) }
        };

        return (await GetIssuesAsync(parameters)).FirstOrDefault();
    }

    private static double ExtractRemaining(Issue _Issue)
    {
        IssueCustomField? estimation = _Issue.CustomFields?.FirstOrDefault(_Field => _Field.Name == "Reste à faire");
        if (estimation == null)
            return 0.0;

        if (estimation.Values != null)
        {
            foreach (var val in estimation.Values)
            {
                string? raw = val.Info;
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

    private async Task<Dictionary<string, string>> GetEpicEnumerationLookupAsync()
    {
        if (m_EpicEnumerationCache != null)
            return m_EpicEnumerationCache;

        await m_EpicEnumerationLock.WaitAsync().ConfigureAwait(false);
        try
        {
            if (m_EpicEnumerationCache != null)
                return m_EpicEnumerationCache;

            try
            {
                Dictionary<string, string>? map = await TryFetchEpicEnumerationsAsync().ConfigureAwait(false);
                if (map != null)
                {
                    m_EpicEnumerationCache = map;
                    return m_EpicEnumerationCache;
                }
            }
            catch
            {
                // Swallow exceptions so missing enumeration data does not block planner execution.
            }

            m_EpicEnumerationCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            return m_EpicEnumerationCache;
        }
        finally
        {
            m_EpicEnumerationLock.Release();
        }
    }

    private async Task<Dictionary<string, string>?> TryFetchEpicEnumerationsAsync()
    {
        Dictionary<string, string>? map = await TryFetchEpicEnumerationsFromJsonAsync().ConfigureAwait(false);
        if (map != null && map.Count > 0)
            return map;

        map = await TryFetchEpicEnumerationsFromHtmlAsync().ConfigureAwait(false);
        if (map != null && map.Count > 0)
            return map;

        return map;
    }

    private async Task<Dictionary<string, string>?> TryFetchEpicEnumerationsFromJsonAsync()
    {
        try
        {
            using HttpResponseMessage response = await m_HttpClient
                .GetAsync($"custom_fields/{EpicCustomFieldId}/enumerations.json")
                .ConfigureAwait(false);

            if (!response.IsSuccessStatusCode)
                return null;

            await using var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
            CustomFieldEnumerationResponse? payload = await JsonSerializer
                .DeserializeAsync<CustomFieldEnumerationResponse>(stream)
                .ConfigureAwait(false);

            if (payload?.CustomFieldEnumerations == null || payload.CustomFieldEnumerations.Count == 0)
                return null;

            return ConvertEnumerationsToLookup(payload.CustomFieldEnumerations);
        }
        catch
        {
            return null;
        }
    }

    private async Task<Dictionary<string, string>?> TryFetchEpicEnumerationsFromHtmlAsync()
    {
        try
        {
            using HttpResponseMessage response = await m_HttpClient
                .GetAsync($"custom_fields/{EpicCustomFieldId}/enumerations")
                .ConfigureAwait(false);

            if (!response.IsSuccessStatusCode)
                return null;

            string html = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(html))
                return null;

            var matches = s_EnumerationRowRegex.Matches(html);
            if (matches.Count == 0)
                return null;

            var enumerations = new List<CustomFieldEnumeration>(matches.Count);
            foreach (Match match in matches.Cast<Match>())
            {
                string idValue = match.Groups["id"].Value;
                string content = match.Groups["content"].Value;

                Match nameMatch = s_EnumerationNameCellRegex.Match(content);
                if (!nameMatch.Success)
                    continue;

                string nameValue = nameMatch.Groups["name"].Value;

                if (!int.TryParse(idValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int id))
                    continue;

                string normalizedName = NormalizeHtmlText(nameValue);
                if (string.IsNullOrWhiteSpace(normalizedName))
                    continue;

                enumerations.Add(new CustomFieldEnumeration
                {
                    Id = id,
                    Name = normalizedName
                });
            }

            if (enumerations.Count == 0)
                return null;

            return ConvertEnumerationsToLookup(enumerations);
        }
        catch
        {
            return null;
        }
    }

    private static Dictionary<string, string> ConvertEnumerationsToLookup(IEnumerable<CustomFieldEnumeration> _Enumerations)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (CustomFieldEnumeration enumeration in _Enumerations)
        {
            if (enumeration == null || string.IsNullOrWhiteSpace(enumeration.Name))
                continue;

            string name = enumeration.Name.Trim();
            string idKey = enumeration.Id.ToString(CultureInfo.InvariantCulture);

            if (!map.ContainsKey(idKey))
                map[idKey] = name;

            if (!map.ContainsKey(name))
                map[name] = name;
        }

        return map;
    }

    private static string NormalizeHtmlText(string _Value)
    {
        if (string.IsNullOrWhiteSpace(_Value))
            return string.Empty;

        string withoutTags = s_HtmlTagRegex.Replace(_Value, " ");
        string decoded = WebUtility.HtmlDecode(withoutTags);
        return Regex.Replace(decoded, "\s+", " ").Trim();
    }

    private sealed class CustomFieldEnumerationResponse
    {
        [JsonPropertyName("custom_field_enumerations")]
        public List<CustomFieldEnumeration> CustomFieldEnumerations { get; set; } = new();
    }

    private sealed class CustomFieldEnumeration
    {
        [JsonPropertyName("id")]
        public int Id { get; set; }

        [JsonPropertyName("name")]
        public string? Name { get; set; }
    }

    #endregion
}
