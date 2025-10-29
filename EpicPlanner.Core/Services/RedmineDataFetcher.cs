using System.Collections.Generic;
using System.Collections.Specialized;
using System.Globalization;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using Redmine.Net.Api;
using Redmine.Net.Api.Net;
using Redmine.Net.Api.Types;

namespace EpicPlanner.Core;

public class RedmineDataFetcher
{
    #region Members

    private const int EpicCustomFieldId = 57;
    private const string ParentIdKey = "parent_id";

    private readonly RedmineManager m_RedmineManager;
    private readonly HttpClient m_HttpClient;
    private readonly SemaphoreSlim m_EpicEnumerationLock = new(1, 1);
    private Dictionary<int, string>? m_EpicEnumerationCache;

    #endregion

    #region Constructor

    public RedmineDataFetcher(string _BaseUrl, string _ApiKey)
    {
        if (string.IsNullOrWhiteSpace(_BaseUrl))
            throw new ArgumentException("Redmine base URL must be provided.", nameof(_BaseUrl));

        m_RedmineManager = new RedmineManager(new RedmineManagerOptionsBuilder()
            .WithHost(_BaseUrl)
            .WithApiKeyAuthentication(_ApiKey));

        string normalizedBaseUrl = _BaseUrl.EndsWith("/") ? _BaseUrl : _BaseUrl + "/";

        m_HttpClient = new HttpClient
        {
            BaseAddress = new Uri(normalizedBaseUrl)
        };
        m_HttpClient.DefaultRequestHeaders.Add("X-Redmine-API-Key", _ApiKey);
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

    public async Task<Dictionary<string, double>> GetPlannedHoursForSprintAsync(
        int _iSprintNumber,
        DateTime _dtSprintStart,
        DateTime _dtSprintEnd,
        IReadOnlyCollection<string> _PlannedEpics)
    {
        var result = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        HashSet<string>? plannedEpicSet = (_PlannedEpics?.Count ?? 0) > 0
            ? new HashSet<string>(_PlannedEpics, StringComparer.OrdinalIgnoreCase)
            : null;

        Dictionary<int, Issue>? parentCache = plannedEpicSet != null
            ? new Dictionary<int, Issue>()
            : null;

        IEnumerable<Issue> issues = await GetSprintIssuesAsync(_iSprintNumber);
        foreach (var issue in issues)
        {
            if (issue.AssignedTo == null || issue.Subject.Contains("[Suivi]") || issue.Subject.Contains("[Analyse]"))
                continue;

            if (plannedEpicSet != null)
            {
                EpicDescriptor descriptor = await ResolveEpicDescriptorAsync(issue, parentCache!).ConfigureAwait(false);
                if (!plannedEpicSet.Contains(descriptor.Name))
                    continue;

                if (issue.Id > 0)
                {
                    parentCache[issue.Id] = issue;
                }
            }

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

    public async Task<List<SprintEpicSummary>> GetEpicSprintSummariesAsync(
        int _iSprintNumber,
        DateTime _dtSprintStart,
        DateTime _dtSprintEnd,
        IReadOnlyCollection<string> _PlannedEpics,
        IReadOnlyDictionary<string, double>? _PlannedCapacityByEpic = null)
    {
        var summaries = new Dictionary<string, SprintEpicSummary>(StringComparer.OrdinalIgnoreCase);
        var descriptors = new Dictionary<string, EpicDescriptor>(StringComparer.OrdinalIgnoreCase);
        var sprintRemaining = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        HashSet<string>? plannedEpicSet = (_PlannedEpics?.Count ?? 0) > 0
            ? new HashSet<string>(_PlannedEpics, StringComparer.OrdinalIgnoreCase)
            : null;
        bool hasSheetPlannedCapacity = _PlannedCapacityByEpic != null && _PlannedCapacityByEpic.Count > 0;

        List<Issue> sprintIssues = (await GetSprintIssuesAsync(_iSprintNumber)).ToList();

        var parentCache = new Dictionary<int, Issue>();
        foreach (var issue in sprintIssues)
        {
            if (issue.Subject.Contains("[Suivi]") || issue.Subject.Contains("[Analyse]"))
                continue;

            EpicDescriptor descriptor = await ResolveEpicDescriptorAsync(issue, parentCache).ConfigureAwait(false);

            if (plannedEpicSet != null && !plannedEpicSet.Contains(descriptor.Name))
                continue;

            if (!descriptors.TryGetValue(descriptor.Name, out EpicDescriptor? existing) ||
                (!existing.HasEnumeration && descriptor.HasEnumeration))
            {
                descriptors[descriptor.Name] = descriptor;
            }

            if (!summaries.TryGetValue(descriptor.Name, out var summary))
            {
                summary = new SprintEpicSummary { Epic = descriptor.Name };
                summaries[descriptor.Name] = summary;
            }

            double planned = issue.EstimatedHours ?? 0.0;
            double remaining = ExtractRemaining(issue);

            summary.PlannedCapacity += planned;

            if (remaining > 0)
            {
                if (!sprintRemaining.ContainsKey(descriptor.Name))
                    sprintRemaining[descriptor.Name] = 0.0;

                sprintRemaining[descriptor.Name] += remaining;
            }
        }

        if (plannedEpicSet != null && plannedEpicSet.Count > 0)
        {
            Dictionary<int, string> enumerationLookup = await GetEpicEnumerationLookupAsync().ConfigureAwait(false);
            var nameToEnumerationId = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var kvp in enumerationLookup)
            {
                string? normalizedName = NormalizeEpicName(kvp.Value);
                if (string.IsNullOrWhiteSpace(normalizedName))
                    continue;

                if (!nameToEnumerationId.ContainsKey(normalizedName))
                    nameToEnumerationId[normalizedName] = kvp.Key;
            }

            foreach (string plannedEpic in plannedEpicSet)
            {
                if (!descriptors.ContainsKey(plannedEpic))
                {
                    int? enumerationId = null;
                    if (nameToEnumerationId.TryGetValue(plannedEpic, out int resolvedId))
                        enumerationId = resolvedId;

                    descriptors[plannedEpic] = new EpicDescriptor(plannedEpic, enumerationId);
                }

                if (!summaries.ContainsKey(plannedEpic))
                    summaries[plannedEpic] = new SprintEpicSummary { Epic = plannedEpic };
            }
        }

        Dictionary<string, EpicTodoCache> todoCaches = await BuildEpicTodoCachesAsync(descriptors.Values).ConfigureAwait(false);
        var todoIssueToEpic = new Dictionary<int, string>();

        foreach (var cacheEntry in todoCaches)
        {
            foreach (int issueId in cacheEntry.Value.TodoIssueIds)
            {
                if (issueId > 0)
                {
                    todoIssueToEpic[issueId] = cacheEntry.Key;
                }
            }
        }

        List<TimeEntryRecord> timeEntries = await GetTimeEntriesAsync(_dtSprintStart, _dtSprintEnd).ConfigureAwait(false);
        HashSet<string> plannedEpicNames = plannedEpicSet != null
            ? new HashSet<string>(plannedEpicSet, StringComparer.OrdinalIgnoreCase)
            : new HashSet<string>(descriptors.Keys, StringComparer.OrdinalIgnoreCase);

        foreach (TimeEntryRecord entry in timeEntries)
        {
            if (entry?.Hours <= 0)
                continue;

            int issueId = entry.Issue?.Id ?? 0;
            if (issueId <= 0)
                continue;

            if (!todoIssueToEpic.TryGetValue(issueId, out string? epicName))
                continue;

            if (!plannedEpicNames.Contains(epicName))
                continue;

            if (!descriptors.TryGetValue(epicName, out EpicDescriptor? descriptor))
                continue;

            if (!summaries.TryGetValue(descriptor.Name, out var summary))
            {
                summary = new SprintEpicSummary { Epic = descriptor.Name };
                summaries[descriptor.Name] = summary;
            }

            summary.Consumed += entry.Hours;
        }

        if (hasSheetPlannedCapacity)
        {
            foreach (var summary in summaries.Values)
            {
                if (_PlannedCapacityByEpic!.TryGetValue(summary.Epic, out double plannedCapacity))
                {
                    summary.PlannedCapacity = plannedCapacity;
                }
            }
        }

        foreach (var summary in summaries.Values)
        {
            if (todoCaches.TryGetValue(summary.Epic, out EpicTodoCache? cache))
            {
                summary.Remaining = cache.TotalRemaining;
            }
            else if (sprintRemaining.TryGetValue(summary.Epic, out double sprintTotal))
            {
                summary.Remaining = sprintTotal;
            }
            else
            {
                summary.Remaining = 0.0;
            }
        }

        int plannedEpicCount = _PlannedEpics?.Count ?? 0;

        if (plannedEpicSet != null && plannedEpicCount > 0)
        {
            var orderedSummaries = new List<SprintEpicSummary>(plannedEpicCount);
            foreach (string plannedEpic in _PlannedEpics!)
            {
                if (summaries.TryGetValue(plannedEpic, out SprintEpicSummary? existing))
                {
                    orderedSummaries.Add(existing);
                }
                else
                {
                    orderedSummaries.Add(new SprintEpicSummary { Epic = plannedEpic });
                }
            }

            return orderedSummaries;
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

    private async Task<List<TimeEntryRecord>> GetTimeEntriesAsync(DateTime _dtSprintStart, DateTime _dtSprintEnd)
    {
        var results = new List<TimeEntryRecord>();
        int offset = 0;
        const int limit = 100;

        try
        {
            while (true)
            {
                string url = $"time_entries.json?from={_dtSprintStart:yyyy-MM-dd}&to={_dtSprintEnd:yyyy-MM-dd}&offset={offset}&limit={limit}";
                using HttpResponseMessage response = await m_HttpClient.GetAsync(url).ConfigureAwait(false);

                if (!response.IsSuccessStatusCode)
                    break;

                await using Stream stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
                TimeEntriesResponse? payload = await JsonSerializer
                    .DeserializeAsync<TimeEntriesResponse>(stream, new JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true
                    })
                    .ConfigureAwait(false);

                if (payload?.TimeEntries == null || payload.TimeEntries.Count == 0)
                    break;

                results.AddRange(payload.TimeEntries.Where(e => e != null));

                offset += payload.TimeEntries.Count;

                int totalCount = payload.TotalCount;
                if (totalCount <= 0 || offset >= totalCount)
                    break;
            }
        }
        catch
        {
            // Swallow exception and return data fetched so far to avoid blocking planner execution.
        }

        return results;
    }

    private async Task<EpicDescriptor> ResolveEpicDescriptorAsync(
        Issue _Issue,
        Dictionary<int, Issue> _ParentCache)
    {
        EpicCustomFieldValue? directValue = await ExtractEpicCustomFieldValueAsync(_Issue).ConfigureAwait(false);
        if (directValue != null)
            return CreateDescriptor(directValue);

        Issue? parent = await GetParentIssueAsync(_Issue, _ParentCache).ConfigureAwait(false);
        if (parent != null)
        {
            EpicCustomFieldValue? parentValue = await ExtractEpicCustomFieldValueAsync(parent).ConfigureAwait(false);
            if (parentValue != null)
                return CreateDescriptor(parentValue);
        }

        string? fromSubject = NormalizeEpicName(ExtractEpicFromSubject(_Issue));
        if (!string.IsNullOrWhiteSpace(fromSubject))
            return new EpicDescriptor(fromSubject, null);

        if (parent != null)
        {
            string? parentSubject = NormalizeEpicName(ExtractEpicFromSubject(parent));
            if (!string.IsNullOrWhiteSpace(parentSubject))
                return new EpicDescriptor(parentSubject, null);
        }

        return new EpicDescriptor("(No Epic)", null);
    }

    private static string? NormalizeEpicName(string? _RawEpicValue)
    {
        if (string.IsNullOrWhiteSpace(_RawEpicValue))
            return null;

        return _RawEpicValue.Trim();
    }

    private async Task<EpicCustomFieldValue?> ExtractEpicCustomFieldValueAsync(Issue _Issue)
    {
        if (_Issue.CustomFields == null)
            return null;

        IssueCustomField? epicField = _Issue.CustomFields
            .FirstOrDefault(cf => cf.Id == EpicCustomFieldId);

        if (epicField == null || epicField.Values == null)
            return null;

        string? rawValue = epicField.Values
            .Select(v => v.Info)
            .FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));

        if (string.IsNullOrWhiteSpace(rawValue))
            return null;

        string trimmedValue = rawValue.Trim();

        int? enumerationId = null;
        string? label = null;

        if (int.TryParse(trimmedValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int enumId) && enumId > 0)
        {
            enumerationId = enumId;

            Dictionary<int, string> lookup = await GetEpicEnumerationLookupAsync().ConfigureAwait(false);
            if (lookup.TryGetValue(enumId, out string? mappedValue) && !string.IsNullOrWhiteSpace(mappedValue))
                label = mappedValue.Trim();
        }

        label ??= trimmedValue;

        return new EpicCustomFieldValue
        {
            EnumerationId = enumerationId,
            Label = label,
            RawValue = trimmedValue
        };
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

    private async Task<Issue?> GetIssueByIdAsync(int _iIssueId)
    {
        var parameters = new NameValueCollection
        {
            { RedmineKeys.ISSUE_ID, _iIssueId.ToString(CultureInfo.InvariantCulture) }
        };

        return (await GetIssuesAsync(parameters)).FirstOrDefault();
    }

    private static EpicDescriptor CreateDescriptor(EpicCustomFieldValue _Value)
    {
        string? normalized = NormalizeEpicName(_Value.Label);
        if (string.IsNullOrWhiteSpace(normalized) && !string.IsNullOrWhiteSpace(_Value.RawValue))
            normalized = NormalizeEpicName(_Value.RawValue);

        string name = normalized ?? "(No Epic)";
        return new EpicDescriptor(name, _Value.EnumerationId);
    }

    private async Task<Dictionary<string, EpicTodoCache>> BuildEpicTodoCachesAsync(IEnumerable<EpicDescriptor> _Descriptors)
    {
        var caches = new Dictionary<string, EpicTodoCache>(StringComparer.OrdinalIgnoreCase);

        var groupedById = _Descriptors
            .Where(d => d.HasEnumeration)
            .GroupBy(d => d.EnumerationId!.Value);

        foreach (var group in groupedById)
        {
            List<Issue> parentIssues = await FetchEpicParentIssuesAsync(group.Key).ConfigureAwait(false);
            if (parentIssues.Count == 0)
            {
                foreach (EpicDescriptor descriptor in group)
                {
                    caches[descriptor.Name] = EpicTodoCache.Empty;
                }

                continue;
            }

            List<Issue> todoIssues = await FetchTodoDescendantsAsync(parentIssues).ConfigureAwait(false);

            double totalRemaining = 0.0;
            HashSet<int> todoIssueIds = new();

            foreach (Issue todo in todoIssues)
            {
                if (todo?.Id > 0)
                    todoIssueIds.Add(todo.Id);

                totalRemaining += ExtractRemaining(todo);
            }

            var cache = new EpicTodoCache(todoIssueIds, totalRemaining);

            foreach (EpicDescriptor descriptor in group)
            {
                caches[descriptor.Name] = cache;
            }
        }

        return caches;
    }

    private async Task<List<Issue>> FetchEpicParentIssuesAsync(int _iEnumerationId)
    {
        var parents = new Dictionary<int, Issue>();
        foreach (int trackerId in new[] { 9, 4 })
        {
            var parameters = new NameValueCollection
            {
                { RedmineKeys.TRACKER_ID, trackerId.ToString(CultureInfo.InvariantCulture) },
                { $"cf_{EpicCustomFieldId}", _iEnumerationId.ToString(CultureInfo.InvariantCulture) },
                { RedmineKeys.STATUS_ID, "*" }
            };

            IEnumerable<Issue> issues = await GetIssuesAsync(parameters).ConfigureAwait(false);
            foreach (Issue issue in issues)
            {
                parents[issue.Id] = issue;
            }
        }

        return parents.Values.ToList();
    }

    private async Task<List<Issue>> FetchTodoDescendantsAsync(IEnumerable<Issue> _ParentIssues)
    {
        var todoIssues = new List<Issue>();
        HashSet<int> countedTodoIds = new();
        HashSet<int> visitedParents = new();
        Queue<int> parentsToVisit = new();

        foreach (Issue parent in _ParentIssues)
        {
            if (parent == null)
                continue;

            int parentId = parent.Id;
            if (parentId > 0 && visitedParents.Add(parentId))
                parentsToVisit.Enqueue(parentId);
        }

        while (parentsToVisit.Count > 0)
        {
            int currentParentId = parentsToVisit.Dequeue();

            var parameters = new NameValueCollection
            {
                { ParentIdKey, currentParentId.ToString(CultureInfo.InvariantCulture) },
                { RedmineKeys.STATUS_ID, "*" }
            };

            IEnumerable<Issue> children = await GetIssuesAsync(parameters).ConfigureAwait(false);
            foreach (Issue child in children)
            {
                if (child == null)
                    continue;

                if (child.Subject.Contains("[Suivi]") || child.Subject.Contains("[Analyse]"))
                    continue;

                int childId = child.Id;

                if (child.Tracker?.Id == 6)
                {
                    if (childId <= 0 || countedTodoIds.Add(childId))
                        todoIssues.Add(child);

                    continue;
                }

                if (childId > 0 && visitedParents.Add(childId))
                    parentsToVisit.Enqueue(childId);
            }
        }

        return todoIssues;
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

    private async Task<Dictionary<int, string>> GetEpicEnumerationLookupAsync()
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
                Dictionary<int, string>? map = await TryFetchEpicEnumerationsAsync().ConfigureAwait(false);
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

            m_EpicEnumerationCache = new Dictionary<int, string>();
            return m_EpicEnumerationCache;
        }
        finally
        {
            m_EpicEnumerationLock.Release();
        }
    }

    private async Task<Dictionary<int, string>?> TryFetchEpicEnumerationsAsync()
    {
        Dictionary<int, string>? map = await TryFetchEpicEnumerationsFromCustomFieldsAsync().ConfigureAwait(false);
        if (map != null && map.Count > 0)
            return map;

        return map;
    }

    private async Task<Dictionary<int, string>?> TryFetchEpicEnumerationsFromCustomFieldsAsync()
    {
        try
        {
            using HttpResponseMessage response = await m_HttpClient
                .GetAsync("custom_fields.json")
                .ConfigureAwait(false);

            if (!response.IsSuccessStatusCode)
                return null;

            await using var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
            CustomFieldsResponse? payload = await JsonSerializer
                .DeserializeAsync<CustomFieldsResponse>(stream)
                .ConfigureAwait(false);

            if (payload?.CustomFields == null || payload.CustomFields.Count == 0)
                return null;

            CustomField? epicField = payload.CustomFields
                .FirstOrDefault(cf => cf.Id == EpicCustomFieldId);

            if (epicField == null)
                return null;

            List<CustomFieldEnumeration>? enumerations = epicField.CustomFieldEnumerations;
            if (enumerations == null || enumerations.Count == 0)
            {
                if (epicField.PossibleValues != null && epicField.PossibleValues.Count > 0)
                {
                    enumerations = epicField.PossibleValues
                        .Where(pv => !string.IsNullOrWhiteSpace(pv.Value) || !string.IsNullOrWhiteSpace(pv.Label))
                        .Select(pv => new CustomFieldEnumeration
                        {
                            Id = TryParseEnumerationId(pv.Value),
                            Value = pv.Value,
                            Label = pv.Label,
                            Name = pv.Label
                        })
                        .ToList();
                }
            }

            if (enumerations == null || enumerations.Count == 0)
                return null;

            return ConvertEnumerationsToLookup(enumerations);
        }
        catch
        {
            return null;
        }
    }

    private static Dictionary<int, string> ConvertEnumerationsToLookup(IEnumerable<CustomFieldEnumeration> _Enumerations)
    {
        var map = new Dictionary<int, string>();
        foreach (CustomFieldEnumeration enumeration in _Enumerations)
        {
            if (enumeration == null)
                continue;

            string? rawValue = enumeration.Value;
            int id = enumeration.Id > 0
                ? enumeration.Id
                : TryParseEnumerationId(rawValue);

            if (id <= 0)
                continue;

            string? name = enumeration.Label ?? enumeration.Name;
            if (string.IsNullOrWhiteSpace(name) && !string.IsNullOrWhiteSpace(rawValue))
            {
                string trimmed = rawValue.Trim();
                if (!string.Equals(trimmed, id.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal))
                    name = trimmed;
            }
            if (string.IsNullOrWhiteSpace(name))
                continue;

            map[id] = name.Trim();
        }

        return map;
    }

    private sealed class EpicTodoCache
    {
        public static readonly EpicTodoCache Empty = new(new HashSet<int>(), 0.0);

        public EpicTodoCache(HashSet<int> _TodoIssueIds, double _dTotalRemaining)
        {
            TodoIssueIds = _TodoIssueIds;
            TotalRemaining = _dTotalRemaining;
        }

        public HashSet<int> TodoIssueIds { get; }

        public double TotalRemaining { get; }
    }

    private sealed class EpicDescriptor
    {
        public EpicDescriptor(string _Name, int? _iEnumerationId)
        {
            Name = _Name;
            EnumerationId = _iEnumerationId;
        }

        public string Name { get; }

        public int? EnumerationId { get; }

        public bool HasEnumeration => EnumerationId.HasValue;
    }

    private sealed class EpicCustomFieldValue
    {
        public int? EnumerationId { get; init; }

        public string? Label { get; init; }

        public string? RawValue { get; init; }
    }

    private sealed class CustomFieldsResponse
    {
        [JsonPropertyName("custom_fields")]
        public List<CustomField> CustomFields { get; set; } = new();
    }

    private sealed class CustomField
    {
        [JsonPropertyName("id")]
        public int Id { get; set; }

        [JsonPropertyName("name")]
        public string? Name { get; set; }

        [JsonPropertyName("custom_field_enumerations")]
        public List<CustomFieldEnumeration>? CustomFieldEnumerations { get; set; }

        [JsonPropertyName("possible_values")]
        public List<CustomFieldPossibleValue>? PossibleValues { get; set; }
    }

    private sealed class CustomFieldEnumeration
    {
        [JsonPropertyName("id")]
        public int Id { get; set; }

        [JsonPropertyName("name")]
        public string? Name { get; set; }

        [JsonPropertyName("value")]
        public string? Value { get; set; }

        [JsonPropertyName("label")]
        public string? Label { get; set; }
    }

    private sealed class CustomFieldPossibleValue
    {
        [JsonPropertyName("value")]
        public string? Value { get; set; }

        [JsonPropertyName("label")]
        public string? Label { get; set; }
    }

    private sealed class TimeEntriesResponse
    {
        [JsonPropertyName("time_entries")]
        public List<TimeEntryRecord> TimeEntries { get; set; } = new();

        [JsonPropertyName("total_count")]
        public int TotalCount { get; set; }

        [JsonPropertyName("offset")]
        public int Offset { get; set; }

        [JsonPropertyName("limit")]
        public int Limit { get; set; }
    }

    private sealed class TimeEntryRecord
    {
        [JsonPropertyName("id")]
        public int Id { get; set; }

        [JsonPropertyName("hours")]
        public double Hours { get; set; }

        [JsonPropertyName("spent_on")]
        public DateTime SpentOn { get; set; }

        [JsonPropertyName("issue")]
        public TimeEntryIssueReference? Issue { get; set; }
    }

    private sealed class TimeEntryIssueReference
    {
        [JsonPropertyName("id")]
        public int Id { get; set; }
    }

    private static int TryParseEnumerationId(string? _Value)
    {
        if (string.IsNullOrWhiteSpace(_Value))
            return 0;

        return int.TryParse(_Value.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)
            ? parsed
            : 0;
    }

    #endregion
}