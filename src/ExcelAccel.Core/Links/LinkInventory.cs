using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ExcelAccel.Core.Links;

/// <summary>The kind of object that carries a link.</summary>
public enum LinkCategory
{
    FormulaReference,
    DefinedName,
    ChartSeries,
    QueryConnection,
    DataValidation,
}

/// <summary>
/// What is known about a source from local metadata alone.
/// </summary>
/// <remarks>
/// The names are deliberately about this Excel session rather than about the
/// file. Nothing here opens the source, probes the filesystem, or touches the
/// network, so "accessible" and "inaccessible" would claim a check that never
/// happened, which AC-LINK-003 forbids. <see cref="NotOpenInSession"/> says
/// exactly what was observed and nothing more.
/// </remarks>
public enum LinkStatus
{
    /// <summary>The source workbook is open in this Excel session.</summary>
    OpenInSession,

    /// <summary>
    /// Not open here. Whether the file exists or is reachable is unknown and was
    /// not checked.
    /// </summary>
    NotOpenInSession,

    /// <summary>A usage of this source no longer resolves, typically <c>#REF!</c>.</summary>
    Broken,

    /// <summary>The source token could not be normalized into anything meaningful.</summary>
    Unsupported,
}

/// <summary>One link-bearing object, exactly as captured.</summary>
public sealed class LinkUsageRecord
{
    public LinkUsageRecord(
        LinkCategory category,
        string sourceToken,
        string? worksheetName,
        string? address,
        string detail,
        bool isBroken)
    {
        Category = category;
        SourceToken = sourceToken ?? string.Empty;
        WorksheetName = string.IsNullOrWhiteSpace(worksheetName) ? null : worksheetName!.Trim();
        Address = string.IsNullOrWhiteSpace(address) ? null : address!.Trim();
        Detail = detail ?? string.Empty;
        IsBroken = isBroken;
    }

    public LinkCategory Category { get; }
    public string SourceToken { get; }
    public string? WorksheetName { get; }
    public string? Address { get; }

    /// <summary>Human-readable identification, such as the owning name or series.</summary>
    public string Detail { get; }

    public bool IsBroken { get; }

    /// <summary>
    /// Only a usage anchored to a cell can be selected. Chart series, queries,
    /// and names are listed but stay non-navigable.
    /// </summary>
    public bool IsCellAnchored =>
        (Category == LinkCategory.FormulaReference || Category == LinkCategory.DataValidation)
        && WorksheetName is not null
        && Address is not null;
}

/// <summary>A source Excel itself reports, whether or not any usage was found.</summary>
public sealed class LinkSourceRecord
{
    public LinkSourceRecord(string rawSource, bool isOpenInSession)
    {
        RawSource = rawSource ?? string.Empty;
        IsOpenInSession = isOpenInSession;
    }

    public string RawSource { get; }
    public bool IsOpenInSession { get; }
}

public sealed class LinkUsage
{
    public LinkUsage(string id, LinkUsageRecord record, bool isNavigable, string? nonNavigableReason)
    {
        Id = id;
        Record = record;
        IsNavigable = isNavigable;
        NonNavigableReason = nonNavigableReason;
    }

    public string Id { get; }
    public LinkUsageRecord Record { get; }
    public bool IsNavigable { get; }
    public string? NonNavigableReason { get; }
}

public sealed class LinkSourceGroup
{
    public LinkSourceGroup(string token, string displaySource, LinkStatus status, IReadOnlyList<LinkUsage> usages)
    {
        Token = token;
        DisplaySource = displaySource;
        Status = status;
        Usages = usages;
    }

    /// <summary>The normalized grouping key: the source file name, lowercased.</summary>
    public string Token { get; }

    /// <summary>
    /// The path Excel reported for this source, or the normalized file name when
    /// Excel reported none. Never a formula fragment.
    /// </summary>
    public string DisplaySource { get; }

    public LinkStatus Status { get; }
    public IReadOnlyList<LinkUsage> Usages { get; }
    public int UsageCount => Usages.Count;
}

public sealed class LinkInventory
{
    public LinkInventory(
        IReadOnlyList<LinkSourceGroup> sources,
        IReadOnlyDictionary<LinkCategory, int> counts,
        int excludedByBound,
        IReadOnlyList<string> coverageGaps)
    {
        Sources = sources;
        Counts = counts;
        ExcludedByBound = excludedByBound;
        CoverageGaps = coverageGaps;
    }

    public IReadOnlyList<LinkSourceGroup> Sources { get; }

    /// <summary>Usage counts per category, so coverage is reported per category.</summary>
    public IReadOnlyDictionary<LinkCategory, int> Counts { get; }

    public int ExcludedByBound { get; }
    public IReadOnlyList<string> CoverageGaps { get; }

    public int UsageCount => Sources.Sum(source => source.UsageCount);
    public bool IsComplete => ExcludedByBound == 0 && CoverageGaps.Count == 0;
}

public static class LinkInventoryBuilder
{
    public const int MaximumSources = 512;
    public const int MaximumUsages = 5_000;

    /// <summary>
    /// Categories no scanner exists for yet. Naming them keeps a partial
    /// inventory honest instead of letting it look complete.
    /// </summary>
    public static readonly IReadOnlyList<LinkCategory> UnscannedCategories = new[]
    {
        LinkCategory.ChartSeries,
        LinkCategory.QueryConnection,
        LinkCategory.DataValidation,
    };

    public static LinkInventory Build(
        IEnumerable<LinkSourceRecord> sources,
        IEnumerable<LinkUsageRecord> usages,
        IEnumerable<LinkCategory>? scannedCategories = null)
    {
        if (sources is null) throw new ArgumentNullException(nameof(sources));
        if (usages is null) throw new ArgumentNullException(nameof(usages));

        var scanned = new HashSet<LinkCategory>(
            scannedCategories ?? new[] { LinkCategory.FormulaReference, LinkCategory.DefinedName });

        var sourceByToken = new Dictionary<string, LinkSourceRecord>(StringComparer.OrdinalIgnoreCase);
        foreach (var source in sources)
        {
            var token = NormalizeToken(source.RawSource);
            if (token.Length == 0 || sourceByToken.ContainsKey(token)) continue;
            sourceByToken[token] = source;
        }

        var orderedUsages = usages
            .OrderBy(usage => NormalizeToken(usage.SourceToken), StringComparer.OrdinalIgnoreCase)
            .ThenBy(usage => usage.Category.ToString(), StringComparer.Ordinal)
            .ThenBy(usage => usage.WorksheetName ?? string.Empty, StringComparer.OrdinalIgnoreCase)
            .ThenBy(usage => usage.Address ?? string.Empty, StringComparer.OrdinalIgnoreCase)
            .ThenBy(usage => usage.Detail, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        var excluded = Math.Max(0, orderedUsages.Length - MaximumUsages);
        var retained = orderedUsages.Take(MaximumUsages).ToArray();

        var counts = Enum.GetValues(typeof(LinkCategory)).Cast<LinkCategory>().ToDictionary(value => value, _ => 0);
        var grouped = new Dictionary<string, List<LinkUsage>>(StringComparer.OrdinalIgnoreCase);
        var display = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var record in retained)
        {
            counts[record.Category]++;
            var token = NormalizeToken(record.SourceToken);
            if (!grouped.TryGetValue(token, out var list))
            {
                list = new List<LinkUsage>();
                grouped[token] = list;
            }

            var reason = NonNavigableReason(record);
            list.Add(new LinkUsage(
                IdFor(record, list.Count),
                record,
                isNavigable: reason is null,
                nonNavigableReason: reason));
        }

        // A source Excel reports but no scanned usage explains still belongs in
        // the list; the usage most likely sits in a category nothing scans yet.
        foreach (var pair in sourceByToken)
        {
            if (!grouped.ContainsKey(pair.Key)) grouped[pair.Key] = new List<LinkUsage>();
            Remember(display, pair.Key, pair.Value.RawSource);
        }

        // Anything Excel did not report a source for is shown as its normalized
        // token. A usage token is a raw formula, never a path, so it must not
        // become the display: doing so would print a formula fragment in a column
        // the user reads as a file, and would leak more than the path redaction
        // in the export is trying to withhold.
        foreach (var token in grouped.Keys)
        {
            if (!display.ContainsKey(token)) display[token] = token;
        }

        var groups = grouped
            .OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)
            .Take(MaximumSources)
            .Select(pair => new LinkSourceGroup(
                pair.Key,
                display.TryGetValue(pair.Key, out var shown) ? shown : pair.Key,
                Status(pair.Key, pair.Value, sourceByToken),
                pair.Value))
            .ToArray();

        var excludedSources = Math.Max(0, grouped.Count - MaximumSources);
        var gaps = new SortedSet<string>(StringComparer.Ordinal);
        foreach (var category in UnscannedCategories)
        {
            if (!scanned.Contains(category))
            {
                gaps.Add("category_not_scanned:" + category.ToString().ToLowerInvariant());
            }
        }

        if (excluded > 0) gaps.Add("usage_count_exceeded_qualified_bound");
        if (excludedSources > 0) gaps.Add("source_count_exceeded_qualified_bound");

        return new LinkInventory(groups, counts, excluded + excludedSources, gaps.ToArray());
    }

    /// <summary>
    /// Reduces a source to its file name so a formula token such as
    /// <c>[Budget.xlsx]Sheet1</c> and a full path to the same workbook land in
    /// one group. Purely textual: no path is resolved and no file is touched.
    /// </summary>
    public static string NormalizeToken(string? source)
    {
        var value = (source ?? string.Empty).Trim();
        if (value.Length == 0) return string.Empty;

        // Strip a formula-style qualifier down to the bracketed workbook.
        var open = value.IndexOf('[');
        var close = value.IndexOf(']');
        if (open >= 0 && close > open)
        {
            value = value.Substring(open + 1, close - open - 1);
        }

        value = value.Trim().Trim('\'');
        var separator = value.LastIndexOfAny(new[] { '\\', '/' });
        if (separator >= 0 && separator + 1 < value.Length)
        {
            value = value.Substring(separator + 1);
        }

        return value.Trim().ToLowerInvariant();
    }

    /// <summary>
    /// Records the display form for a source. Only a source Excel itself reported
    /// contributes one, because that is the only input guaranteed to be a path.
    /// </summary>
    private static void Remember(Dictionary<string, string> display, string token, string candidate)
    {
        var value = (candidate ?? string.Empty).Trim();
        if (value.Length == 0) return;
        if (!display.TryGetValue(token, out var existing) || value.Length > existing.Length)
        {
            display[token] = value;
        }
    }

    private static LinkStatus Status(
        string token,
        IReadOnlyList<LinkUsage> usages,
        IReadOnlyDictionary<string, LinkSourceRecord> sources)
    {
        if (token.Length == 0) return LinkStatus.Unsupported;
        if (usages.Any(usage => usage.Record.IsBroken)) return LinkStatus.Broken;
        return sources.TryGetValue(token, out var source) && source.IsOpenInSession
            ? LinkStatus.OpenInSession
            : LinkStatus.NotOpenInSession;
    }

    private static string IdFor(LinkUsageRecord record, int ordinal)
    {
        const char Separator = '';
        return NormalizeToken(record.SourceToken)
            + Separator + record.Category
            + Separator + (record.WorksheetName ?? string.Empty)
            + Separator + (record.Address ?? string.Empty)
            + Separator + ordinal.ToString(CultureInfo.InvariantCulture);
    }

    private static string? NonNavigableReason(LinkUsageRecord record)
    {
        if (record.IsBroken) return "This usage no longer resolves.";
        if (record.IsCellAnchored) return null;
        switch (record.Category)
        {
            case LinkCategory.DefinedName: return "A defined name is not a cell; open the name inventory to inspect it.";
            case LinkCategory.ChartSeries: return "Chart series selection is not a qualified navigation path yet.";
            case LinkCategory.QueryConnection: return "A query or connection is not a selectable worksheet object.";
            case LinkCategory.DataValidation: return "This validation usage is not anchored to a known cell.";
            default: return "This usage is not anchored to a selectable object.";
        }
    }
}

/// <summary>Deterministic local filtering over a completed inventory.</summary>
public static class LinkInventorySearch
{
    public static IReadOnlyList<LinkSourceGroup> Filter(
        LinkInventory inventory,
        string? query,
        LinkStatus? status = null,
        LinkCategory? category = null)
    {
        if (inventory is null) throw new ArgumentNullException(nameof(inventory));
        var needle = (query ?? string.Empty).Trim();

        return inventory.Sources
            .Where(source => status is null || source.Status == status)
            .Where(source => category is null || source.Usages.Any(usage => usage.Record.Category == category))
            .Where(source => needle.Length == 0
                || source.DisplaySource.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0
                || source.Usages.Any(usage =>
                    (usage.Record.WorksheetName ?? string.Empty).IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0
                    || usage.Record.Detail.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0))
            .ToArray();
    }

    /// <summary>
    /// Source paths are excluded unless the user opts in. A link path routinely
    /// names a drive, a client, or a deal.
    /// </summary>
    public static IReadOnlyList<string> ExportFields(bool includePaths)
    {
        var fields = new List<string> { "source", "status", "category", "worksheet", "address", "detail", "navigable" };
        if (includePaths) fields[0] = "source_path";
        return fields;
    }

    public static IReadOnlyList<string> ExportRow(LinkSourceGroup source, LinkUsage usage, bool includePaths)
    {
        if (source is null) throw new ArgumentNullException(nameof(source));
        if (usage is null) throw new ArgumentNullException(nameof(usage));
        return new List<string>
        {
            includePaths ? source.DisplaySource : source.Token,
            source.Status.ToString(),
            usage.Record.Category.ToString(),
            usage.Record.WorksheetName ?? string.Empty,
            usage.Record.Address ?? string.Empty,
            usage.Record.Detail,
            usage.IsNavigable ? "true" : "false",
        };
    }

    public static string CoverageStatement(LinkInventory inventory)
    {
        if (inventory is null) throw new ArgumentNullException(nameof(inventory));
        var head = string.Format(
            CultureInfo.InvariantCulture,
            "{0}: {1:N0} source(s), {2:N0} usage(s)",
            inventory.IsComplete ? "Complete" : "Partial",
            inventory.Sources.Count,
            inventory.UsageCount);
        return inventory.CoverageGaps.Count == 0
            ? head + "."
            : head + "; " + string.Join(", ", inventory.CoverageGaps) + ".";
    }
}
