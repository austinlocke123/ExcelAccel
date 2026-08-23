using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAccel.Core.Auditing;

namespace ExcelAccel.Core.Links;

/// <summary>
/// Projects a link inventory into the shared trace presentation, so the existing
/// read-only view renders it without a second window being written.
/// </summary>
public static class LinkInventoryPresentation
{
    private static readonly IReadOnlyList<TraceColumn> Columns = new[]
    {
        new TraceColumn("Source", 240),
        new TraceColumn("Status", 140),
        new TraceColumn("Category", 140),
        new TraceColumn("Where", 160),
        new TraceColumn("Detail", 260),
    };

    public static TraceResultPresentation Create(
        LinkInventory inventory,
        string workbookId,
        IReadOnlyList<LinkSourceGroup>? filtered = null)
    {
        if (inventory is null) throw new ArgumentNullException(nameof(inventory));
        if (string.IsNullOrWhiteSpace(workbookId)) throw new ArgumentException("A workbook is required.", nameof(workbookId));

        var sources = filtered ?? inventory.Sources;
        var rows = new List<TraceRow>();
        foreach (var source in sources)
        {
            if (source.Usages.Count == 0)
            {
                // Excel reports the source but nothing scanned explains it. Saying
                // so is more useful than omitting a link the user can see in
                // Excel's own dialog.
                rows.Add(new TraceRow(new[]
                {
                    source.DisplaySource,
                    Label(source.Status),
                    string.Empty,
                    string.Empty,
                    "No usage found in a scanned category.",
                }));
                continue;
            }

            foreach (var usage in source.Usages)
            {
                rows.Add(new TraceRow(
                    new[]
                    {
                        source.DisplaySource,
                        Label(source.Status),
                        usage.Record.Category.ToString(),
                        Where(usage.Record),
                        Detail(usage),
                    },
                    TryNavigationTarget(usage, workbookId)));
            }
        }

        return new TraceResultPresentation(
            "External links",
            inventory.IsComplete ? AuditTraceStatus.Complete : AuditTraceStatus.Partial,
            Headline(inventory, sources),
            LinkInventorySearch.CoverageStatement(inventory),
            Columns,
            rows,
            SummaryLines(inventory),
            refusalCode: null);
    }

    /// <summary>
    /// Status wording states what was observed in this session and never implies
    /// a filesystem or network check that did not happen (AC-LINK-003).
    /// </summary>
    public static string Label(LinkStatus status)
    {
        switch (status)
        {
            case LinkStatus.OpenInSession: return "Open in Excel";
            case LinkStatus.NotOpenInSession: return "Not open (not checked)";
            case LinkStatus.Broken: return "Broken reference";
            default: return "Unsupported source";
        }
    }

    private static string Headline(LinkInventory inventory, IReadOnlyList<LinkSourceGroup> shown) =>
        shown.Count == inventory.Sources.Count
            ? string.Format(CultureInfo.InvariantCulture, "{0:N0} source(s), {1:N0} usage(s).", inventory.Sources.Count, inventory.UsageCount)
            : string.Format(CultureInfo.InvariantCulture, "{0:N0} of {1:N0} source(s) shown.", shown.Count, inventory.Sources.Count);

    private static string Where(LinkUsageRecord record)
    {
        if (record.WorksheetName is null) return string.Empty;
        return record.Address is null ? record.WorksheetName : record.WorksheetName + "!" + record.Address;
    }

    private static string Detail(LinkUsage usage)
    {
        var parts = new List<string>();
        if (usage.Record.Detail.Length > 0) parts.Add(usage.Record.Detail);
        if (usage.NonNavigableReason is not null) parts.Add(usage.NonNavigableReason);
        return string.Join(" | ", parts);
    }

    private static IReadOnlyList<string> SummaryLines(LinkInventory inventory)
    {
        var lines = inventory.Counts
            .Where(pair => pair.Value > 0)
            .OrderBy(pair => pair.Key.ToString(), StringComparer.Ordinal)
            .Select(pair => string.Format(CultureInfo.InvariantCulture, "{0}: {1:N0} usage(s)", pair.Key, pair.Value))
            .ToList();

        foreach (var category in LinkInventoryBuilder.UnscannedCategories)
        {
            if (inventory.CoverageGaps.Contains("category_not_scanned:" + category.ToString().ToLowerInvariant()))
            {
                lines.Add(category + ": not scanned");
            }
        }

        if (inventory.ExcludedByBound > 0)
        {
            lines.Add(string.Format(
                CultureInfo.InvariantCulture, "Excluded by the qualified bound: {0:N0}", inventory.ExcludedByBound));
        }

        return lines;
    }

    /// <summary>
    /// Only a cell-anchored usage is offered as a jump, and the target is always
    /// inside this workbook. Navigating a link never opens its source.
    /// </summary>
    public static AuditCellIdentity? TryNavigationTarget(LinkUsage usage, string workbookId)
    {
        if (usage is null) throw new ArgumentNullException(nameof(usage));
        if (!usage.IsNavigable || !usage.Record.IsCellAnchored) return null;
        return new AuditCellIdentity(workbookId, usage.Record.WorksheetName!, usage.Record.Address!);
    }
}
