using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.Formulas;

namespace ExcelAccel.Core.Names;

/// <summary>
/// Projects a name inventory into the shared trace presentation, so the existing
/// read-only view renders it without a second window ever being written.
/// </summary>
public static class NameInventoryPresentation
{
    private static readonly IReadOnlyList<TraceColumn> Columns = new[]
    {
        new TraceColumn("Name", 200),
        new TraceColumn("Scope", 130),
        new TraceColumn("Target", 100),
        new TraceColumn("Refers to", 300),
        new TraceColumn("Note", 220),
    };

    public static TraceResultPresentation Create(
        NameInventory inventory,
        string workbookId,
        IReadOnlyList<NameInventoryEntry>? filtered = null)
    {
        if (inventory is null) throw new ArgumentNullException(nameof(inventory));
        if (string.IsNullOrWhiteSpace(workbookId)) throw new ArgumentException("A workbook is required.", nameof(workbookId));

        var entries = filtered ?? inventory.Entries;
        var rows = entries.Select(entry => new TraceRow(
            new[]
            {
                entry.Record.Name,
                Scope(entry.Record),
                entry.TargetKind.ToString(),
                entry.Record.RefersTo,
                Note(entry),
            },
            TryNavigationTarget(entry, workbookId)));

        var status = inventory.IsComplete ? AuditTraceStatus.Complete : AuditTraceStatus.Partial;
        return new TraceResultPresentation(
            "Named ranges",
            status,
            Headline(inventory, entries),
            NameInventorySearch.CoverageStatement(inventory),
            Columns,
            rows,
            SummaryLines(inventory),
            refusalCode: null);
    }

    private static string Headline(NameInventory inventory, IReadOnlyList<NameInventoryEntry> entries) =>
        entries.Count == inventory.Entries.Count
            ? string.Format(CultureInfo.InvariantCulture, "{0:N0} name(s).", entries.Count)
            : string.Format(CultureInfo.InvariantCulture, "{0:N0} of {1:N0} name(s) shown.", entries.Count, inventory.Entries.Count);

    private static string Scope(NameRecord record) =>
        record.Scope == NameScopeKind.Workbook ? "Workbook" : "Sheet: " + record.WorksheetName;

    private static string Note(NameInventoryEntry entry)
    {
        var notes = new List<string>();
        if (entry.HasDuplicateDisplayName) notes.Add("Duplicate name in another scope");
        if (!entry.Record.IsVisible) notes.Add("Hidden");
        if (entry.Record.IsBuiltIn) notes.Add("Built-in");
        if (entry.NonNavigableReason is not null) notes.Add(entry.NonNavigableReason);
        return string.Join(" | ", notes);
    }

    private static IReadOnlyList<string> SummaryLines(NameInventory inventory)
    {
        var lines = inventory.Counts
            .Where(pair => pair.Value > 0)
            .OrderBy(pair => pair.Key.ToString(), StringComparer.Ordinal)
            .Select(pair => string.Format(CultureInfo.InvariantCulture, "{0}: {1:N0}", pair.Key, pair.Value))
            .ToList();
        if (inventory.ExcludedByBound > 0)
        {
            lines.Add(string.Format(
                CultureInfo.InvariantCulture,
                "Excluded by the qualified bound: {0:N0}",
                inventory.ExcludedByBound));
        }

        return lines;
    }

    /// <summary>
    /// Derives the selectable target of a range name. Only a plain, single,
    /// sheet-qualified reference in this workbook is navigable; anything else
    /// stays inspectable but is not offered as a jump.
    /// </summary>
    public static AuditCellIdentity? TryNavigationTarget(NameInventoryEntry entry, string workbookId)
    {
        if (entry is null) throw new ArgumentNullException(nameof(entry));
        if (!entry.IsNavigable || entry.TargetKind != NameTargetKind.Range)
        {
            return null;
        }

        var value = entry.Record.RefersTo.Trim();
        var expression = value.Length > 0 && value[0] == '=' ? value : "=" + value;
        var parsed = new FormulaParser().Parse(expression, FormulaParseOptions.DefaultA1);
        if (!parsed.IsSuccess || parsed.Document is null || parsed.Document.References.Count != 1)
        {
            return null;
        }

        var reference = parsed.Document.References[0];
        if (reference.Qualifier is null)
        {
            return null;
        }

        var worksheet = Unquote(reference.Qualifier);
        if (worksheet.Length == 0 || worksheet.IndexOf('[') >= 0)
        {
            return null;
        }

        var address = AddressOf(reference.SourceText);
        return address.Length == 0
            ? null
            : new AuditCellIdentity(workbookId, worksheet, address);
    }

    /// <summary>Strips the sheet qualifier, leaving the address Excel can select.</summary>
    private static string AddressOf(string sourceText)
    {
        var text = sourceText ?? string.Empty;
        var separator = text.LastIndexOf('!');
        var address = separator >= 0 ? text.Substring(separator + 1) : text;
        return address.Replace("$", string.Empty).Trim();
    }

    private static string Unquote(string qualifier)
    {
        var name = qualifier.Trim();
        if (name.Length >= 2 && name[0] == '\'' && name[name.Length - 1] == '\'')
        {
            name = name.Substring(1, name.Length - 2).Replace("''", "'");
        }

        // A sheet qualifier arrives with its trailing '!' in some notations.
        return name.TrimEnd('!').Trim();
    }
}
