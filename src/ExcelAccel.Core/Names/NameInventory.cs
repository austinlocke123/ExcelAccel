using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAccel.Core.Formulas;

namespace ExcelAccel.Core.Names;

public enum NameScopeKind { Workbook, Worksheet }

/// <summary>What a defined name points at, as far as local metadata can tell.</summary>
public enum NameTargetKind
{
    /// <summary>A cell or range on a worksheet in this workbook.</summary>
    Range,

    /// <summary>A literal value rather than a reference.</summary>
    Constant,

    /// <summary>An expression that is neither a plain reference nor a literal.</summary>
    Formula,

    /// <summary>Reaches another workbook.</summary>
    External,

    /// <summary>The target no longer resolves, typically <c>#REF!</c>.</summary>
    Broken,

    /// <summary>Outside the qualified subset; kept visible rather than dropped.</summary>
    Unresolved,
}

/// <summary>
/// One name exactly as Excel reports it, before any classification. Keeping the
/// raw shape separate is what lets the whole classifier be tested without Excel.
/// </summary>
public sealed class NameRecord
{
    public NameRecord(string name, NameScopeKind scope, string? worksheetName, string refersTo, bool isVisible, bool isBuiltIn)
    {
        Name = string.IsNullOrWhiteSpace(name)
            ? throw new ArgumentException("A name is required.", nameof(name))
            : name.Trim();
        Scope = scope;
        WorksheetName = scope == NameScopeKind.Worksheet
            ? (!string.IsNullOrWhiteSpace(worksheetName)
                ? worksheetName!.Trim()
                : throw new ArgumentException("A worksheet-scoped name requires a worksheet.", nameof(worksheetName)))
            : null;
        RefersTo = refersTo ?? string.Empty;
        IsVisible = isVisible;
        IsBuiltIn = isBuiltIn;
    }

    public string Name { get; }
    public NameScopeKind Scope { get; }
    public string? WorksheetName { get; }
    public string RefersTo { get; }
    public bool IsVisible { get; }
    public bool IsBuiltIn { get; }
}

public sealed class NameInventoryEntry
{
    public NameInventoryEntry(
        string id,
        NameRecord record,
        NameTargetKind targetKind,
        bool isNavigable,
        bool hasDuplicateDisplayName,
        string? nonNavigableReason)
    {
        Id = id;
        Record = record;
        TargetKind = targetKind;
        IsNavigable = isNavigable;
        HasDuplicateDisplayName = hasDuplicateDisplayName;
        NonNavigableReason = nonNavigableReason;
    }

    /// <summary>Stable within one inventory: scope, worksheet, and name.</summary>
    public string Id { get; }

    public NameRecord Record { get; }
    public NameTargetKind TargetKind { get; }
    public bool IsNavigable { get; }

    /// <summary>
    /// The same display name exists in more than one scope. Excel allows it, and
    /// it is a common source of a formula resolving to the wrong one.
    /// </summary>
    public bool HasDuplicateDisplayName { get; }

    /// <summary>Why navigation is unavailable, when it is.</summary>
    public string? NonNavigableReason { get; }
}

public sealed class NameInventory
{
    public NameInventory(
        IReadOnlyList<NameInventoryEntry> entries,
        IReadOnlyDictionary<NameTargetKind, int> counts,
        int excludedByBound,
        IReadOnlyList<string> coverageGaps)
    {
        Entries = entries;
        Counts = counts;
        ExcludedByBound = excludedByBound;
        CoverageGaps = coverageGaps;
    }

    public IReadOnlyList<NameInventoryEntry> Entries { get; }
    public IReadOnlyDictionary<NameTargetKind, int> Counts { get; }

    /// <summary>
    /// Names dropped because the workbook exceeded the qualified bound. Reported
    /// rather than hidden: a truncated inventory that looks complete is worse
    /// than one that says so.
    /// </summary>
    public int ExcludedByBound { get; }

    public IReadOnlyList<string> CoverageGaps { get; }

    public bool IsComplete => ExcludedByBound == 0 && CoverageGaps.Count == 0;
}

/// <summary>
/// Recognises names Excel owns rather than the user. Excel exposes no built-in
/// flag, so they are identified by the reserved local names and the internal
/// prefixes it uses for its own bookkeeping.
/// </summary>
public static class ReservedNames
{
    private static readonly string[] Exact =
    {
        "Print_Area", "Print_Titles", "Criteria", "Database", "Extract",
        "Consolidate_Area", "Sheet_Title", "FilterDatabase", "_FilterDatabase",
    };

    /// <summary>
    /// Real workbooks carry names such as <c>_xlfn.SINGLE</c>, function shims
    /// Excel adds for itself. A real-Excel run surfaced one in an inventory the
    /// user would expect to hold only their own names.
    /// </summary>
    private static readonly string[] InternalPrefixes = { "_xlfn.", "_xlref", "_xludf.", "_xlchart." };

    public static bool IsReserved(string? localName)
    {
        var name = (localName ?? string.Empty).Trim();
        if (name.Length == 0)
        {
            return false;
        }

        foreach (var prefix in InternalPrefixes)
        {
            if (name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) return true;
        }

        foreach (var reserved in Exact)
        {
            if (string.Equals(name, reserved, StringComparison.OrdinalIgnoreCase)) return true;
        }

        return false;
    }
}

public static class NameInventoryBuilder
{
    /// <summary>Matches the existing audit-side name ceiling.</summary>
    public const int MaximumNames = 4_096;

    public static NameInventory Build(IEnumerable<NameRecord> records, bool includeHidden, bool includeBuiltIn)
    {
        if (records is null) throw new ArgumentNullException(nameof(records));

        var ordered = records
            .Where(record => includeHidden || record.IsVisible)
            .Where(record => includeBuiltIn || !record.IsBuiltIn)
            .OrderBy(record => record.Scope == NameScopeKind.Workbook ? 0 : 1)
            .ThenBy(record => record.WorksheetName ?? string.Empty, StringComparer.OrdinalIgnoreCase)
            .ThenBy(record => record.Name, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        var excluded = Math.Max(0, ordered.Length - MaximumNames);
        var retained = ordered.Take(MaximumNames).ToArray();

        var duplicates = retained
            .GroupBy(record => record.Name, StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() > 1)
            .Select(group => group.Key);
        var duplicateNames = new HashSet<string>(duplicates, StringComparer.OrdinalIgnoreCase);

        var entries = new List<NameInventoryEntry>(retained.Length);
        var counts = Enum.GetValues(typeof(NameTargetKind)).Cast<NameTargetKind>().ToDictionary(value => value, _ => 0);
        var gaps = new SortedSet<string>(StringComparer.Ordinal);

        foreach (var record in retained)
        {
            var kind = Classify(record.RefersTo);
            counts[kind]++;
            if (kind == NameTargetKind.Unresolved)
            {
                gaps.Add("name_expression_outside_qualified_subset");
            }

            var reason = NonNavigableReason(kind);
            entries.Add(new NameInventoryEntry(
                IdFor(record),
                record,
                kind,
                isNavigable: reason is null,
                hasDuplicateDisplayName: duplicateNames.Contains(record.Name),
                nonNavigableReason: reason));
        }

        if (excluded > 0)
        {
            gaps.Add("name_count_exceeded_qualified_bound");
        }

        return new NameInventory(entries, counts, excluded, gaps.ToArray());
    }

    public static string IdFor(NameRecord record) =>
        (record.Scope == NameScopeKind.Workbook ? "wb" : "ws")
        + "" + (record.WorksheetName ?? string.Empty)
        + "" + record.Name;

    /// <summary>
    /// Classifies a <c>RefersTo</c> expression using only local metadata. Nothing
    /// here opens or contacts an external source.
    /// </summary>
    public static NameTargetKind Classify(string? refersTo)
    {
        var value = (refersTo ?? string.Empty).Trim();
        if (value.Length == 0)
        {
            return NameTargetKind.Unresolved;
        }

        // Excel reports a lost target as #REF! inside the expression, and that
        // check has to come first: a broken external reference is still broken.
        if (value.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            return NameTargetKind.Broken;
        }

        var expression = value[0] == '=' ? value : "=" + value;
        var parsed = new FormulaParser().Parse(expression, FormulaParseOptions.DefaultA1);
        if (!parsed.IsSuccess || parsed.Document is null)
        {
            return NameTargetKind.Unresolved;
        }

        var references = parsed.Document.References;
        if (references.Any(reference => reference.Qualifier is not null && reference.Qualifier.IndexOf('[') >= 0))
        {
            return NameTargetKind.External;
        }

        if (references.Count == 0)
        {
            return IsLiteral(parsed.Document) ? NameTargetKind.Constant : NameTargetKind.Formula;
        }

        // A plain reference is a range; anything with operators or functions
        // around it is an expression that happens to contain one.
        return IsPlainReference(parsed.Document) ? NameTargetKind.Range : NameTargetKind.Formula;
    }

    /// <summary>
    /// Meaningful tokens exclude whitespace and the leading '=' prefix, which the
    /// parser reports as a token of its own.
    /// </summary>
    private static IReadOnlyList<FormulaToken> Meaningful(FormulaSyntaxDocument document) =>
        document.Tokens
            .Where(token => token.Kind != FormulaTokenKind.Whitespace && token.Kind != FormulaTokenKind.Prefix)
            .ToArray();

    private static bool IsPlainReference(FormulaSyntaxDocument document) =>
        document.References.Count == 1 && Meaningful(document).Count == 1;

    private static bool IsLiteral(FormulaSyntaxDocument document)
    {
        var meaningful = Meaningful(document);
        if (meaningful.Count == 1)
        {
            return meaningful[0].Kind == FormulaTokenKind.Number || meaningful[0].Kind == FormulaTokenKind.StringLiteral;
        }

        // A signed literal such as -1 arrives as an operator plus a number.
        return meaningful.Count == 2
            && meaningful[0].Kind == FormulaTokenKind.Operator
            && meaningful[1].Kind == FormulaTokenKind.Number;
    }

    private static string? NonNavigableReason(NameTargetKind kind)
    {
        switch (kind)
        {
            case NameTargetKind.Range: return null;
            case NameTargetKind.Constant: return "This name holds a constant, so there is nothing to select.";
            case NameTargetKind.Formula: return "This name holds an expression rather than a reference.";
            case NameTargetKind.External: return "This name reaches another workbook, which is never opened automatically.";
            case NameTargetKind.Broken: return "This name's target no longer resolves.";
            default: return "This name's expression is outside the qualified subset.";
        }
    }
}

/// <summary>
/// Deterministic local filtering over a completed inventory. No workbook is
/// re-read, so typing in the search box cannot trigger a scan.
/// </summary>
public static class NameInventorySearch
{
    public static IReadOnlyList<NameInventoryEntry> Filter(
        NameInventory inventory,
        string? query,
        NameScopeKind? scope = null,
        NameTargetKind? targetKind = null,
        bool brokenOnly = false)
    {
        if (inventory is null) throw new ArgumentNullException(nameof(inventory));

        var needle = (query ?? string.Empty).Trim();
        return inventory.Entries
            .Where(entry => scope is null || entry.Record.Scope == scope)
            .Where(entry => targetKind is null || entry.TargetKind == targetKind)
            .Where(entry => !brokenOnly || entry.TargetKind == NameTargetKind.Broken)
            .Where(entry => needle.Length == 0
                || entry.Record.Name.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0
                || entry.Record.RefersTo.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0
                || (entry.Record.WorksheetName ?? string.Empty).IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0)
            .ToArray();
    }

    /// <summary>
    /// The default export row. Raw expressions are excluded unless the user opts
    /// in, because a name's target can carry a path or a business term.
    /// </summary>
    public static IReadOnlyList<string> ExportFields(bool includeExpressions)
    {
        var fields = new List<string> { "name", "scope", "worksheet", "visible", "built_in", "target_kind", "navigable", "duplicate_name" };
        if (includeExpressions) fields.Add("refers_to");
        return fields;
    }

    public static IReadOnlyList<string> ExportRow(NameInventoryEntry entry, bool includeExpressions)
    {
        if (entry is null) throw new ArgumentNullException(nameof(entry));
        var row = new List<string>
        {
            entry.Record.Name,
            entry.Record.Scope == NameScopeKind.Workbook ? "workbook" : "worksheet",
            entry.Record.WorksheetName ?? string.Empty,
            entry.Record.IsVisible ? "true" : "false",
            entry.Record.IsBuiltIn ? "true" : "false",
            entry.TargetKind.ToString().ToLowerInvariant(),
            entry.IsNavigable ? "true" : "false",
            entry.HasDuplicateDisplayName ? "true" : "false",
        };
        if (includeExpressions) row.Add(entry.Record.RefersTo);
        return row;
    }

    public static string CoverageStatement(NameInventory inventory)
    {
        if (inventory is null) throw new ArgumentNullException(nameof(inventory));
        if (inventory.IsComplete)
        {
            return string.Format(
                CultureInfo.InvariantCulture, "Complete: {0:N0} name(s).", inventory.Entries.Count);
        }

        var parts = new List<string>
        {
            string.Format(CultureInfo.InvariantCulture, "Partial: {0:N0} name(s)", inventory.Entries.Count),
        };
        if (inventory.ExcludedByBound > 0)
        {
            parts.Add(string.Format(CultureInfo.InvariantCulture, "{0:N0} beyond the qualified bound", inventory.ExcludedByBound));
        }

        if (inventory.CoverageGaps.Count > 0)
        {
            parts.Add(string.Join(", ", inventory.CoverageGaps));
        }

        return string.Join("; ", parts) + ".";
    }
}
