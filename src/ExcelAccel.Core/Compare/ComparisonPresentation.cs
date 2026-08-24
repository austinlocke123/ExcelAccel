using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAccel.Core.Auditing;

namespace ExcelAccel.Core.Compare;

/// <summary>
/// Projects a comparison into the shared trace presentation, so the existing
/// read-only view renders it without another window being written.
/// </summary>
public static class ComparisonPresentation
{
    private static readonly IReadOnlyList<TraceColumn> Columns = new[]
    {
        new TraceColumn("Position", 110),
        new TraceColumn("Category", 130),
        new TraceColumn("Source", 220),
        new TraceColumn("Target", 220),
        new TraceColumn("Note", 240),
    };

    /// <summary>
    /// Rows navigate to the source side. The target side is reached through
    /// <c>compare.result.navigate_target</c>, because a row carries one target
    /// and silently picking one side would be worse than saying which.
    /// </summary>
    public static TraceResultPresentation Create(ComparisonResult result, bool navigateTarget = false)
    {
        if (result is null) throw new ArgumentNullException(nameof(result));

        var side = navigateTarget ? result.Target : result.Source;
        var rows = result.Differences.Select(difference => new TraceRow(
            new[]
            {
                Position(difference),
                Label(difference.Category),
                Describe(difference.Source),
                Describe(difference.Target),
                difference.Statement,
            },
            NavigationTarget(difference, side, navigateTarget)));

        return new TraceResultPresentation(
            navigateTarget ? "Comparison (target side)" : "Comparison",
            result.IsComplete ? AuditTraceStatus.Complete : AuditTraceStatus.Partial,
            Headline(result),
            CoverageStatement(result),
            Columns,
            rows,
            SummaryLines(result),
            refusalCode: null);
    }

    public static string Label(DifferenceCategory category)
    {
        switch (category)
        {
            case DifferenceCategory.Formula: return "Formula";
            case DifferenceCategory.Constant: return "Value";
            case DifferenceCategory.DisplayedValue: return "Displayed value";
            case DifferenceCategory.NumberFormat: return "Number format";
            case DifferenceCategory.TypeMismatch: return "Type mismatch";
            default: return "Unsupported";
        }
    }

    public static string CoverageStatement(ComparisonResult result)
    {
        if (result is null) throw new ArgumentNullException(nameof(result));
        var head = string.Format(
            CultureInfo.InvariantCulture,
            "{0}: {1:N0} difference(s) across {2:N0} compared cell(s)",
            result.IsComplete ? "Complete" : "Partial",
            result.Differences.Count + result.ExcludedByBound,
            result.ComparedCells);
        return result.CoverageGaps.Count == 0
            ? head + "."
            : head + "; " + string.Join(", ", result.CoverageGaps) + ".";
    }

    private static string Headline(ComparisonResult result) =>
        string.Format(
            CultureInfo.InvariantCulture,
            "{0} vs {1}: {2:N0} difference(s).",
            result.Source.Identity,
            result.Target.Identity,
            result.Differences.Count + result.ExcludedByBound);

    private static string Position(CellDifference difference) =>
        string.Format(CultureInfo.InvariantCulture, "r{0}c{1}", difference.Row + 1, difference.Column + 1);

    /// <summary>
    /// A short excerpt. Raw content stays bounded here and is excluded from the
    /// export by default.
    /// </summary>
    private static string Describe(ComparisonCell cell)
    {
        var text = cell.IsFormula ? cell.Formula : cell.Value;
        if (text.Length == 0) return "(empty)";
        return text.Length <= 80 ? text : text.Substring(0, 77) + "...";
    }

    private static IReadOnlyList<string> SummaryLines(ComparisonResult result)
    {
        var lines = result.Counts
            .Where(pair => pair.Value > 0)
            .OrderBy(pair => pair.Key.ToString(), StringComparer.Ordinal)
            .Select(pair => string.Format(CultureInfo.InvariantCulture, "{0}: {1:N0}", Label(pair.Key), pair.Value))
            .ToList();

        var shapes = result.Differences.Count(d => d.FormulaKind == FormulaDifferenceKind.EquivalentShape);
        if (shapes > 0)
        {
            lines.Add(string.Format(CultureInfo.InvariantCulture, "Of the formulas, {0:N0} are the same shape with different references.", shapes));
        }

        if (result.ExcludedByBound > 0)
        {
            lines.Add(string.Format(CultureInfo.InvariantCulture, "Beyond the result bound: {0:N0}", result.ExcludedByBound));
        }

        return lines;
    }

    private static AuditCellIdentity? NavigationTarget(CellDifference difference, ComparisonBlock block, bool useTarget)
    {
        var cell = useTarget ? difference.Target : difference.Source;
        try
        {
            return new AuditCellIdentity(block.WorkbookId, block.WorksheetName, cell.Address);
        }
        catch (ArgumentException)
        {
            // A position without a usable address stays listed but non-navigable.
            return null;
        }
    }

    /// <summary>The export row. Raw formulas and values are excluded by default.</summary>
    public static IReadOnlyList<string> ExportFields(bool includeContent)
    {
        var fields = new List<string> { "position", "category", "formula_difference", "note" };
        if (includeContent)
        {
            fields.Add("source_content");
            fields.Add("target_content");
        }

        return fields;
    }

    public static IReadOnlyList<string> ExportRow(CellDifference difference, bool includeContent)
    {
        if (difference is null) throw new ArgumentNullException(nameof(difference));
        var row = new List<string>
        {
            Position(difference),
            Label(difference.Category),
            difference.FormulaKind.ToString(),
            difference.Statement,
        };
        if (includeContent)
        {
            row.Add(difference.Source.IsFormula ? difference.Source.Formula : difference.Source.Value);
            row.Add(difference.Target.IsFormula ? difference.Target.Formula : difference.Target.Value);
        }

        return row;
    }
}
