using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAccel.Core.Formulas;

namespace ExcelAccel.Core.Auditing;

/// <summary>
/// Deterministic read-only projection of a parsed formula tree.
///
/// It renders structure only. It performs no subexpression evaluation, produces
/// no prose explanation, assigns no complexity or quality score, and never
/// touches a workbook.
/// </summary>
public sealed class FormulaInspectorReport
{
    private FormulaInspectorReport(
        AuditCellIdentity target,
        string formula,
        FormulaTreeResult tree,
        string headline,
        string completenessStatement,
        IReadOnlyList<string> summaryLines)
    {
        Target = target;
        Formula = formula;
        Tree = tree;
        Headline = headline;
        CompletenessStatement = completenessStatement;
        SummaryLines = summaryLines;
    }

    public AuditCellIdentity Target { get; }

    public string Formula { get; }

    public FormulaTreeResult Tree { get; }

    public string Headline { get; }

    public string CompletenessStatement { get; }

    public IReadOnlyList<string> SummaryLines { get; }

    public AuditTraceStatus Status => Tree.IsComplete
        ? AuditTraceStatus.Complete
        : Tree.Root is null ? AuditTraceStatus.Refused : AuditTraceStatus.Partial;

    public static FormulaInspectorReport Create(AuditCellIdentity target, string formula, FormulaTreeResult tree)
    {
        if (target is null) throw new ArgumentNullException(nameof(target));
        if (formula is null) throw new ArgumentNullException(nameof(formula));
        if (tree is null) throw new ArgumentNullException(nameof(tree));

        var location = AuditPresentationLabels.Location(target);
        var completeness = tree.IsComplete
            ? "The whole formula is represented; no part of it is omitted."
            : "This formula is not fully represented as a tree; the stated limitation applies.";

        var summary = new List<string>
        {
            "Cell: " + location,
            "Workbook: " + target.WorkbookId,
            "Nodes: " + AuditPresentationLabels.Count(tree.NodeCount),
            "Status: " + AuditPresentationLabels.Status(
                tree.IsComplete ? AuditTraceStatus.Complete : tree.Root is null ? AuditTraceStatus.Refused : AuditTraceStatus.Partial),
        };
        if (!string.IsNullOrWhiteSpace(tree.LimitationCode)) summary.Add("Limitation code: " + tree.LimitationCode);
        if (tree.LimitationSpan is { } span)
        {
            summary.Add("Limitation span: [" +
                span.Start.ToString(CultureInfo.InvariantCulture) + "+" +
                span.Length.ToString(CultureInfo.InvariantCulture) + "]");
        }

        if (!string.IsNullOrWhiteSpace(tree.Message)) summary.Add("Message: " + tree.Message);
        summary.Add("No subexpression is evaluated, scored, or explained.");

        var headline = tree.Root is null
            ? "The formula in " + location + " could not be represented as a tree: " + tree.Message + " (code " + tree.LimitationCode + ")"
            : AuditPresentationLabels.Count(tree.NodeCount) + " node" + (tree.NodeCount == 1 ? string.Empty : "s") +
                " in " + location + (tree.IsComplete ? "." : ", partially represented.");

        return new FormulaInspectorReport(target, formula, tree, headline, completeness, summary.AsReadOnly());
    }

    public TraceResultPresentation ToPresentation() => new TraceResultPresentation(
        "ExcelAccel Formula Inspector",
        Status,
        Headline,
        CompletenessStatement,
        new[]
        {
            new TraceColumn("Structure", 320),
            new TraceColumn("Kind", 140),
            new TraceColumn("Source span", 120),
            new TraceColumn("Source text", 260),
        },
        Rows(),
        SummaryLines,
        Tree.LimitationCode);

    /// <summary>
    /// Pre-order rows with indentation, which is also the keyboard focus order:
    /// moving down the list walks the tree exactly as it nests.
    /// </summary>
    private IEnumerable<TraceRow> Rows()
    {
        if (Tree.Root is null) return Array.Empty<TraceRow>();
        return Tree.Root.Flatten().Select(entry => new TraceRow(new[]
        {
            new string(' ', entry.Depth * 3) + Label(entry.Node),
            entry.Node.Kind.ToString(),
            "[" + entry.Node.Span.Start.ToString(CultureInfo.InvariantCulture) + "+" +
                entry.Node.Span.Length.ToString(CultureInfo.InvariantCulture) + "]",
            Excerpt(entry.Node.Span),
        }));
    }

    private static string Label(FormulaSyntaxNode node) => node.Kind switch
    {
        FormulaNodeKind.Function => node.Text + "( )",
        FormulaNodeKind.Group => "( )",
        _ => node.Text,
    };

    private string Excerpt(FormulaSourceSpan span)
    {
        if (span.Start < 0 || span.Start >= Formula.Length) return string.Empty;
        var length = Math.Min(span.Length, Formula.Length - span.Start);
        return Formula.Substring(span.Start, Math.Max(length, 0));
    }
}
