using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAccel.Core.Auditing;

public sealed class TraceColumn
{
    public TraceColumn(string header, int width)
    {
        Header = !string.IsNullOrWhiteSpace(header) ? header : throw new ArgumentException("A column header is required.", nameof(header));
        Width = width > 0 ? width : throw new ArgumentOutOfRangeException(nameof(width));
    }

    public string Header { get; }

    public int Width { get; }
}

/// <summary>
/// The display-ready shape every auditing result projects into: a headline, a
/// completeness statement, a table, and a summary. It carries no analysis and no
/// host type, so one view can render every trace result and the views cannot
/// drift apart in how they present the same states.
/// </summary>
public sealed class TraceResultPresentation
{
    public TraceResultPresentation(
        string title,
        AuditTraceStatus status,
        string headline,
        string completenessStatement,
        IReadOnlyList<TraceColumn> columns,
        IEnumerable<IReadOnlyList<string>> rows,
        IReadOnlyList<string> summaryLines,
        string? refusalCode)
    {
        Title = !string.IsNullOrWhiteSpace(title) ? title : throw new ArgumentException("A title is required.", nameof(title));
        Status = status;
        Headline = headline ?? throw new ArgumentNullException(nameof(headline));
        CompletenessStatement = completenessStatement ?? throw new ArgumentNullException(nameof(completenessStatement));
        Columns = columns ?? throw new ArgumentNullException(nameof(columns));
        if (Columns.Count == 0) throw new ArgumentException("At least one column is required.", nameof(columns));
        Rows = Array.AsReadOnly((rows ?? throw new ArgumentNullException(nameof(rows))).ToArray());
        if (Rows.Any(row => row is null || row.Count != Columns.Count))
        {
            throw new ArgumentException("Every row must supply exactly one value per column.", nameof(rows));
        }

        SummaryLines = summaryLines ?? throw new ArgumentNullException(nameof(summaryLines));
        RefusalCode = refusalCode;
    }

    public string Title { get; }

    public AuditTraceStatus Status { get; }

    public string Headline { get; }

    public string CompletenessStatement { get; }

    public IReadOnlyList<TraceColumn> Columns { get; }

    public IReadOnlyList<IReadOnlyList<string>> Rows { get; }

    public IReadOnlyList<string> SummaryLines { get; }

    public string? RefusalCode { get; }
}
