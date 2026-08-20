using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAccel.Core.Auditing;

namespace ExcelAccel.Core.ModelCheck;

/// <summary>
/// A deterministic peer region: a contiguous run of cells down one column of one
/// worksheet. Grouping is purely spatial, so it does not depend on rule order,
/// and it never infers meaning from labels or content.
/// </summary>
public sealed class PeerRegion
{
    internal PeerRegion(string worksheetName, int column, IReadOnlyList<ModelCheckCell> cells)
    {
        WorksheetName = worksheetName;
        Column = column;
        Cells = cells;
    }

    public string WorksheetName { get; }

    public int Column { get; }

    /// <summary>Cells in ascending row order.</summary>
    public IReadOnlyList<ModelCheckCell> Cells { get; }

    public string Label =>
        WorksheetName + "!" + AuditAddress.Cell(FirstRow, Column) + ":" + AuditAddress.Cell(LastRow, Column);

    public int FirstRow { get; internal set; }

    public int LastRow { get; internal set; }

    /// <summary>
    /// Splits the snapshot into contiguous column runs. A blank cell breaks a run
    /// when the configuration says so; a gap in row numbers always breaks one.
    /// </summary>
    public static IReadOnlyList<PeerRegion> Build(ModelCheckSnapshot snapshot, ModelCheckConfiguration configuration)
    {
        if (snapshot is null) throw new ArgumentNullException(nameof(snapshot));
        if (configuration is null) throw new ArgumentNullException(nameof(configuration));

        var positioned = new List<(string Sheet, int Column, int Row, ModelCheckCell Cell)>();
        foreach (var cell in snapshot.Cells)
        {
            if (!AuditAddress.TryParse(cell.Identity.Address, out var rectangle) || !rectangle.IsSingleCell) continue;
            positioned.Add((cell.Identity.WorksheetName, rectangle.FirstColumn, rectangle.FirstRow, cell));
        }

        var regions = new List<PeerRegion>();
        foreach (var group in positioned
            .GroupBy(item => new { item.Sheet, item.Column })
            .OrderBy(group => group.Key.Sheet, StringComparer.OrdinalIgnoreCase)
            .ThenBy(group => group.Key.Column))
        {
            var run = new List<ModelCheckCell>();
            var firstRow = 0;
            var previousRow = int.MinValue;
            foreach (var item in group.OrderBy(item => item.Row))
            {
                var breaksRun = item.Row != previousRow + 1 ||
                    (configuration.TreatBlanksAsPeerBreaks && item.Cell.IsBlank);
                if (breaksRun && run.Count != 0)
                {
                    regions.Add(Close(group.Key.Sheet, group.Key.Column, run, firstRow, previousRow));
                    run = new List<ModelCheckCell>();
                }

                if (configuration.TreatBlanksAsPeerBreaks && item.Cell.IsBlank)
                {
                    previousRow = item.Row;
                    continue;
                }

                if (run.Count == 0) firstRow = item.Row;
                run.Add(item.Cell);
                previousRow = item.Row;
            }

            if (run.Count != 0) regions.Add(Close(group.Key.Sheet, group.Key.Column, run, firstRow, previousRow));
        }

        return regions;
    }

    private static PeerRegion Close(string sheet, int column, IReadOnlyList<ModelCheckCell> run, int firstRow, int lastRow) =>
        new PeerRegion(sheet, column, run) { FirstRow = firstRow, LastRow = lastRow };

    /// <summary>
    /// The most common non-null value and its count, breaking ties by ordinal
    /// comparison so the baseline never depends on enumeration order.
    /// </summary>
    internal static (string? Baseline, int Count) Majority(IEnumerable<string?> values)
    {
        var counts = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var value in values)
        {
            if (value is null) continue;
            counts[value] = counts.TryGetValue(value, out var count) ? count + 1 : 1;
        }

        if (counts.Count == 0) return (null, 0);
        var best = counts
            .OrderByDescending(entry => entry.Value)
            .ThenBy(entry => entry.Key, StringComparer.Ordinal)
            .First();
        return (best.Key, best.Value);
    }

    internal static string RowLabel(ModelCheckCell cell) =>
        AuditAddress.TryParse(cell.Identity.Address, out var rectangle)
            ? rectangle.FirstRow.ToString(CultureInfo.InvariantCulture)
            : cell.Identity.Address;
}
