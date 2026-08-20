using System;

namespace ExcelAccel.Core.Auditing;

/// <summary>
/// The raw region a worksheet reports as used. This is an untrusted input: an
/// Excel worksheet routinely reports a used range far larger than its real
/// content because of stray formatting, so it is never a resource bound on its
/// own.
/// </summary>
public sealed class UsedRegionBounds
{
    public UsedRegionBounds(string worksheetName, int firstRow, int firstColumn, int rowCount, int columnCount)
    {
        WorksheetName = !string.IsNullOrWhiteSpace(worksheetName)
            ? worksheetName
            : throw new ArgumentException("A worksheet name is required.", nameof(worksheetName));
        FirstRow = firstRow;
        FirstColumn = firstColumn;
        RowCount = rowCount;
        ColumnCount = columnCount;
    }

    public string WorksheetName { get; }

    public int FirstRow { get; }

    public int FirstColumn { get; }

    public int RowCount { get; }

    public int ColumnCount { get; }

    public bool IsEmpty => RowCount <= 0 || ColumnCount <= 0;
}

/// <summary>
/// A bounded, banded read plan over one worksheet region. Every ceiling is
/// applied here, in pure code, so the bound does not depend on what Excel
/// reports. A region that cannot be read within the ceilings is refused rather
/// than partially guessed at.
/// </summary>
public sealed class DependentScanRegion
{
    /// <summary>Largest region a worksheet-scope scan will read at all.</summary>
    public const int MaximumRegionCells = 250_000;

    /// <summary>
    /// Largest single block read. This matches the bounded formula-block
    /// ceiling the Excel adapter already enforces.
    /// </summary>
    public const int MaximumBlockCells = 10_000;

    private DependentScanRegion(
        string worksheetName,
        int firstRow,
        int firstColumn,
        int rowCount,
        int columnCount,
        int blockRowCount)
    {
        WorksheetName = worksheetName;
        FirstRow = firstRow;
        FirstColumn = firstColumn;
        RowCount = rowCount;
        ColumnCount = columnCount;
        BlockRowCount = blockRowCount;
        BlockCount = ((rowCount - 1) / blockRowCount) + 1;
    }

    public string WorksheetName { get; }

    public int FirstRow { get; }

    public int FirstColumn { get; }

    public int RowCount { get; }

    public int ColumnCount { get; }

    public int BlockRowCount { get; }

    public int BlockCount { get; }

    public long CellCount => (long)RowCount * ColumnCount;

    public int LastRow => FirstRow + RowCount - 1;

    public int LastColumn => FirstColumn + ColumnCount - 1;

    /// <summary>
    /// Builds a bounded read plan, or explains why the region cannot be read.
    /// An empty region yields a plan with no blocks rather than a refusal.
    /// </summary>
    public static bool TryCreate(UsedRegionBounds bounds, out DependentScanRegion? region, out string? refusalCode, out string? message)
    {
        if (bounds is null) throw new ArgumentNullException(nameof(bounds));
        region = null;
        refusalCode = null;
        message = null;

        if (bounds.IsEmpty)
        {
            region = new DependentScanRegion(bounds.WorksheetName, 1, 1, 0, 0, 1);
            return true;
        }

        if (bounds.FirstRow < 1 || bounds.FirstColumn < 1 ||
            bounds.FirstRow > AuditAddress.MaximumRow || bounds.FirstColumn > AuditAddress.MaximumColumn ||
            (long)bounds.FirstRow + bounds.RowCount - 1 > AuditAddress.MaximumRow ||
            (long)bounds.FirstColumn + bounds.ColumnCount - 1 > AuditAddress.MaximumColumn)
        {
            refusalCode = AuditRefusalCodes.ScanRegionUnsupported;
            message = "The worksheet reported a used region outside the addressable grid.";
            return false;
        }

        if (bounds.ColumnCount > MaximumBlockCells)
        {
            refusalCode = AuditRefusalCodes.ScanRegionTooLarge;
            message = $"The used region spans {bounds.ColumnCount:N0} columns, which exceeds the {MaximumBlockCells:N0}-cell ceiling for a single row band. " + InflationGuidance;
            return false;
        }

        var cellCount = (long)bounds.RowCount * bounds.ColumnCount;
        if (cellCount > MaximumRegionCells)
        {
            refusalCode = AuditRefusalCodes.ScanRegionTooLarge;
            message = $"The used region covers {cellCount:N0} cells, which exceeds the {MaximumRegionCells:N0}-cell worksheet scan ceiling. " + InflationGuidance;
            return false;
        }

        var blockRowCount = Math.Max(1, MaximumBlockCells / bounds.ColumnCount);
        region = new DependentScanRegion(
            bounds.WorksheetName, bounds.FirstRow, bounds.FirstColumn, bounds.RowCount, bounds.ColumnCount, blockRowCount);
        return true;
    }

    private const string InflationGuidance =
        "Excel often reports a used range far larger than the real content because of stray formatting; " +
        "clear the unused formatting or scan a smaller worksheet.";

    /// <summary>Returns the rectangle covered by one banded read.</summary>
    public AuditRectangle Block(int index)
    {
        if (index < 0 || index >= BlockCount || RowCount == 0)
        {
            throw new ArgumentOutOfRangeException(nameof(index));
        }

        var firstRow = FirstRow + ((long)index * BlockRowCount);
        var lastRow = Math.Min(firstRow + BlockRowCount - 1, LastRow);
        return new AuditRectangle((int)firstRow, FirstColumn, (int)lastRow, LastColumn);
    }
}
