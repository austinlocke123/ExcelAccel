using System;

namespace ExcelAccel.Core.Commands;

public sealed class SelectionSnapshot
{
    public SelectionSnapshot(SelectionContext context, long cellCount, bool? hasFormula, string numberFormat)
    {
        if (cellCount < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(cellCount));
        }

        Context = context ?? throw new ArgumentNullException(nameof(context));
        CellCount = cellCount;
        HasFormula = hasFormula;
        NumberFormat = numberFormat ?? string.Empty;
    }

    public SelectionContext Context { get; }

    public long CellCount { get; }

    public bool? HasFormula { get; }

    public string NumberFormat { get; }
}
