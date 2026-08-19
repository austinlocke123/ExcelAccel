using System;
using ExcelAccel.Core.Collaboration;

namespace ExcelAccel.Core.Commands;

public sealed class SelectionSnapshot
{
    public SelectionSnapshot(
        SelectionContext context,
        long cellCount,
        bool? hasFormula,
        string numberFormat,
        SelectionSafetyState? safety = null,
        WorkbookCollaborationState? collaboration = null)
    {
        if (cellCount < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(cellCount));
        }

        Context = context ?? throw new ArgumentNullException(nameof(context));
        CellCount = cellCount;
        HasFormula = hasFormula;
        NumberFormat = numberFormat ?? string.Empty;
        Safety = safety ?? SelectionSafetyState.Safe();
        Collaboration = collaboration ?? WorkbookCollaborationState.Unknown;
    }

    public SelectionContext Context { get; }

    public long CellCount { get; }

    public bool? HasFormula { get; }

    public string NumberFormat { get; }

    public SelectionSafetyState Safety { get; }

    public WorkbookCollaborationState Collaboration { get; }
}
