using System;

namespace ExcelAccel.Core.Commands;

public sealed class SelectionSafetyState
{
    public SelectionSafetyState(
        int areaCount,
        bool hasMergedCells,
        bool worksheetProtected,
        bool workbookReadOnly,
        bool hasLegacyArray,
        bool hasDynamicArraySpill,
        bool dynamicArraySpillCheckSupported = true)
    {
        if (areaCount < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(areaCount));
        }

        AreaCount = areaCount;
        HasMergedCells = hasMergedCells;
        WorksheetProtected = worksheetProtected;
        WorkbookReadOnly = workbookReadOnly;
        HasLegacyArray = hasLegacyArray;
        HasDynamicArraySpill = hasDynamicArraySpill;
        DynamicArraySpillCheckSupported = dynamicArraySpillCheckSupported;
    }

    public int AreaCount { get; }

    public bool HasMergedCells { get; }

    public bool WorksheetProtected { get; }

    public bool WorkbookReadOnly { get; }

    public bool HasLegacyArray { get; }

    public bool HasDynamicArraySpill { get; }

    public bool DynamicArraySpillCheckSupported { get; }

    public static SelectionSafetyState Safe() => new SelectionSafetyState(1, false, false, false, false, false);
}
