using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAccel.Application.AutoColor;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Commands;
using ExcelAccel.Core.Reliability;

namespace ExcelAccel.ExcelInterop;

/// <summary>
/// Reads the cells a recolour considers and writes their font colours.
/// </summary>
/// <remarks>
/// Writes are grouped: every cell taking the same colour is combined into one
/// range and coloured in a single COM call, so a recolour costs roughly one write
/// per distinct colour per batch rather than one per cell. Batches are bounded
/// because a range built from too many discontiguous areas is itself slow.
///
/// Reads are still per cell, because kind, formula, and font colour are needed
/// together and a cell's error state is only visible through its displayed text.
/// That is the slower half and the obvious place to optimise next; it is bounded
/// by the planner's cell ceiling, which is why it is acceptable for now.
/// </remarks>
public sealed class ExcelAutoColorAdapter : IAutoColorPort
{
    /// <summary>Cells combined into one Union before it is flushed.</summary>
    private const int UnionBatchSize = 256;

    private readonly Func<object> _getApplication;
    private readonly Action _verifyExcelThread;
    private readonly ExcelSelectionAdapter _selection;

    public ExcelAutoColorAdapter(Func<object> getApplication, Action verifyExcelThread)
    {
        _getApplication = getApplication ?? throw new ArgumentNullException(nameof(getApplication));
        _verifyExcelThread = verifyExcelThread ?? throw new ArgumentNullException(nameof(verifyExcelThread));
        _selection = new ExcelSelectionAdapter(getApplication, verifyExcelThread);
    }

    public SelectionSnapshot CaptureSelection() => _selection.CaptureSelection();

    public IReadOnlyList<AutoColorCellSnapshot> CaptureCells(AutoColorScope scope)
    {
        _verifyExcelThread();
        return ExcelComRetry.Execute(() => CaptureCellsOnce(scope));
    }

    public void WriteFontColors(IEnumerable<AutoColorChange> changes)
    {
        _verifyExcelThread();
        if (changes is null) throw new ArgumentNullException(nameof(changes));
        var ordered = changes.ToArray();
        if (ordered.Length == 0) return;
        ExcelComRetry.Execute(() => WriteFontColorsOnce(ordered));
    }

    /// <summary>
    /// Undo reads and writes the whole block through this port, so the block
    /// property is handled here and everything else defers to the selection
    /// adapter's existing property handling.
    /// </summary>
    public bool TryRead(SelectionContext target, string propertyId, out string value)
        => TryReadCore(target, propertyId, null, out value);

    public bool TryRead(
        SelectionContext target,
        string propertyId,
        string referenceValue,
        out string value)
        => TryReadCore(target, propertyId, referenceValue, out value);

    private bool TryReadCore(
        SelectionContext target,
        string propertyId,
        string? referenceValue,
        out string value)
    {
        if (!string.Equals(propertyId, FontColorBlock.ReceiptPropertyId, StringComparison.Ordinal))
        {
            return _selection.TryRead(target, propertyId, out value);
        }

        value = string.Empty;
        try
        {
            // A font-colour receipt contains only cells AutoColor changed. Use
            // that value as the exact address manifest; reading the whole
            // original selection would include unchanged cells and make a valid
            // undo look stale.
            var addresses = string.IsNullOrEmpty(referenceValue)
                ? null
                : FontColorBlock.Deserialize(referenceValue)
                    .Select(cell => cell.Key)
                    .ToArray();
            var cells = ReadBlock(target, addresses);
            value = FontColorBlock.Serialize(cells);
            return true;
        }
        catch (Exception)
        {
            return false;
        }
    }

    public bool TryWrite(SelectionContext target, string propertyId, string value)
    {
        if (!string.Equals(propertyId, FontColorBlock.ReceiptPropertyId, StringComparison.Ordinal))
        {
            return _selection.TryWrite(target, propertyId, value);
        }

        try
        {
            var cells = FontColorBlock.Deserialize(value);
            WriteFontColors(cells.Select(cell => new AutoColorChange(
                cell.Key, AutoColorCategory.Unsupported, string.Empty, cell.Value)));
            return true;
        }
        catch (Exception)
        {
            return false;
        }
    }

    private IReadOnlyList<KeyValuePair<string, string>> ReadBlock(
        SelectionContext target,
        IReadOnlyList<string>? addresses = null)
    {
        object? applicationObject = null;
        object? workbookObject = null;
        object? worksheetsObject = null;
        object? worksheetObject = null;
        object? rangeObject = null;
        try
        {
            applicationObject = _getApplication();
            ExcelCommandReadiness.RequireReady(applicationObject);
            workbookObject = ((dynamic)applicationObject).ActiveWorkbook;
            if (workbookObject is null)
            {
                throw new CommandRefusedException(
                    RefusalCodes.SelectionUnsupported, "An active workbook is required.", "Open the planned workbook and retry.");
            }

            worksheetsObject = ((dynamic)workbookObject).Worksheets;
            worksheetObject = ((dynamic)worksheetsObject)[target.WorksheetName];
            if (addresses is not null)
            {
                return ReadAddresses(worksheetObject, addresses);
            }

            rangeObject = ((dynamic)worksheetObject).Range[target.Address];
            return ReadRange(rangeObject);
        }
        finally
        {
            ComRelease.Owned(rangeObject);
            ComRelease.Owned(worksheetObject);
            ComRelease.Owned(worksheetsObject);
            ComRelease.Owned(workbookObject);
            ComRelease.Owned(applicationObject);
        }
    }

    private static IReadOnlyList<KeyValuePair<string, string>> ReadRange(object rangeObject)
    {
        var result = new List<KeyValuePair<string, string>>();
        object? cellsObject = null;
        try
        {
            cellsObject = ((dynamic)rangeObject).Cells;
            int count = ((dynamic)cellsObject).Count;
            for (var index = 1; index <= count && result.Count < FontColorBlock.MaximumCells; index++)
            {
                object? cellObject = null;
                object? fontObject = null;
                try
                {
                    cellObject = ((dynamic)cellsObject)[index];
                    fontObject = ((dynamic)cellObject).Font;
                    result.Add(new KeyValuePair<string, string>(
                        (string)((dynamic)cellObject).Address(false, false),
                        ColorToHex(((dynamic)fontObject).Color)));
                }
                finally
                {
                    ComRelease.Owned(fontObject);
                    ComRelease.Owned(cellObject);
                }
            }
        }
        finally
        {
            ComRelease.Owned(cellsObject);
        }

        return result;
    }

    private static IReadOnlyList<KeyValuePair<string, string>> ReadAddresses(
        object worksheetObject,
        IReadOnlyList<string> addresses)
    {
        var result = new List<KeyValuePair<string, string>>(addresses.Count);
        foreach (var address in addresses)
        {
            object? cellObject = null;
            object? fontObject = null;
            try
            {
                cellObject = ((dynamic)worksheetObject).Range[address];
                fontObject = ((dynamic)cellObject).Font;
                result.Add(new KeyValuePair<string, string>(
                    (string)((dynamic)cellObject).Address(false, false),
                    ColorToHex(((dynamic)fontObject).Color)));
            }
            finally
            {
                ComRelease.Owned(fontObject);
                ComRelease.Owned(cellObject);
            }
        }

        return result;
    }

    private IReadOnlyList<AutoColorCellSnapshot> CaptureCellsOnce(AutoColorScope scope)
    {
        object? applicationObject = null;
        object? workbookObject = null;
        object? targetObject = null;
        try
        {
            applicationObject = _getApplication();
            ExcelCommandReadiness.RequireReady(applicationObject);
            workbookObject = ((dynamic)applicationObject).ActiveWorkbook;
            if (workbookObject is null)
            {
                throw new CommandRefusedException(
                    RefusalCodes.SelectionUnsupported, "An active workbook is required.", "Open a workbook and retry.");
            }

            if (scope == AutoColorScope.Worksheet)
            {
                object? sheetObject = null;
                try
                {
                    sheetObject = ((dynamic)applicationObject).ActiveSheet;
                    targetObject = ((dynamic)sheetObject).UsedRange;
                }
                finally
                {
                    ComRelease.Owned(sheetObject);
                }
            }
            else
            {
                targetObject = ((dynamic)applicationObject).Selection;
            }

            if (targetObject is null)
            {
                throw new CommandRefusedException(
                    RefusalCodes.SelectionUnsupported, "A target range is required.", "Select cells and retry.");
            }

            return CaptureRange(targetObject);
        }
        finally
        {
            ComRelease.Owned(targetObject);
            ComRelease.Owned(workbookObject);
            ComRelease.Owned(applicationObject);
        }
    }

    private static IReadOnlyList<AutoColorCellSnapshot> CaptureRange(object rangeObject)
    {
        var snapshots = new List<AutoColorCellSnapshot>();
        object? cellsObject = null;
        try
        {
            cellsObject = ((dynamic)rangeObject).Cells;
            int count = ((dynamic)cellsObject).Count;
            if (count > AutoColorPlanner.MaximumPlannableCells)
            {
                throw new CommandRefusedException(
                    RefusalCodes.ResourceLimit,
                    $"The target holds {count:N0} cells, beyond the qualified planning bound.",
                    "Select a smaller range and retry.");
            }

            for (var index = 1; index <= count; index++)
            {
                object? cellObject = null;
                object? fontObject = null;
                try
                {
                    cellObject = ((dynamic)cellsObject)[index];
                    fontObject = ((dynamic)cellObject).Font;
                    snapshots.Add(new AutoColorCellSnapshot(
                        (string)((dynamic)cellObject).Address(false, false),
                        ScalarKindOf(cellObject),
                        FormulaOf(cellObject),
                        ColorToHex(((dynamic)fontObject).Color)));
                }
                finally
                {
                    ComRelease.Owned(fontObject);
                    ComRelease.Owned(cellObject);
                }
            }
        }
        finally
        {
            ComRelease.Owned(cellsObject);
        }

        return snapshots;
    }

    private static string FormulaOf(object cellObject)
    {
        try
        {
            var formula = (string)((dynamic)cellObject).Formula ?? string.Empty;
            return formula.Length > 0 && formula[0] == '=' ? formula : string.Empty;
        }
        catch (Exception)
        {
            return string.Empty;
        }
    }

    private static CellScalarKind ScalarKindOf(object cellObject)
    {
        object? value;
        try { value = ((dynamic)cellObject).Value2; }
        catch (Exception) { return CellScalarKind.Unsupported; }

        if (value is null) return CellScalarKind.Empty;

        // Excel surfaces a cell error as an int error code through Value2, and a
        // formula that evaluated to an error must classify as an error rather
        // than as the number its code happens to be.
        try
        {
            if ((bool)((dynamic)cellObject).HasFormula || value is string)
            {
                object? text;
                try { text = ((dynamic)cellObject).Text; }
                catch (Exception) { text = null; }
                if (text is string display && display.StartsWith("#", StringComparison.Ordinal))
                {
                    return CellScalarKind.Error;
                }
            }
        }
        catch (Exception)
        {
        }

        switch (value)
        {
            case bool _: return CellScalarKind.Boolean;
            case string text: return text.Length == 0 ? CellScalarKind.Empty : CellScalarKind.Text;
            case double _:
            case int _:
            case long _:
            case decimal _:
                return CellScalarKind.Number;
            default:
                return CellScalarKind.Unsupported;
        }
    }

    private void WriteFontColorsOnce(IReadOnlyList<AutoColorChange> changes)
    {
        object? applicationObject = null;
        object? worksheetObject = null;
        try
        {
            applicationObject = _getApplication();
            ExcelCommandReadiness.RequireReady(applicationObject);
            worksheetObject = ((dynamic)applicationObject).ActiveSheet;
            if (worksheetObject is null)
            {
                throw new CommandRefusedException(
                    RefusalCodes.SelectionUnsupported, "An active worksheet is required.", "Open a worksheet and retry.");
            }

            var application = applicationObject;
            var worksheet = worksheetObject;
            ApplicationStateGuard.Run(
                new ExcelApplicationStateAdapter(applicationObject),
                ApplicationStateChangeSet.PropertyMutation(),
                () =>
                {
                    foreach (var group in changes.GroupBy(change => change.AfterColor, StringComparer.OrdinalIgnoreCase))
                    {
                        WriteOneColor(application, worksheet, group.Select(change => change.Address).ToArray(), group.Key);
                    }
                });
        }
        finally
        {
            ComRelease.Owned(worksheetObject);
            ComRelease.Owned(applicationObject);
        }
    }

    /// <summary>
    /// Colours every address taking one colour, in bounded Union batches so the
    /// cost is one write per batch rather than one per cell.
    /// </summary>
    private static void WriteOneColor(object applicationObject, object worksheetObject, IReadOnlyList<string> addresses, string color)
    {
        var oleColor = HexToOle(color);
        for (var start = 0; start < addresses.Count; start += UnionBatchSize)
        {
            object? batchObject = null;
            object? fontObject = null;
            try
            {
                var batch = addresses.Skip(start).Take(UnionBatchSize).ToArray();
                batchObject = BuildUnion(applicationObject, worksheetObject, batch);
                if (batchObject is null) continue;
                fontObject = ((dynamic)batchObject).Font;
                ((dynamic)fontObject).Color = oleColor;
            }
            finally
            {
                ComRelease.Owned(fontObject);
                ComRelease.Owned(batchObject);
            }
        }
    }

    private static object? BuildUnion(object applicationObject, object worksheetObject, IReadOnlyList<string> addresses)
    {
        // Excel's own address grammar accepts a comma-separated list up to a
        // length limit, and one Range call is cheaper than repeated Union calls.
        var combined = string.Join(",", addresses);
        if (combined.Length <= 250)
        {
            return ((dynamic)worksheetObject).Range[combined];
        }

        object? accumulated = null;
        try
        {
            foreach (var address in addresses)
            {
                object? next = null;
                try
                {
                    next = ((dynamic)worksheetObject).Range[address];
                    if (accumulated is null)
                    {
                        accumulated = next;
                        next = null;
                        continue;
                    }

                    var merged = (object)((dynamic)applicationObject).Union(accumulated, next);
                    ComRelease.Owned(accumulated);
                    accumulated = merged;
                }
                finally
                {
                    ComRelease.Owned(next);
                }
            }

            var result = accumulated;
            accumulated = null;
            return result;
        }
        finally
        {
            // Ownership transfers to the caller only on success. If a later
            // Range or Union call fails, release the partial union here.
            ComRelease.Owned(accumulated);
        }
    }

    private static string ColorToHex(object value)
    {
        if (value is null) return "#000000";
        var color = Convert.ToInt32(value, CultureInfo.InvariantCulture);
        var red = color & 0xFF;
        var green = (color >> 8) & 0xFF;
        var blue = (color >> 16) & 0xFF;
        return $"#{red:X2}{green:X2}{blue:X2}";
    }

    private static int HexToOle(string value)
    {
        if (value is null || value.Length != 7 || value[0] != '#')
        {
            throw new ArgumentException("A #RRGGBB color is required.", nameof(value));
        }

        var red = int.Parse(value.Substring(1, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        var green = int.Parse(value.Substring(3, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        var blue = int.Parse(value.Substring(5, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        return red | (green << 8) | (blue << 16);
    }
}
