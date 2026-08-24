using System;
using System.Collections.Generic;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Compare;
using ExcelAccel.Core.Compare;

namespace ExcelAccel.ExcelInterop;

/// <summary>
/// Captures the two sides of a comparison from workbooks that are already open.
/// </summary>
/// <remarks>
/// Read-only by construction: it reads <c>Formula</c>, <c>Value2</c>, and
/// <c>NumberFormat</c> from an existing range, and never calls
/// <c>Workbooks.Open</c>, <c>Save</c>, <c>Close</c>, or <c>Calculate</c>. The
/// source side is remembered as an address, and re-read only when the comparison
/// runs, so a stale capture cannot present old content as current.
/// </remarks>
public sealed class ExcelComparisonAdapter : IComparisonPort
{
    private readonly Func<object> _getApplication;
    private readonly Action _verifyExcelThread;
    private readonly ComparisonAnchor? _source;

    public ExcelComparisonAdapter(Func<object> getApplication, Action verifyExcelThread, ComparisonAnchor? source = null)
    {
        _getApplication = getApplication ?? throw new ArgumentNullException(nameof(getApplication));
        _verifyExcelThread = verifyExcelThread ?? throw new ArgumentNullException(nameof(verifyExcelThread));
        _source = source;
    }

    /// <summary>Captures the current selection's identity for later comparison.</summary>
    public ComparisonAnchor CaptureAnchor()
    {
        _verifyExcelThread();
        return ExcelComRetry.Execute(() => WithSelection((workbook, worksheet, range) => new ComparisonAnchor(
            (string)((dynamic)workbook).Name,
            (string)((dynamic)worksheet).Name,
            (string)((dynamic)range).Address(false, false))));
    }

    public ComparisonBlock? CaptureSource()
    {
        if (_source is null) return null;
        _verifyExcelThread();
        return ExcelComRetry.Execute(() => ReadAnchor(_source));
    }

    public ComparisonBlock CaptureTarget()
    {
        _verifyExcelThread();
        return ExcelComRetry.Execute(() => WithSelection((workbook, worksheet, range) => ReadRange(
            (string)((dynamic)workbook).Name,
            (string)((dynamic)worksheet).Name,
            (string)((dynamic)range).Address(false, false),
            range)));
    }

    private ComparisonBlock ReadAnchor(ComparisonAnchor anchor)
    {
        object? applicationObject = null;
        object? workbooksObject = null;
        object? workbookObject = null;
        object? worksheetsObject = null;
        object? worksheetObject = null;
        object? rangeObject = null;
        try
        {
            applicationObject = _getApplication();
            ExcelCommandReadiness.RequireReady(applicationObject);
            workbooksObject = ((dynamic)applicationObject).Workbooks;

            // The source must still be open. Nothing is opened to satisfy a
            // comparison; a closed source is refused so a stale capture can never
            // be presented as current.
            try { workbookObject = ((dynamic)workbooksObject)[anchor.WorkbookId]; }
            catch (Exception)
            {
                throw new CommandRefusedException(
                    RefusalCodes.StaleContext,
                    $"The captured source workbook '{anchor.WorkbookId}' is no longer open.",
                    "Reopen it and capture the source again.");
            }

            worksheetsObject = ((dynamic)workbookObject).Worksheets;
            try { worksheetObject = ((dynamic)worksheetsObject)[anchor.WorksheetName]; }
            catch (Exception)
            {
                throw new CommandRefusedException(
                    RefusalCodes.StaleContext,
                    $"The captured source worksheet '{anchor.WorksheetName}' no longer exists.",
                    "Capture the source again.");
            }

            rangeObject = ((dynamic)worksheetObject).Range[anchor.Address];
            return ReadRange(anchor.WorkbookId, anchor.WorksheetName, anchor.Address, rangeObject);
        }
        finally
        {
            ComRelease.Owned(rangeObject);
            ComRelease.Owned(worksheetObject);
            ComRelease.Owned(worksheetsObject);
            ComRelease.Owned(workbookObject);
            ComRelease.Owned(workbooksObject);
            ComRelease.Owned(applicationObject);
        }
    }

    private T WithSelection<T>(Func<object, object, object, T> read)
    {
        object? applicationObject = null;
        object? workbookObject = null;
        object? worksheetObject = null;
        object? selectionObject = null;
        object? areasObject = null;
        object? areaObject = null;
        try
        {
            applicationObject = _getApplication();
            ExcelCommandReadiness.RequireReady(applicationObject);
            workbookObject = ((dynamic)applicationObject).ActiveWorkbook;
            worksheetObject = ((dynamic)applicationObject).ActiveSheet;
            selectionObject = ((dynamic)applicationObject).Selection;
            if (workbookObject is null || worksheetObject is null || selectionObject is null)
            {
                throw new CommandRefusedException(
                    RefusalCodes.SelectionUnsupported, "A cell selection is required.", "Select a range and retry.");
            }

            areasObject = ((dynamic)selectionObject).Areas;
            int areaCount = ((dynamic)areasObject).Count;
            if (areaCount != 1)
            {
                throw new CommandRefusedException(
                    RefusalCodes.MultiAreaUnsupported,
                    "Comparison requires one rectangular selection.",
                    "Select a single range and retry.");
            }

            areaObject = ((dynamic)areasObject)[1];
            return read(workbookObject, worksheetObject, areaObject);
        }
        finally
        {
            ComRelease.Owned(areaObject);
            ComRelease.Owned(areasObject);
            ComRelease.Owned(selectionObject);
            ComRelease.Owned(worksheetObject);
            ComRelease.Owned(workbookObject);
            ComRelease.Owned(applicationObject);
        }
    }

    private static ComparisonBlock ReadRange(string workbookId, string worksheetName, string address, object rangeObject)
    {
        object? rowsObject = null;
        object? columnsObject = null;
        object? cellsObject = null;
        try
        {
            rowsObject = ((dynamic)rangeObject).Rows;
            columnsObject = ((dynamic)rangeObject).Columns;
            int rowCount = ((dynamic)rowsObject).Count;
            int columnCount = ((dynamic)columnsObject).Count;
            if ((long)rowCount * columnCount > SameShapeComparer.MaximumCells)
            {
                throw new CommandRefusedException(
                    RefusalCodes.ResourceLimit,
                    $"The range covers {(long)rowCount * columnCount:N0} cells, beyond the qualified comparison bound.",
                    "Compare a smaller range.");
            }

            var cells = new List<ComparisonCell>(rowCount * columnCount);
            cellsObject = ((dynamic)rangeObject).Cells;
            for (var row = 1; row <= rowCount; row++)
            {
                for (var column = 1; column <= columnCount; column++)
                {
                    object? cellObject = null;
                    try
                    {
                        cellObject = ((dynamic)cellsObject)[row, column];
                        cells.Add(new ComparisonCell(
                            (string)((dynamic)cellObject).Address(false, false),
                            FormulaOf(cellObject),
                            ValueOf(cellObject),
                            NumberFormatOf(cellObject),
                            TextOf(cellObject)));
                    }
                    finally
                    {
                        ComRelease.Owned(cellObject);
                    }
                }
            }

            return new ComparisonBlock(workbookId, worksheetName, address, rowCount, columnCount, cells);
        }
        finally
        {
            ComRelease.Owned(cellsObject);
            ComRelease.Owned(columnsObject);
            ComRelease.Owned(rowsObject);
        }
    }

    private static string FormulaOf(object cellObject)
    {
        try
        {
            var formula = (string)((dynamic)cellObject).Formula ?? string.Empty;
            return formula.Length > 0 && formula[0] == '=' ? formula : string.Empty;
        }
        catch (Exception) { return string.Empty; }
    }

    private static string ValueOf(object cellObject)
    {
        try
        {
            object? value = ((dynamic)cellObject).Value2;
            return value is null
                ? string.Empty
                : Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
        }
        catch (Exception) { return string.Empty; }
    }

    private static string NumberFormatOf(object cellObject)
    {
        try { return (string)((dynamic)cellObject).NumberFormat ?? string.Empty; }
        catch (Exception) { return string.Empty; }
    }

    private static string TextOf(object cellObject)
    {
        try { return (string)((dynamic)cellObject).Text ?? string.Empty; }
        catch (Exception) { return string.Empty; }
    }
}

/// <summary>
/// The identity of a captured comparison side. Only the address is remembered;
/// the content is re-read when the comparison runs.
/// </summary>
public sealed class ComparisonAnchor
{
    public ComparisonAnchor(string workbookId, string worksheetName, string address)
    {
        WorkbookId = workbookId;
        WorksheetName = worksheetName;
        Address = address;
    }

    public string WorkbookId { get; }
    public string WorksheetName { get; }
    public string Address { get; }

    public override string ToString() => WorkbookId + "!" + WorksheetName + "!" + Address;
}
