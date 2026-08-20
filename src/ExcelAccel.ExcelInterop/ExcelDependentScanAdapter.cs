using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using ExcelAccel.Application.Auditing;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.ExcelInterop;

/// <summary>
/// Read-only worksheet scan boundary for direct dependents. It reports what the
/// worksheet says and reads exactly the bands it is asked for; every ceiling and
/// the banding plan are decided in pure code by the coordinator. It writes
/// nothing, selects nothing, and opens no external workbook.
/// </summary>
public sealed class ExcelDependentScanAdapter : IDependentScanPort
{
    /// <summary>Ceiling on defined names enumerated for one scan.</summary>
    public const int MaximumNames = 4_096;

    private readonly ExcelSelectionAdapter _selection;
    private readonly Func<object> _getApplication;
    private readonly Action _verifyExcelThread;

    public ExcelDependentScanAdapter(Func<object> getApplication, Action verifyExcelThread)
    {
        _getApplication = getApplication ?? throw new ArgumentNullException(nameof(getApplication));
        _verifyExcelThread = verifyExcelThread ?? throw new ArgumentNullException(nameof(verifyExcelThread));
        _selection = new ExcelSelectionAdapter(_getApplication, _verifyExcelThread);
    }

    public AuditCellIdentity CaptureTarget()
    {
        var snapshot = _selection.CaptureSelection();
        if (snapshot.Safety.AreaCount != 1)
        {
            throw new CommandRefusedException(
                RefusalCodes.MultiAreaUnsupported,
                "Direct dependents require one rectangular selection.",
                "Select one cell or one rectangular range and retry.");
        }

        var context = snapshot.Context;
        if (!AuditAddress.TryParse(context.Address, out _))
        {
            throw new CommandRefusedException(
                RefusalCodes.SelectionUnsupported,
                "Direct dependents are qualified only for a single A1 cell or rectangular range target.",
                "Select one cell or one rectangular range and retry.");
        }

        return new AuditCellIdentity(context.WorkbookId, context.WorksheetName, context.Address);
    }

    /// <summary>Every visible worksheet, in workbook order.</summary>
    public IReadOnlyList<string> CaptureWorksheetNames()
    {
        _verifyExcelThread();
        return ExcelComRetry.Execute(CaptureWorksheetNamesOnce);
    }

    public UsedRegionBounds CaptureUsedRegion(string worksheetName)
    {
        if (string.IsNullOrWhiteSpace(worksheetName)) throw new ArgumentException("A worksheet name is required.", nameof(worksheetName));
        _verifyExcelThread();
        return ExcelComRetry.Execute(() => CaptureUsedRegionOnce(worksheetName));
    }

    public IReadOnlyList<AuditFormulaCell> CaptureBlock(string worksheetName, AuditRectangle band)
    {
        if (string.IsNullOrWhiteSpace(worksheetName)) throw new ArgumentException("A worksheet name is required.", nameof(worksheetName));
        var workbookId = CaptureTarget().WorkbookId;
        var block = _selection.CaptureFormulaBlock(new SelectionContext(workbookId, worksheetName, band.ToString()));
        var formulas = new List<AuditFormulaCell>();
        for (var row = 0; row < block.Contents.RowCount; row++)
        {
            for (var column = 0; column < block.Contents.ColumnCount; column++)
            {
                var cell = block.Contents[row, column];
                if (!cell.IsFormula) continue;
                formulas.Add(new AuditFormulaCell(
                    new AuditCellIdentity(
                        workbookId,
                        worksheetName,
                        AuditAddress.Cell(block.FirstRow + row, block.FirstColumn + column)),
                    cell.InvariantValue));
            }
        }

        return formulas;
    }

    public IReadOnlyList<AuditNameBinding> CaptureNames(DependentScanScope scope)
    {
        if (scope is null) throw new ArgumentNullException(nameof(scope));
        _verifyExcelThread();
        try
        {
            return ExcelComRetry.Execute(() => CaptureNamesOnce(scope));
        }
        catch (COMException)
        {
            return Array.Empty<AuditNameBinding>();
        }
    }

    private UsedRegionBounds CaptureUsedRegionOnce(string worksheetName)
    {
        object? applicationObject = null;
        object? workbookObject = null;
        object? worksheetsObject = null;
        object? worksheetObject = null;
        object? usedObject = null;
        object? rowsObject = null;
        object? columnsObject = null;
        try
        {
            applicationObject = _getApplication();
            ExcelCommandReadiness.RequireReady(applicationObject);
            workbookObject = ((dynamic)applicationObject).ActiveWorkbook;
            if (workbookObject is null) throw MissingWorkbook();
            worksheetsObject = ((dynamic)workbookObject).Worksheets;
            worksheetObject = ((dynamic)worksheetsObject)[worksheetName];
            usedObject = ((dynamic)worksheetObject).UsedRange;
            if (usedObject is null) return new UsedRegionBounds(worksheetName, 1, 1, 0, 0);
            rowsObject = ((dynamic)usedObject).Rows;
            columnsObject = ((dynamic)usedObject).Columns;
            return new UsedRegionBounds(
                worksheetName,
                Convert.ToInt32(((dynamic)usedObject).Row, CultureInfo.InvariantCulture),
                Convert.ToInt32(((dynamic)usedObject).Column, CultureInfo.InvariantCulture),
                Convert.ToInt32(((dynamic)rowsObject).Count, CultureInfo.InvariantCulture),
                Convert.ToInt32(((dynamic)columnsObject).Count, CultureInfo.InvariantCulture));
        }
        finally
        {
            ComRelease.Owned(columnsObject);
            ComRelease.Owned(rowsObject);
            ComRelease.Owned(usedObject);
            ComRelease.Owned(worksheetObject);
            ComRelease.Owned(worksheetsObject);
            ComRelease.Owned(workbookObject);
        }
    }

    private IReadOnlyList<string> CaptureWorksheetNamesOnce()
    {
        var names = new List<string>();
        object? applicationObject = null;
        object? workbookObject = null;
        object? worksheetsObject = null;
        try
        {
            applicationObject = _getApplication();
            ExcelCommandReadiness.RequireReady(applicationObject);
            workbookObject = ((dynamic)applicationObject).ActiveWorkbook;
            if (workbookObject is null) throw MissingWorkbook();
            worksheetsObject = ((dynamic)workbookObject).Worksheets;
            var count = Convert.ToInt32(((dynamic)worksheetsObject).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                object? worksheetObject = null;
                try
                {
                    worksheetObject = ((dynamic)worksheetsObject)[index];
                    var name = Convert.ToString(((dynamic)worksheetObject).Name, CultureInfo.InvariantCulture);
                    if (!string.IsNullOrWhiteSpace(name)) names.Add(name!);
                }
                finally
                {
                    ComRelease.Owned(worksheetObject);
                }
            }

            return names;
        }
        finally
        {
            ComRelease.Owned(worksheetsObject);
            ComRelease.Owned(workbookObject);
        }
    }

    private IReadOnlyList<AuditNameBinding> CaptureNamesOnce(DependentScanScope scope)
    {
        var bindings = new List<AuditNameBinding>();
        object? applicationObject = null;
        object? workbookObject = null;
        object? worksheetsObject = null;
        object? worksheetObject = null;
        try
        {
            applicationObject = _getApplication();
            ExcelCommandReadiness.RequireReady(applicationObject);
            workbookObject = ((dynamic)applicationObject).ActiveWorkbook;
            if (workbookObject is null) throw MissingWorkbook();
            var anchorWorksheet = scope.WorksheetName ?? CaptureTarget().WorksheetName;
            Collect(((dynamic)workbookObject).Names, AuditNameScope.Workbook, scope, anchorWorksheet, bindings);
            worksheetsObject = ((dynamic)workbookObject).Worksheets;
            worksheetObject = ((dynamic)worksheetsObject)[scope.WorksheetName ?? CaptureTarget().WorksheetName];
            Collect(((dynamic)worksheetObject).Names, AuditNameScope.Worksheet, scope, anchorWorksheet, bindings);
            return bindings;
        }
        finally
        {
            ComRelease.Owned(worksheetObject);
            ComRelease.Owned(worksheetsObject);
            ComRelease.Owned(workbookObject);
        }
    }

    private static void Collect(
        object? namesObject,
        AuditNameScope nameScope,
        DependentScanScope scope,
        string anchorWorksheet,
        ICollection<AuditNameBinding> bindings)
    {
        if (namesObject is null) return;
        try
        {
            var count = Convert.ToInt32(((dynamic)namesObject).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count && bindings.Count < MaximumNames; index++)
            {
                object? nameObject = null;
                try
                {
                    nameObject = ((dynamic)namesObject).Item(index);
                    var binding = TryBind(nameObject, nameScope, scope, anchorWorksheet);
                    if (binding is not null) bindings.Add(binding);
                }
                catch (COMException)
                {
                    // A name that cannot be read stays absent, which the index
                    // already reports as an explicit coverage gap.
                }
                finally
                {
                    ComRelease.Owned(nameObject);
                }
            }
        }
        finally
        {
            ComRelease.Owned(namesObject);
        }
    }

    /// <summary>
    /// Binds only a simple local single-target definition. Anything else is
    /// omitted so the reverse index reports it as a coverage gap rather than
    /// guessing. External definitions are never followed.
    /// </summary>
    private static AuditNameBinding? TryBind(object nameObject, AuditNameScope nameScope, DependentScanScope scope, string anchorWorksheet)
    {
        var name = Convert.ToString(((dynamic)nameObject).Name, CultureInfo.InvariantCulture);
        if (string.IsNullOrWhiteSpace(name)) return null;
        if (nameScope == AuditNameScope.Worksheet)
        {
            var separator = name!.LastIndexOf('!');
            if (separator >= 0) name = name.Substring(separator + 1);
            if (string.IsNullOrWhiteSpace(name)) return null;
        }

        var refersTo = Convert.ToString(((dynamic)nameObject).RefersTo, CultureInfo.InvariantCulture);
        if (string.IsNullOrWhiteSpace(refersTo)) return null;
        var anchor = new AuditCellIdentity(scope.WorkbookId, anchorWorksheet, "A1");
        var plan = new DirectPrecedentAnalyzer().CreateCapturePlan(anchor, refersTo!);
        if (plan.LocalTargets.Count != 1 || plan.NameCandidates.Count != 0) return null;
        return new AuditNameBinding(
            name!,
            nameScope,
            plan.LocalTargets[0],
            nameScope == AuditNameScope.Worksheet ? anchorWorksheet : null);
    }

    private static CommandRefusedException MissingWorkbook() => new CommandRefusedException(
        RefusalCodes.SelectionUnsupported,
        "An open workbook is required for a dependent scan.",
        "Open a workbook and retry.");
}
