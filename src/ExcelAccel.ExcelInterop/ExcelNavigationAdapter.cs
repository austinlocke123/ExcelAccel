using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Navigation;

namespace ExcelAccel.ExcelInterop;

public sealed class ExcelNavigationAdapter : INavigationPort
{
    private readonly Func<object> _getApplication;
    private readonly Action _verifyExcelThread;
    public ExcelNavigationAdapter(Func<object> getApplication, Action verifyExcelThread)
    {
        _getApplication = getApplication ?? throw new ArgumentNullException(nameof(getApplication));
        _verifyExcelThread = verifyExcelThread ?? throw new ArgumentNullException(nameof(verifyExcelThread));
    }

    public NavigationLocation CaptureLocation()
    {
        _verifyExcelThread();
        return ExcelComRetry.Execute(CaptureOnce);
    }

    public IReadOnlyList<string> GetVisibleWorksheetNames()
    {
        _verifyExcelThread();
        return ExcelComRetry.Execute(GetVisibleSheetsOnce);
    }

    public NavigationLocation ResolveTarget(NavigationTargetKind target)
    {
        _verifyExcelThread();
        return ExcelComRetry.Execute(() => ResolveOnce(target));
    }

    public bool TryNavigate(NavigationLocation target)
    {
        _verifyExcelThread();
        try { return ExcelComRetry.Execute(() => NavigateOnce(target)); }
        catch (CommandRefusedException) { return false; }
    }

    private NavigationLocation CaptureOnce()
    {
        object? applicationObject = null;
        object? workbookObject = null;
        object? sheetObject = null;
        object? selectionObject = null;
        try
        {
            applicationObject = _getApplication();
            ExcelCommandReadiness.RequireReady(applicationObject);
            dynamic app = applicationObject;
            workbookObject = app.ActiveWorkbook;
            sheetObject = app.ActiveSheet;
            selectionObject = app.Selection;
            if (workbookObject is null || sheetObject is null || selectionObject is null) throw MissingContext();
            return new NavigationLocation(WorkbookId(workbookObject), Convert.ToString(((dynamic)sheetObject).Name, CultureInfo.InvariantCulture) ?? string.Empty,
                Convert.ToString(((dynamic)selectionObject).Address[false, false, 1, false], CultureInfo.InvariantCulture) ?? string.Empty);
        }
        finally
        {
            ComRelease.Owned(selectionObject);
            ComRelease.Owned(sheetObject);
            ComRelease.Owned(workbookObject);
        }
    }

    private IReadOnlyList<string> GetVisibleSheetsOnce()
    {
        object? appObject = null;
        object? workbookObject = null;
        object? sheetsObject = null;
        var names = new List<string>();
        try
        {
            appObject = _getApplication();
            ExcelCommandReadiness.RequireReady(appObject);
            dynamic app = appObject;
            workbookObject = app.ActiveWorkbook;
            if (workbookObject is null) throw MissingContext();
            sheetsObject = ((dynamic)workbookObject).Worksheets;
            var count = Convert.ToInt32(((dynamic)sheetsObject).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                object? sheet = null;
                try
                {
                    sheet = ((dynamic)sheetsObject)[index];
                    if (Convert.ToInt32(((dynamic)sheet).Visible, CultureInfo.InvariantCulture) == -1)
                        names.Add(Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture) ?? string.Empty);
                }
                finally { ComRelease.Owned(sheet); }
            }
            return names;
        }
        finally
        {
            ComRelease.Owned(sheetsObject);
            ComRelease.Owned(workbookObject);
        }
    }

    private NavigationLocation ResolveOnce(NavigationTargetKind target)
    {
        var origin = CaptureOnce();
        object? appObject = null;
        object? sheetObject = null;
        object? selectionObject = null;
        object? usedObject = null;
        object? cellsObject = null;
        object? rowsObject = null;
        object? columnsObject = null;
        object? targetObject = null;
        try
        {
            appObject = _getApplication();
            dynamic app = appObject;
            sheetObject = app.ActiveSheet;
            selectionObject = app.Selection;
            if (sheetObject is null || selectionObject is null) throw MissingContext();
            dynamic sheet = sheetObject;
            dynamic selection = selectionObject;
            switch (target)
            {
                case NavigationTargetKind.A1: targetObject = sheet.Range["A1"]; break;
                case NavigationTargetKind.UsedFirst:
                    usedObject = sheet.UsedRange;
                    cellsObject = ((dynamic)usedObject).Cells;
                    targetObject = ((dynamic)cellsObject)[1, 1];
                    break;
                case NavigationTargetKind.UsedLast:
                    usedObject = sheet.UsedRange;
                    cellsObject = ((dynamic)usedObject).Cells;
                    rowsObject = ((dynamic)usedObject).Rows;
                    columnsObject = ((dynamic)usedObject).Columns;
                    targetObject = ((dynamic)cellsObject)[((dynamic)rowsObject).Count, ((dynamic)columnsObject).Count];
                    break;
                default: targetObject = selection.End[Direction(target)]; break;
            }
            var address = Convert.ToString(((dynamic)targetObject).Address[false, false, 1, false], CultureInfo.InvariantCulture) ?? string.Empty;
            return new NavigationLocation(origin.WorkbookId, origin.WorksheetName, address);
        }
        finally
        {
            ComRelease.Owned(targetObject);
            ComRelease.Owned(columnsObject);
            ComRelease.Owned(rowsObject);
            ComRelease.Owned(cellsObject);
            ComRelease.Owned(usedObject);
            ComRelease.Owned(selectionObject);
            ComRelease.Owned(sheetObject);
        }
    }

    private bool NavigateOnce(NavigationLocation target)
    {
        object? appObject = null;
        object? workbookObject = null;
        object? sheetsObject = null;
        object? sheetObject = null;
        object? rangeObject = null;
        try
        {
            appObject = _getApplication();
            ExcelCommandReadiness.RequireReady(appObject);
            dynamic app = appObject;
            workbookObject = app.ActiveWorkbook;
            if (workbookObject is null || WorkbookId(workbookObject) != target.WorkbookId) return false;
            sheetsObject = ((dynamic)workbookObject).Worksheets;
            sheetObject = ((dynamic)sheetsObject)[target.WorksheetName];
            ((dynamic)sheetObject).Activate();
            rangeObject = ((dynamic)sheetObject).Range[target.Address];
            ((dynamic)rangeObject).Select();
            return true;
        }
        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { return false; }
        catch (System.Runtime.InteropServices.COMException) { return false; }
        finally
        {
            ComRelease.Owned(rangeObject);
            ComRelease.Owned(sheetObject);
            ComRelease.Owned(sheetsObject);
            ComRelease.Owned(workbookObject);
        }
    }

    private static int Direction(NavigationTargetKind target)
    {
        switch (target)
        {
            case NavigationTargetKind.RegionEdgeUp: return -4162;
            case NavigationTargetKind.RegionEdgeDown: return -4121;
            case NavigationTargetKind.RegionEdgeLeft: return -4159;
            case NavigationTargetKind.RegionEdgeRight: return -4161;
            default: throw new ArgumentOutOfRangeException(nameof(target));
        }
    }

    private static string WorkbookId(object workbook)
    {
        var id = Convert.ToString(((dynamic)workbook).FullName, CultureInfo.InvariantCulture);
        return string.IsNullOrWhiteSpace(id) ? Convert.ToString(((dynamic)workbook).Name, CultureInfo.InvariantCulture) ?? string.Empty : id;
    }

    private static CommandRefusedException MissingContext() =>
        new CommandRefusedException(RefusalCodes.SelectionUnsupported, "Navigation requires an active workbook, worksheet, and cell selection.", "Open a workbook and select a cell.");
}
