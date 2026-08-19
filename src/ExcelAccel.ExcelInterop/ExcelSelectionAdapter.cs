using System;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Commands;
using ExcelAccel.Core.Reliability;

namespace ExcelAccel.ExcelInterop;

public sealed class ExcelSelectionAdapter : ISelectionPort
{
    private readonly Func<object> _getApplication;
    private readonly Action _verifyExcelThread;

    public ExcelSelectionAdapter(Func<object> getApplication, Action verifyExcelThread)
    {
        _getApplication = getApplication ?? throw new ArgumentNullException(nameof(getApplication));
        _verifyExcelThread = verifyExcelThread ?? throw new ArgumentNullException(nameof(verifyExcelThread));
    }

    public SelectionSnapshot CaptureSelection()
    {
        _verifyExcelThread();
        return ExcelComRetry.Execute(CaptureSelectionOnce);
    }

    private SelectionSnapshot CaptureSelectionOnce()
    {
        object? applicationObject = null;
        object? workbookObject = null;
        object? worksheetObject = null;
        object? selectionObject = null;
        object? areasObject = null;

        try
        {
            applicationObject = _getApplication();
            ExcelCommandReadiness.RequireReady(applicationObject);
            dynamic application = applicationObject;
            workbookObject = application.ActiveWorkbook;
            worksheetObject = application.ActiveSheet;
            selectionObject = application.Selection;

            if (workbookObject is null || worksheetObject is null || selectionObject is null)
            {
                throw new CommandRefusedException(
                    RefusalCodes.SelectionUnsupported,
                    "An open workbook and a cell range selection are required.",
                    "Open a workbook and select one rectangular cell range.");
            }

            dynamic workbook = workbookObject;
            dynamic worksheet = worksheetObject;
            dynamic selection = selectionObject;

            string workbookId = Convert.ToString(workbook.FullName, CultureInfo.InvariantCulture) ?? string.Empty;
            if (string.IsNullOrWhiteSpace(workbookId))
            {
                workbookId = Convert.ToString(workbook.Name, CultureInfo.InvariantCulture) ?? string.Empty;
            }

            string worksheetName = Convert.ToString(worksheet.Name, CultureInfo.InvariantCulture) ?? string.Empty;
            string address = Convert.ToString(selection.Address[false, false, 1, false], CultureInfo.InvariantCulture) ?? string.Empty;
            long cellCount = Convert.ToInt64(selection.CountLarge, CultureInfo.InvariantCulture);
            bool? hasFormula = ConvertFormulaState(selection.HasFormula);
            string numberFormat = Convert.ToString(selection.NumberFormat, CultureInfo.InvariantCulture) ?? "(mixed)";
            areasObject = selection.Areas;
            dynamic areas = areasObject;
            int areaCount = Convert.ToInt32(areas.Count, CultureInfo.InvariantCulture);
            bool hasMergedCells = ConvertUnsafeBoolean(selection.MergeCells);
            bool worksheetProtected = Convert.ToBoolean(worksheet.ProtectContents, CultureInfo.InvariantCulture);
            bool workbookReadOnly = Convert.ToBoolean(workbook.ReadOnly, CultureInfo.InvariantCulture);
            bool hasLegacyArray = ConvertUnsafeBoolean(selection.HasArray);
            bool spillCheckSupported = TryReadOptionalUnsafeBoolean(selectionObject, "HasSpill", out var hasDynamicArraySpill);
            var collaboration = ExcelWorkbookCollaborationAdapter.Capture(workbookObject);

            return new SelectionSnapshot(
                new SelectionContext(workbookId, worksheetName, address),
                cellCount,
                hasFormula,
                numberFormat,
                new SelectionSafetyState(
                    areaCount,
                    hasMergedCells,
                    worksheetProtected,
                    workbookReadOnly,
                    hasLegacyArray,
                    hasDynamicArraySpill,
                    spillCheckSupported),
                collaboration);
        }
        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException exception)
        {
            throw new CommandRefusedException(
                RefusalCodes.SelectionUnsupported,
                $"The current Excel selection is not a supported cell range: {exception.Message}",
                "Select one rectangular cell range and try again.");
        }
        finally
        {
            ComRelease.Owned(areasObject);
            ComRelease.Owned(selectionObject);
            ComRelease.Owned(worksheetObject);
            ComRelease.Owned(workbookObject);
            // ExcelDnaUtil owns the root Application RCW; never release a borrowed root.
        }
    }

    public void SetNumberFormat(string formatCode)
    {
        _verifyExcelThread();
        if (string.IsNullOrWhiteSpace(formatCode))
        {
            throw new ArgumentException("A number format is required.", nameof(formatCode));
        }

        ExcelComRetry.Execute(() => SetNumberFormatOnce(formatCode));
    }

    private void SetNumberFormatOnce(string formatCode)
    {
        object? applicationObject = null;
        object? selectionObject = null;

        try
        {
            applicationObject = _getApplication();
            ExcelCommandReadiness.RequireReady(applicationObject);
            dynamic application = applicationObject;
            selectionObject = application.Selection;
            if (selectionObject is null)
            {
                throw new CommandRefusedException(
                    RefusalCodes.SelectionUnsupported,
                    "A cell range selection is required.",
                    "Select one rectangular cell range and try again.");
            }

            ApplicationStateGuard.Run(
                new ExcelApplicationStateAdapter(applicationObject),
                ApplicationStateChangeSet.PropertyMutation(),
                () =>
                {
                    dynamic selection = selectionObject;
                    selection.NumberFormat = formatCode;
                });
        }
        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException exception)
        {
            throw new CommandRefusedException(
                RefusalCodes.SelectionUnsupported,
                $"The current Excel selection cannot be formatted: {exception.Message}",
                "Select one rectangular cell range and try again.");
        }
        finally
        {
            ComRelease.Owned(selectionObject);
            // ExcelDnaUtil owns the root Application RCW; never release a borrowed root.
        }
    }

    private static bool? ConvertFormulaState(object value)
    {
        if (value is bool boolean)
        {
            return boolean;
        }

        return null;
    }

    private static bool ConvertUnsafeBoolean(object value)
    {
        return !(value is bool boolean) || boolean;
    }

    private static bool TryReadOptionalUnsafeBoolean(object target, string propertyName, out bool value)
    {
        try
        {
            var propertyValue = target.GetType().InvokeMember(
                propertyName,
                BindingFlags.GetProperty,
                binder: null,
                target,
                args: null,
                CultureInfo.InvariantCulture);
            value = ConvertUnsafeBoolean(propertyValue);
            return true;
        }
        catch (MissingMethodException)
        {
            value = false;
            return false;
        }
        catch (COMException)
        {
            value = false;
            return false;
        }
    }
}
