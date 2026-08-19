using System;
using System.Globalization;
using ExcelDna.Integration;
using ExcelAccel.Core.Commands;
using ExcelAccel.ExcelAddIn.Reliability;

namespace ExcelAccel.ExcelAddIn.Interop;

internal sealed class ExcelSelectionAdapter : ISelectionPort
{
    public SelectionSnapshot CaptureSelection()
    {
        RuntimeState.VerifyExcelThread();
        object? applicationObject = null;
        object? workbookObject = null;
        object? worksheetObject = null;
        object? selectionObject = null;

        try
        {
            applicationObject = ExcelDnaUtil.Application;
            dynamic application = applicationObject;
            workbookObject = application.ActiveWorkbook;
            worksheetObject = application.ActiveSheet;
            selectionObject = application.Selection;

            if (workbookObject is null || worksheetObject is null || selectionObject is null)
            {
                throw new CommandRefusedException("An open workbook and a cell range selection are required.");
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

            return new SelectionSnapshot(
                new SelectionContext(workbookId, worksheetName, address),
                cellCount,
                hasFormula,
                numberFormat);
        }
        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException exception)
        {
            throw new CommandRefusedException($"The current Excel selection is not a supported cell range: {exception.Message}");
        }
        finally
        {
            ComRelease.Owned(selectionObject);
            ComRelease.Owned(worksheetObject);
            ComRelease.Owned(workbookObject);
            // ExcelDnaUtil owns the root Application RCW; never release a borrowed root.
        }
    }

    public void SetNumberFormat(string formatCode)
    {
        RuntimeState.VerifyExcelThread();
        if (string.IsNullOrWhiteSpace(formatCode))
        {
            throw new ArgumentException("A number format is required.", nameof(formatCode));
        }

        object? applicationObject = null;
        object? selectionObject = null;

        try
        {
            applicationObject = ExcelDnaUtil.Application;
            dynamic application = applicationObject;
            selectionObject = application.Selection;
            if (selectionObject is null)
            {
                throw new CommandRefusedException("A cell range selection is required.");
            }

            using (var state = ExcelStateGuard.SuppressScreenUpdating(applicationObject))
            {
                dynamic selection = selectionObject;
                selection.NumberFormat = formatCode;
            }
        }
        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException exception)
        {
            throw new CommandRefusedException($"The current Excel selection cannot be formatted: {exception.Message}");
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
}
