using System;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Formatting;
using ExcelAccel.Application.Undo;
using ExcelAccel.Application.Styles;
using ExcelAccel.Application.Formulas;
using ExcelAccel.Application.SelectionTools;
using ExcelAccel.Core.Commands;
using ExcelAccel.Core.Reliability;

namespace ExcelAccel.ExcelInterop;

public sealed class ExcelSelectionAdapter : IFormattingPort, IPropertyReceiptPort, IStylePort, IFormulaBlockPort, ISelectionMatchPort
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

    public string ReadFormattingProperty(string propertyId)
    {
        _verifyExcelThread();
        if (string.IsNullOrWhiteSpace(propertyId))
        {
            throw new ArgumentException("A formatting property ID is required.", nameof(propertyId));
        }

        return ExcelComRetry.Execute(() => ReadFormattingPropertyOnce(propertyId));
    }

    public void WriteFormattingProperty(string propertyId, string invariantValue)
    {
        _verifyExcelThread();
        if (string.IsNullOrWhiteSpace(propertyId))
        {
            throw new ArgumentException("A formatting property ID is required.", nameof(propertyId));
        }

        ExcelComRetry.Execute(() => WriteFormattingPropertyOnce(propertyId, invariantValue ?? string.Empty));
    }

    public bool TryRead(SelectionContext target, string propertyId, out string value)
    {
        value = string.Empty;
        try
        {
            if (propertyId == FormulaBlockCommand.ReceiptPropertyId)
                value = CaptureFormulaBlock(target).Contents.Serialize();
            else
            {
                if (!target.Equals(CaptureSelection().Context)) return false;
                value = ReadFormattingProperty(propertyId);
            }
            return true;
        }
        catch (CommandRefusedException) { return false; }
        catch (System.Runtime.InteropServices.COMException) { return false; }
    }

    public bool TryWrite(SelectionContext target, string propertyId, string value)
    {
        try
        {
            if (propertyId == FormulaBlockCommand.ReceiptPropertyId)
                WriteFormulaBlock(target, FormulaCellBlock.Deserialize(value));
            else
            {
                if (!target.Equals(CaptureSelection().Context)) return false;
                WriteFormattingProperty(propertyId, value);
            }
            return true;
        }
        catch (CommandRefusedException) { return false; }
        catch (System.Runtime.InteropServices.COMException) { return false; }
    }

    public FormulaBlockSnapshot CaptureFormulaBlock()
    {
        _verifyExcelThread();
        var selection = CaptureSelection();
        return ExcelComRetry.Execute(() => CaptureFormulaBlockOnce(selection));
    }

    public FormulaBlockSnapshot CaptureFormulaBlock(SelectionContext target)
    {
        _verifyExcelThread();
        if (target is null) throw new ArgumentNullException(nameof(target));
        return ExcelComRetry.Execute(() => CaptureTargetFormulaBlockOnce(target));
    }

    public void WriteFormulaBlock(FormulaCellBlock contents)
    {
        _verifyExcelThread();
        if (contents is null) throw new ArgumentNullException(nameof(contents));
        ExcelComRetry.Execute(() => WriteFormulaBlockOnce(contents));
    }

    public void WriteFormulaBlock(SelectionContext target, FormulaCellBlock contents)
    {
        _verifyExcelThread();
        if (target is null) throw new ArgumentNullException(nameof(target));
        if (contents is null) throw new ArgumentNullException(nameof(contents));
        ExcelComRetry.Execute(() => WriteTargetFormulaBlockOnce(target, contents));
    }

    private FormulaBlockSnapshot CaptureTargetFormulaBlockOnce(SelectionContext target)
    {
        object? applicationObject = null;
        object? workbookObject = null;
        object? worksheetsObject = null;
        object? worksheetObject = null;
        object? rangeObject = null;
        object? areasObject = null;
        object? rowsObject = null;
        object? columnsObject = null;
        try
        {
            applicationObject = _getApplication();
            ExcelCommandReadiness.RequireReady(applicationObject);
            dynamic application = applicationObject;
            workbookObject = application.ActiveWorkbook;
            if (workbookObject is null) throw new CommandRefusedException(RefusalCodes.SelectionUnsupported, "An active workbook is required.", "Open the planned workbook and retry.");
            RequireWorkbookIdentity(workbookObject, target.WorkbookId);
            worksheetsObject = ((dynamic)workbookObject).Worksheets;
            worksheetObject = ((dynamic)worksheetsObject)[target.WorksheetName];
            dynamic worksheet = worksheetObject;
            rangeObject = worksheet.Range[target.Address];
            dynamic range = rangeObject;
            var address = Convert.ToString(range.Address[false, false, 1, false], CultureInfo.InvariantCulture) ?? string.Empty;
            var context = new SelectionContext(target.WorkbookId, target.WorksheetName, address);
            if (!target.Equals(context)) throw new CommandRefusedException(RefusalCodes.StaleContext, "The planned range no longer resolves to the same address.", "Refresh the plan and retry.");
            rowsObject = range.Rows;
            columnsObject = range.Columns;
            var rowCount = Convert.ToInt32(((dynamic)rowsObject).Count, CultureInfo.InvariantCulture);
            var columnCount = Convert.ToInt32(((dynamic)columnsObject).Count, CultureInfo.InvariantCulture);
            var firstRow = Convert.ToInt32(range.Row, CultureInfo.InvariantCulture);
            var firstColumn = Convert.ToInt32(range.Column, CultureInfo.InvariantCulture);
            var cellCount = checked((long)rowCount * columnCount);
            if (cellCount > FormulaCellBlock.MaximumCells) throw new CommandRefusedException(RefusalCodes.ResourceLimit, "The range exceeds the bounded formula cell limit.", "Use a smaller range.");
            areasObject = range.Areas;
            var spillCheckSupported = TryReadOptionalUnsafeBoolean(rangeObject, "HasSpill", out var hasSpill);
            var safety = new SelectionSafetyState(
                Convert.ToInt32(((dynamic)areasObject).Count, CultureInfo.InvariantCulture),
                ConvertUnsafeBoolean(range.MergeCells),
                Convert.ToBoolean(worksheet.ProtectContents, CultureInfo.InvariantCulture),
                Convert.ToBoolean(((dynamic)workbookObject).ReadOnly, CultureInfo.InvariantCulture),
                ConvertUnsafeBoolean(range.HasArray),
                spillCheckSupported && hasSpill,
                spillCheckSupported);
            var selection = new SelectionSnapshot(context, cellCount, ConvertFormulaState(range.HasFormula),
                Convert.ToString(range.NumberFormat, CultureInfo.InvariantCulture) ?? "(mixed)", safety,
                ExcelWorkbookCollaborationAdapter.Capture(workbookObject));
            object formulaValues = range.Formula;
            object valueValues = range.Value2;
            var cells = new FormulaCellValue[rowCount * columnCount];
            for (var row = 0; row < rowCount; row++)
                for (var column = 0; column < columnCount; column++)
                    cells[(row * columnCount) + column] = ConvertFormulaCell(
                        ArrayValue(formulaValues, row, column), ArrayValue(valueValues, row, column));
            return new FormulaBlockSnapshot(selection, firstRow, firstColumn, new FormulaCellBlock(rowCount, columnCount, cells));
        }
        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException exception)
        {
            throw new CommandRefusedException(RefusalCodes.SelectionUnsupported, $"The planned formula range is unavailable: {exception.Message}", "Open the planned workbook and worksheet, then refresh the plan.");
        }
        finally
        {
            ComRelease.Owned(columnsObject);
            ComRelease.Owned(rowsObject);
            ComRelease.Owned(areasObject);
            ComRelease.Owned(rangeObject);
            ComRelease.Owned(worksheetObject);
            ComRelease.Owned(worksheetsObject);
            ComRelease.Owned(workbookObject);
        }
    }

    private void WriteTargetFormulaBlockOnce(SelectionContext target, FormulaCellBlock contents)
    {
        object? applicationObject = null;
        object? workbookObject = null;
        object? worksheetsObject = null;
        object? worksheetObject = null;
        object? rangeObject = null;
        object? rowsObject = null;
        object? columnsObject = null;
        try
        {
            applicationObject = _getApplication();
            ExcelCommandReadiness.RequireReady(applicationObject);
            dynamic application = applicationObject;
            workbookObject = application.ActiveWorkbook;
            if (workbookObject is null) throw new CommandRefusedException(RefusalCodes.SelectionUnsupported, "An active workbook is required.", "Open the planned workbook and retry.");
            RequireWorkbookIdentity(workbookObject, target.WorkbookId);
            worksheetsObject = ((dynamic)workbookObject).Worksheets;
            worksheetObject = ((dynamic)worksheetsObject)[target.WorksheetName];
            rangeObject = ((dynamic)worksheetObject).Range[target.Address];
            dynamic range = rangeObject;
            rowsObject = range.Rows;
            columnsObject = range.Columns;
            var rows = Convert.ToInt32(((dynamic)rowsObject).Count, CultureInfo.InvariantCulture);
            var columns = Convert.ToInt32(((dynamic)columnsObject).Count, CultureInfo.InvariantCulture);
            if (rows != contents.RowCount || columns != contents.ColumnCount)
                throw new CommandRefusedException(RefusalCodes.StaleContext, "The planned formula range shape changed.", "Refresh the plan and retry.");
            object? writeValue;
            if (contents.CellCount == 1) writeValue = ToExcelFormulaValue(contents[0, 0]);
            else
            {
                var values = new object?[rows, columns];
                for (var row = 0; row < rows; row++)
                    for (var column = 0; column < columns; column++) values[row, column] = ToExcelFormulaValue(contents[row, column]);
                writeValue = values;
            }
            ApplicationStateGuard.Run(new ExcelApplicationStateAdapter(applicationObject),
                ApplicationStateChangeSet.PropertyMutation(), () =>
                {
                    range.ClearContents();
                    range.Formula = writeValue;
                });
        }
        finally
        {
            ComRelease.Owned(columnsObject);
            ComRelease.Owned(rowsObject);
            ComRelease.Owned(rangeObject);
            ComRelease.Owned(worksheetObject);
            ComRelease.Owned(worksheetsObject);
            ComRelease.Owned(workbookObject);
        }
    }

    public void SelectAreas(SelectionContext sourceContext, System.Collections.Generic.IReadOnlyList<SelectionArea> areas)
    {
        _verifyExcelThread();
        if (sourceContext is null) throw new ArgumentNullException(nameof(sourceContext));
        if (areas is null || areas.Count == 0) throw new ArgumentException("At least one selection area is required.", nameof(areas));
        ExcelComRetry.Execute(() => SelectAreasOnce(sourceContext, areas));
    }

    public System.Collections.Generic.IReadOnlyList<string> CaptureSelectedAreaAddresses()
    {
        _verifyExcelThread();
        return ExcelComRetry.Execute(CaptureSelectedAreaAddressesOnce);
    }

    private void SelectAreasOnce(SelectionContext sourceContext, System.Collections.Generic.IReadOnlyList<SelectionArea> areas)
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
            dynamic application = applicationObject;
            workbookObject = application.ActiveWorkbook;
            if (workbookObject is null) throw new CommandRefusedException(RefusalCodes.SelectionUnsupported, "An active workbook is required.", "Open the planned workbook and retry.");
            RequireWorkbookIdentity(workbookObject, sourceContext.WorkbookId);
            worksheetsObject = ((dynamic)workbookObject).Worksheets;
            worksheetObject = ((dynamic)worksheetsObject)[sourceContext.WorksheetName];
            var address = string.Join(",", System.Linq.Enumerable.Select(areas, area => area.Address));
            rangeObject = ((dynamic)worksheetObject).Range[address];
            ((dynamic)rangeObject).Select();
        }
        finally
        {
            ComRelease.Owned(rangeObject);
            ComRelease.Owned(worksheetObject);
            ComRelease.Owned(worksheetsObject);
            ComRelease.Owned(workbookObject);
        }
    }

    private System.Collections.Generic.IReadOnlyList<string> CaptureSelectedAreaAddressesOnce()
    {
        object? applicationObject = null;
        object? selectionObject = null;
        object? areasObject = null;
        var captured = new System.Collections.Generic.List<System.Tuple<int, int, string>>();
        try
        {
            applicationObject = _getApplication();
            ExcelCommandReadiness.RequireReady(applicationObject);
            selectionObject = ((dynamic)applicationObject).Selection;
            if (selectionObject is null) throw new CommandRefusedException(RefusalCodes.SelectionUnsupported, "A cell range selection is required.", "Select a cell range and retry.");
            areasObject = ((dynamic)selectionObject).Areas;
            var count = Convert.ToInt32(((dynamic)areasObject).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                object? areaObject = null;
                try
                {
                    areaObject = ((dynamic)areasObject)[index];
                    dynamic area = areaObject;
                    captured.Add(System.Tuple.Create(
                        Convert.ToInt32(area.Row, CultureInfo.InvariantCulture),
                        Convert.ToInt32(area.Column, CultureInfo.InvariantCulture),
                        Convert.ToString(area.Address[false, false, 1, false], CultureInfo.InvariantCulture) ?? string.Empty));
                }
                finally { ComRelease.Owned(areaObject); }
            }
            captured.Sort((left, right) =>
            {
                var row = left.Item1.CompareTo(right.Item1);
                return row != 0 ? row : left.Item2.CompareTo(right.Item2);
            });
            return captured.ConvertAll(value => value.Item3).AsReadOnly();
        }
        finally
        {
            ComRelease.Owned(areasObject);
            ComRelease.Owned(selectionObject);
        }
    }

    private static void RequireWorkbookIdentity(object workbookObject, string expectedWorkbookId)
    {
        dynamic workbook = workbookObject;
        var actual = Convert.ToString(workbook.FullName, CultureInfo.InvariantCulture) ?? string.Empty;
        if (string.IsNullOrWhiteSpace(actual)) actual = Convert.ToString(workbook.Name, CultureInfo.InvariantCulture) ?? string.Empty;
        if (!string.Equals(actual, expectedWorkbookId, StringComparison.OrdinalIgnoreCase))
            throw new CommandRefusedException(RefusalCodes.StaleContext, "The active workbook is not the planned workbook.", "Activate the planned workbook and refresh the command.");
    }

    private FormulaBlockSnapshot CaptureFormulaBlockOnce(SelectionSnapshot snapshot)
    {
        object? applicationObject = null;
        object? selectionObject = null;
        object? rowsObject = null;
        object? columnsObject = null;
        try
        {
            applicationObject = _getApplication();
            ExcelCommandReadiness.RequireReady(applicationObject);
            dynamic application = applicationObject;
            selectionObject = application.Selection;
            if (selectionObject is null) throw new CommandRefusedException(RefusalCodes.SelectionUnsupported, "A cell range selection is required.", "Select one rectangular range.");
            dynamic selection = selectionObject;
            var address = Convert.ToString(selection.Address[false, false, 1, false], CultureInfo.InvariantCulture) ?? string.Empty;
            if (!string.Equals(address, snapshot.Context.Address, StringComparison.OrdinalIgnoreCase))
                throw new CommandRefusedException(RefusalCodes.StaleContext, "The selection changed during formula capture.", "Retry the command without changing the selection.");
            rowsObject = selection.Rows;
            columnsObject = selection.Columns;
            var rowCount = Convert.ToInt32(((dynamic)rowsObject).Count, CultureInfo.InvariantCulture);
            var columnCount = Convert.ToInt32(((dynamic)columnsObject).Count, CultureInfo.InvariantCulture);
            var firstRow = Convert.ToInt32(selection.Row, CultureInfo.InvariantCulture);
            var firstColumn = Convert.ToInt32(selection.Column, CultureInfo.InvariantCulture);
            if ((long)rowCount * columnCount > FormulaCellBlock.MaximumCells)
                throw new CommandRefusedException(RefusalCodes.ResourceLimit, "The selection exceeds the bounded formula cell limit.", "Select a smaller range.");
            object formulaValues = selection.Formula;
            object valueValues = selection.Value2;
            var cells = new FormulaCellValue[rowCount * columnCount];
            for (var row = 0; row < rowCount; row++)
                for (var column = 0; column < columnCount; column++)
                    cells[(row * columnCount) + column] = ConvertFormulaCell(
                        ArrayValue(formulaValues, row, column), ArrayValue(valueValues, row, column));
            return new FormulaBlockSnapshot(snapshot, firstRow, firstColumn, new FormulaCellBlock(rowCount, columnCount, cells));
        }
        finally
        {
            ComRelease.Owned(columnsObject);
            ComRelease.Owned(rowsObject);
            ComRelease.Owned(selectionObject);
        }
    }

    private void WriteFormulaBlockOnce(FormulaCellBlock contents)
    {
        object? applicationObject = null;
        object? selectionObject = null;
        object? rowsObject = null;
        object? columnsObject = null;
        try
        {
            applicationObject = _getApplication();
            ExcelCommandReadiness.RequireReady(applicationObject);
            dynamic application = applicationObject;
            selectionObject = application.Selection;
            if (selectionObject is null) throw new CommandRefusedException(RefusalCodes.SelectionUnsupported, "A cell range selection is required.", "Select one rectangular range.");
            dynamic selection = selectionObject;
            rowsObject = selection.Rows;
            columnsObject = selection.Columns;
            var rowCount = Convert.ToInt32(((dynamic)rowsObject).Count, CultureInfo.InvariantCulture);
            var columnCount = Convert.ToInt32(((dynamic)columnsObject).Count, CultureInfo.InvariantCulture);
            if (rowCount != contents.RowCount || columnCount != contents.ColumnCount)
                throw new CommandRefusedException(RefusalCodes.StaleContext, "The selection shape changed before the formula write.", "Refresh the preview and retry.");
            object? writeValue;
            if (contents.CellCount == 1) writeValue = ToExcelFormulaValue(contents[0, 0]);
            else
            {
                var values = new object?[rowCount, columnCount];
                for (var row = 0; row < rowCount; row++)
                    for (var column = 0; column < columnCount; column++)
                        values[row, column] = ToExcelFormulaValue(contents[row, column]);
                writeValue = values;
            }
            ApplicationStateGuard.Run(
                new ExcelApplicationStateAdapter(applicationObject),
                ApplicationStateChangeSet.PropertyMutation(),
                () =>
                {
                    selection.ClearContents();
                    selection.Formula = writeValue;
                });
        }
        finally
        {
            ComRelease.Owned(columnsObject);
            ComRelease.Owned(rowsObject);
            ComRelease.Owned(selectionObject);
        }
    }

    private static object? ArrayValue(object value, int row, int column)
    {
        if (!(value is Array array)) return row == 0 && column == 0 ? value : null;
        if (array.Rank != 2) throw new CommandRefusedException(RefusalCodes.ExcelCapabilityMissing, "Excel returned an unsupported formula matrix shape.", "Retry with a smaller rectangular selection.");
        return array.GetValue(row + array.GetLowerBound(0), column + array.GetLowerBound(1));
    }

    private static FormulaCellValue ConvertFormulaCell(object? formulaValue, object? calculatedValue)
    {
        if (formulaValue is null || formulaValue is DBNull)
        {
            if (calculatedValue is null || calculatedValue is DBNull) return FormulaCellValue.Blank();
            formulaValue = calculatedValue;
        }
        if (formulaValue is ErrorWrapper || calculatedValue is ErrorWrapper)
        {
            if (formulaValue is string errorFormula && errorFormula.StartsWith("=", StringComparison.Ordinal))
                return FormulaCellValue.Formula(errorFormula);
            throw new CommandRefusedException(RefusalCodes.SelectionUnsupported, "Constant Excel error values are outside the formula receipt contract.", "Remove constant error cells from the selection.");
        }
        if (formulaValue is string text)
        {
            if (text.Length == 0 && (calculatedValue is null || calculatedValue is DBNull ||
                (calculatedValue is string emptyCalculated && emptyCalculated.Length == 0)))
                return FormulaCellValue.Blank();
            if (text.StartsWith("=", StringComparison.Ordinal) &&
                !(calculatedValue is string calculatedText && string.Equals(text, calculatedText, StringComparison.Ordinal)))
                return FormulaCellValue.Formula(text);
            if (calculatedValue is bool calculatedBoolean) return FormulaCellValue.Boolean(calculatedBoolean);
            if (calculatedValue is not null && IsNumeric(calculatedValue))
                return FormulaCellValue.Number(Convert.ToDouble(calculatedValue, CultureInfo.InvariantCulture));
            return FormulaCellValue.Text(calculatedValue as string ?? text.TrimStart('\''));
        }
        if (formulaValue is bool boolean) return FormulaCellValue.Boolean(boolean);
        if (IsNumeric(formulaValue)) return FormulaCellValue.Number(Convert.ToDouble(formulaValue, CultureInfo.InvariantCulture));
        throw new CommandRefusedException(RefusalCodes.ExcelCapabilityMissing,
            $"Excel returned unsupported cell value type '{formulaValue.GetType().FullName}'.", "Remove unsupported cell types from the selection.");
    }

    private static object? ToExcelFormulaValue(FormulaCellValue value)
    {
        switch (value.Kind)
        {
            case FormulaCellKind.Blank: return null;
            case FormulaCellKind.Formula: return value.InvariantValue;
            case FormulaCellKind.Text: return value.InvariantValue.StartsWith("=", StringComparison.Ordinal) ? "'" + value.InvariantValue : value.InvariantValue;
            case FormulaCellKind.Number: return value.AsNumber();
            case FormulaCellKind.Boolean: return value.InvariantValue == "true";
            default: throw new InvalidOperationException("Unsupported formula cell kind.");
        }
    }

    private static bool IsNumeric(object value)
    {
        switch (Type.GetTypeCode(value.GetType()))
        {
            case TypeCode.Byte:
            case TypeCode.SByte:
            case TypeCode.UInt16:
            case TypeCode.UInt32:
            case TypeCode.UInt64:
            case TypeCode.Int16:
            case TypeCode.Int32:
            case TypeCode.Int64:
            case TypeCode.Decimal:
            case TypeCode.Double:
            case TypeCode.Single: return true;
            default: return false;
        }
    }

    private string ReadFormattingPropertyOnce(string propertyId)
    {
        object? applicationObject = null;
        object? selectionObject = null;
        object? childObject = null;
        object? ownedObject2 = null;
        try
        {
            applicationObject = _getApplication();
            ExcelCommandReadiness.RequireReady(applicationObject);
            dynamic application = applicationObject;
            selectionObject = application.Selection;
            if (selectionObject is null)
            {
                throw new CommandRefusedException("A cell range selection is required.");
            }

            dynamic selection = selectionObject;
            switch (propertyId)
            {
                case "number_format":
                    return InvariantText(selection.NumberFormat);
                case "font_color":
                    childObject = selection.Font;
                    return OleColorToHex(((dynamic)childObject).Color);
                case "font_name":
                    childObject = selection.Font;
                    return InvariantText(((dynamic)childObject).Name);
                case "font_bold":
                    childObject = selection.Font;
                    return InvariantBoolean(((dynamic)childObject).Bold);
                case "font_italic":
                    childObject = selection.Font;
                    return InvariantBoolean(((dynamic)childObject).Italic);
                case "fill_color":
                    childObject = selection.Interior;
                    return OleColorToHex(((dynamic)childObject).Color);
                case "font_size":
                    childObject = selection.Font;
                    return Convert.ToDouble(((dynamic)childObject).Size, CultureInfo.InvariantCulture).ToString("0.####", CultureInfo.InvariantCulture);
                case "horizontal_alignment":
                    return HorizontalAlignmentToken(selection.HorizontalAlignment);
                case "vertical_alignment":
                    return VerticalAlignmentToken(selection.VerticalAlignment);
                case "indent_level":
                    return selection.IndentLevel is null ? "(mixed)" : Convert.ToInt32(selection.IndentLevel, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture);
                case "underline":
                    childObject = selection.Font;
                    return UnderlineToken(((dynamic)childObject).Underline);
                case "row_height":
                    return InvariantDimension(selection.RowHeight);
                case "column_width":
                    return InvariantDimension(selection.ColumnWidth);
                case "borders":
                    ownedObject2 = selection.Borders;
                    childObject = ((dynamic)ownedObject2)[9];
                    return Convert.ToInt32(((dynamic)childObject).LineStyle, CultureInfo.InvariantCulture) == -4142 ? "none" : "sum_bar";
                case "gridlines":
                    childObject = application.ActiveWindow;
                    return Convert.ToBoolean(((dynamic)childObject).DisplayGridlines, CultureInfo.InvariantCulture).ToString().ToLowerInvariant();
                case "zoom":
                    childObject = application.ActiveWindow;
                    return Convert.ToInt32(((dynamic)childObject).Zoom, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture);
                case "freeze_panes":
                    childObject = application.ActiveWindow;
                    return Convert.ToBoolean(((dynamic)childObject).FreezePanes, CultureInfo.InvariantCulture).ToString().ToLowerInvariant();
                default:
                    throw new CommandRefusedException(
                        RefusalCodes.CommandUnavailable,
                        $"Formatting property '{propertyId}' is not implemented by the Excel adapter.",
                        "Use a currently supported formatting command.");
            }
        }
        finally
        {
            ComRelease.Owned(childObject);
            ComRelease.Owned(ownedObject2);
            ComRelease.Owned(selectionObject);
        }
    }

    private void WriteFormattingPropertyOnce(string propertyId, string invariantValue)
    {
        object? applicationObject = null;
        object? selectionObject = null;
        object? childObject = null;
        object? ownedObject2 = null;
        try
        {
            applicationObject = _getApplication();
            ExcelCommandReadiness.RequireReady(applicationObject);
            dynamic application = applicationObject;
            selectionObject = application.Selection;
            if (selectionObject is null)
            {
                throw new CommandRefusedException("A cell range selection is required.");
            }

            ApplicationStateGuard.Run(
                new ExcelApplicationStateAdapter(applicationObject),
                ApplicationStateChangeSet.PropertyMutation(),
                () =>
                {
                    dynamic selection = selectionObject;
                    switch (propertyId)
                    {
                        case "number_format":
                            selection.NumberFormat = invariantValue;
                            break;
                        case "font_color":
                            childObject = selection.Font;
                            ((dynamic)childObject).Color = HexToOleColor(invariantValue);
                            break;
                        case "font_name":
                            childObject = selection.Font;
                            ((dynamic)childObject).Name = invariantValue;
                            break;
                        case "font_bold":
                            childObject = selection.Font;
                            ((dynamic)childObject).Bold = bool.Parse(invariantValue);
                            break;
                        case "font_italic":
                            childObject = selection.Font;
                            ((dynamic)childObject).Italic = bool.Parse(invariantValue);
                            break;
                        case "fill_color":
                            childObject = selection.Interior;
                            ((dynamic)childObject).Color = HexToOleColor(invariantValue);
                            break;
                        case "font_size":
                            childObject = selection.Font;
                            ((dynamic)childObject).Size = double.Parse(invariantValue, CultureInfo.InvariantCulture);
                            break;
                        case "horizontal_alignment":
                            selection.HorizontalAlignment = HorizontalAlignmentValue(invariantValue);
                            break;
                        case "vertical_alignment":
                            selection.VerticalAlignment = VerticalAlignmentValue(invariantValue);
                            break;
                        case "indent_level":
                            selection.IndentLevel = int.Parse(invariantValue, CultureInfo.InvariantCulture);
                            break;
                        case "underline":
                            childObject = selection.Font;
                            ((dynamic)childObject).Underline = UnderlineValue(invariantValue);
                            break;
                        case "row_height":
                            childObject = selection.EntireRow;
                            if (invariantValue == "autofit") ((dynamic)childObject).AutoFit();
                            else ((dynamic)childObject).RowHeight = double.Parse(invariantValue, CultureInfo.InvariantCulture);
                            break;
                        case "column_width":
                            childObject = selection.EntireColumn;
                            if (invariantValue == "autofit") ((dynamic)childObject).AutoFit();
                            else ((dynamic)childObject).ColumnWidth = double.Parse(invariantValue, CultureInfo.InvariantCulture);
                            break;
                        case "borders":
                            ownedObject2 = selection.Borders;
                            if (invariantValue == "none") ((dynamic)ownedObject2).LineStyle = -4142;
                            else
                            {
                                childObject = ((dynamic)ownedObject2)[9];
                                ((dynamic)childObject).LineStyle = -4119;
                            }
                            break;
                        case "gridlines":
                            childObject = application.ActiveWindow;
                            ((dynamic)childObject).DisplayGridlines = bool.Parse(invariantValue);
                            break;
                        case "zoom":
                            childObject = application.ActiveWindow;
                            ((dynamic)childObject).Zoom = int.Parse(invariantValue, CultureInfo.InvariantCulture);
                            break;
                        case "freeze_panes":
                            childObject = application.ActiveWindow;
                            if (bool.Parse(invariantValue)) selection.Select();
                            ((dynamic)childObject).FreezePanes = bool.Parse(invariantValue);
                            break;
                        default:
                            throw new CommandRefusedException(
                                RefusalCodes.CommandUnavailable,
                                $"Formatting property '{propertyId}' is not implemented by the Excel adapter.",
                                "Use a currently supported formatting command.");
                    }
                });
        }
        finally
        {
            ComRelease.Owned(childObject);
            ComRelease.Owned(ownedObject2);
            ComRelease.Owned(selectionObject);
        }
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

    private static int HexToOleColor(string value)
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

    private static string OleColorToHex(object value)
    {
        if (value is null) return "(mixed)";
        var color = Convert.ToInt32(value, CultureInfo.InvariantCulture);
        var red = color & 0xFF;
        var green = (color >> 8) & 0xFF;
        var blue = (color >> 16) & 0xFF;
        return $"#{red:X2}{green:X2}{blue:X2}";
    }

    private static string InvariantDimension(object value) =>
        value is null ? "(mixed)" : Convert.ToDouble(value, CultureInfo.InvariantCulture).ToString("0.####", CultureInfo.InvariantCulture);

    private static string InvariantText(object value) =>
        value is null ? "(mixed)" : Convert.ToString(value, CultureInfo.InvariantCulture) ?? "(mixed)";

    private static string InvariantBoolean(object value) =>
        value is bool boolean ? boolean.ToString().ToLowerInvariant() : "(mixed)";

    private static string HorizontalAlignmentToken(object value)
    {
        if (value is null) return "(mixed)";
        switch (Convert.ToInt32(value, CultureInfo.InvariantCulture))
        {
            case -4131: return "left";
            case -4108: return "center";
            case -4152: return "right";
            case 7: return "center_across";
            default: return "general";
        }
    }

    private static int HorizontalAlignmentValue(string value)
    {
        switch (value)
        {
            case "left": return -4131;
            case "center": return -4108;
            case "right": return -4152;
            case "center_across": return 7;
            case "general": return 1;
            default: throw new ArgumentException("Unsupported horizontal alignment.", nameof(value));
        }
    }

    private static string VerticalAlignmentToken(object value)
    {
        if (value is null) return "(mixed)";
        switch (Convert.ToInt32(value, CultureInfo.InvariantCulture))
        {
            case -4160: return "top";
            case -4108: return "center";
            default: return "bottom";
        }
    }

    private static int VerticalAlignmentValue(string value)
    {
        switch (value)
        {
            case "top": return -4160;
            case "center": return -4108;
            case "bottom": return -4107;
            default: throw new ArgumentException("Unsupported vertical alignment.", nameof(value));
        }
    }

    private static string UnderlineToken(object value)
    {
        if (value is null) return "(mixed)";
        switch (Convert.ToInt32(value, CultureInfo.InvariantCulture))
        {
            case 2: return "single";
            case -4119: return "double";
            default: return "none";
        }
    }

    private static int UnderlineValue(string value)
    {
        switch (value)
        {
            case "single": return 2;
            case "double": return -4119;
            case "none": return -4142;
            default: throw new ArgumentException("Unsupported underline style.", nameof(value));
        }
    }
}
