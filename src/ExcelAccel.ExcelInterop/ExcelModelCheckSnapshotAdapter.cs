using System;
using System.Collections.Generic;
using ExcelAccel.Application.Auditing;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Formatting;
using ExcelAccel.Application.Formulas;
using ExcelAccel.Application.ModelCheck;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.Commands;
using ExcelAccel.Core.ModelCheck;

namespace ExcelAccel.ExcelInterop;

/// <summary>
/// Read-only capture of the cells a Model Check scan runs against: formulas,
/// classifications, number formats, and captured error text. It reads through
/// the existing Excel-thread and COM-retry boundaries, writes nothing, selects
/// nothing, and never recalculates.
/// </summary>
public sealed class ExcelModelCheckSnapshotAdapter : IModelCheckSnapshotPort
{
    private readonly ExcelSelectionAdapter _selection;
    private readonly ExcelDependentScanAdapter _scan;

    public ExcelModelCheckSnapshotAdapter(Func<object> getApplication, Action verifyExcelThread)
    {
        if (getApplication is null) throw new ArgumentNullException(nameof(getApplication));
        if (verifyExcelThread is null) throw new ArgumentNullException(nameof(verifyExcelThread));
        _selection = new ExcelSelectionAdapter(getApplication, verifyExcelThread);
        _scan = new ExcelDependentScanAdapter(getApplication, verifyExcelThread);
    }

    public AuditCellIdentity CaptureTarget()
    {
        var snapshot = _selection.CaptureSelection();
        if (snapshot.Safety.AreaCount != 1)
        {
            throw new CommandRefusedException(
                RefusalCodes.MultiAreaUnsupported,
                "Model Check requires one rectangular selection.",
                "Select one cell or one rectangular range and retry.");
        }

        var context = snapshot.Context;
        return new AuditCellIdentity(context.WorkbookId, context.WorksheetName, context.Address);
    }

    public IReadOnlyList<string> CaptureWorksheetNames() => _scan.CaptureWorksheetNames();

    public UsedRegionBounds CaptureUsedRegion(string worksheetName) => _scan.CaptureUsedRegion(worksheetName);

    public IReadOnlyList<ModelCheckCell> CaptureBlock(string worksheetName, AuditRectangle band)
    {
        if (string.IsNullOrWhiteSpace(worksheetName)) throw new ArgumentException("A worksheet name is required.", nameof(worksheetName));
        var target = CaptureTarget();
        var context = new SelectionContext(target.WorkbookId, worksheetName, band.ToString());
        var formulas = _selection.CaptureFormulaBlock(context);
        var formats = TryCaptureFormats(context);

        var cells = new List<ModelCheckCell>();
        for (var row = 0; row < formulas.Contents.RowCount; row++)
        {
            for (var column = 0; column < formulas.Contents.ColumnCount; column++)
            {
                var cell = formulas.Contents[row, column];
                var underlying = formulas.UnderlyingValues?[row, column];
                var isError = underlying?.Kind == UnderlyingValueKind.Error;
                var classification = isError
                    ? AuditCellClassification.Error
                    : cell.IsFormula
                        ? AuditCellClassification.Formula
                        : cell.IsBlank
                            ? AuditCellClassification.Blank
                            : AuditCellClassification.Value;

                cells.Add(new ModelCheckCell(
                    new AuditCellIdentity(
                        target.WorkbookId,
                        worksheetName,
                        AuditAddress.Cell(formulas.FirstRow + row, formulas.FirstColumn + column)),
                    cell.IsFormula ? cell.InvariantValue : null,
                    classification,
                    ReadFormat(formats, row, column),
                    isError ? underlying!.InvariantValue : null));
            }
        }

        return cells;
    }

    private FormatBlockSnapshot? TryCaptureFormats(SelectionContext context)
    {
        try
        {
            return _selection.CaptureFormatBlock(context);
        }
        catch (CommandRefusedException)
        {
            // The bounded formats read has a tighter ceiling than the formula
            // read. Without it every cell reports the neutral format, which the
            // format rule treats as a single consistent baseline rather than a
            // false exception.
            return null;
        }
    }

    private static string ReadFormat(FormatBlockSnapshot? formats, int row, int column)
    {
        if (formats is null) return string.Empty;
        if (row >= formats.Contents.RowCount || column >= formats.Contents.ColumnCount) return string.Empty;
        return formats.Contents[row, column].NumberFormat ?? string.Empty;
    }
}
