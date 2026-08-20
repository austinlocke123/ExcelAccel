using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using ExcelAccel.Application.Auditing;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Formulas;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.ExcelInterop;

public sealed class ExcelReferenceSnapshotAdapter : IDirectPrecedentSnapshotPort, IWorkbookPresencePort, IFormulaLookupPort
{
    public const int MaximumCapturedCells = 10_000;
    private readonly ExcelSelectionAdapter _selection;
    private readonly Func<object> _getApplication;
    private readonly Action _verifyExcelThread;

    public ExcelReferenceSnapshotAdapter(Func<object> getApplication, Action verifyExcelThread)
    {
        _getApplication = getApplication ?? throw new ArgumentNullException(nameof(getApplication));
        _verifyExcelThread = verifyExcelThread ?? throw new ArgumentNullException(nameof(verifyExcelThread));
        _selection = new ExcelSelectionAdapter(_getApplication, _verifyExcelThread);
    }

    public FormulaTargetCapture CaptureTarget()
    {
        var block = _selection.CaptureFormulaBlock();
        if (block.Contents.CellCount != 1)
            throw new CommandRefusedException(RefusalCodes.SelectionUnsupported,
                "Direct precedents require exactly one selected cell.", "Select one formula cell and retry.");
        var cell = block.Contents[0, 0];
        if (!cell.IsFormula)
            throw new CommandRefusedException(RefusalCodes.SelectionUnsupported,
                "Direct precedents require one formula cell.", "Select one formula cell and retry.");
        var context = block.Selection.Context;
        return new FormulaTargetCapture(
            new AuditCellIdentity(context.WorkbookId, context.WorksheetName, context.Address),
            cell.InvariantValue);
    }

    public ReferenceSnapshotIndex CaptureIndex(DirectPrecedentCapturePlan plan)
    {
        if (plan is null) throw new ArgumentNullException(nameof(plan));
        var cells = new Dictionary<AuditCellIdentity, AuditCellClassification>();
        var names = new List<AuditNameBinding>();
        var capturedCells = 0;
        foreach (var target in plan.LocalTargets)
        {
            if (capturedCells >= MaximumCapturedCells) break;
            try
            {
                var block = _selection.CaptureFormulaBlock(new SelectionContext(target.WorkbookId, target.WorksheetName, target.Address));
                capturedCells = checked(capturedCells + block.Contents.CellCount);
                if (capturedCells > MaximumCapturedCells) break;
                cells[target] = Classify(block);
            }
            catch (CommandRefusedException exception) when (
                exception.RefusalCode == RefusalCodes.ResourceLimit ||
                exception.RefusalCode == RefusalCodes.SelectionUnsupported ||
                exception.RefusalCode == RefusalCodes.ExcelCapabilityMissing)
            {
                // The analyzer will retain this requested edge with Unknown
                // classification and explicitly mark the result partial.
            }
        }
        foreach (var candidate in plan.NameCandidates)
        {
            if (capturedCells >= MaximumCapturedCells) break;
            var binding = TryResolveName(plan, candidate);
            if (binding is null) continue;
            names.Add(binding);
            if (cells.ContainsKey(binding.Target)) continue;
            try
            {
                var block = _selection.CaptureFormulaBlock(new SelectionContext(
                    binding.Target.WorkbookId, binding.Target.WorksheetName, binding.Target.Address));
                capturedCells = checked(capturedCells + block.Contents.CellCount);
                if (capturedCells > MaximumCapturedCells) break;
                cells[binding.Target] = Classify(block);
            }
            catch (CommandRefusedException) { }
        }
        return new ReferenceSnapshotIndex(cells, names);
    }

    /// <summary>
    /// Reads one cell's formula for a precedent traversal step. A cell that is
    /// not a formula returns null, which ends that branch of the chain; it is a
    /// leaf, not a coverage gap.
    /// </summary>
    public string? TryReadFormula(AuditCellIdentity cell)
    {
        if (cell is null) throw new ArgumentNullException(nameof(cell));
        try
        {
            var block = _selection.CaptureFormulaBlock(new SelectionContext(cell.WorkbookId, cell.WorksheetName, cell.Address));
            if (block.Contents.CellCount != 1) return null;
            var contents = block.Contents[0, 0];
            return contents.IsFormula ? contents.InvariantValue : null;
        }
        catch (CommandRefusedException)
        {
            return null;
        }
    }

    public bool SourceMatches(FormulaTargetCapture capture)
    {
        if (capture is null) throw new ArgumentNullException(nameof(capture));
        try
        {
            var target = capture.Target;
            var block = _selection.CaptureFormulaBlock(new SelectionContext(target.WorkbookId, target.WorksheetName, target.Address));
            return block.Contents.CellCount == 1 && block.Contents[0, 0].IsFormula &&
                string.Equals(block.Contents[0, 0].InvariantValue, capture.Formula, StringComparison.Ordinal);
        }
        catch (CommandRefusedException)
        {
            return false;
        }
    }

    /// <summary>
    /// Reports whether the workbook that produced a captured analysis is still
    /// open. It enumerates only already-open workbooks and never opens one.
    /// A transient COM failure is reported as <see cref="WorkbookPresence.Unknown"/>
    /// so a caller can neither claim the source is live nor discard it silently.
    /// </summary>
    public WorkbookPresence Probe(string workbookId)
    {
        if (string.IsNullOrWhiteSpace(workbookId)) return WorkbookPresence.Closed;
        _verifyExcelThread();
        try
        {
            return ExcelComRetry.Execute(() => ProbeOnce(workbookId));
        }
        catch (COMException)
        {
            return WorkbookPresence.Unknown;
        }
        catch (CommandRefusedException)
        {
            return WorkbookPresence.Unknown;
        }
    }

    private WorkbookPresence ProbeOnce(string workbookId)
    {
        object? workbooksObject = null;
        try
        {
            workbooksObject = ((dynamic)_getApplication()).Workbooks;
            var count = Convert.ToInt32(((dynamic)workbooksObject).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                object? workbookObject = null;
                try
                {
                    workbookObject = ((dynamic)workbooksObject)[index];
                    if (Matches(workbookObject, workbookId)) return WorkbookPresence.Open;
                }
                finally
                {
                    ComRelease.Owned(workbookObject);
                }
            }

            return WorkbookPresence.Closed;
        }
        finally
        {
            ComRelease.Owned(workbooksObject);
        }
    }

    private static bool Matches(object workbookObject, string workbookId)
    {
        var fullName = Convert.ToString(((dynamic)workbookObject).FullName, CultureInfo.InvariantCulture) ?? string.Empty;
        if (string.Equals(fullName, workbookId, StringComparison.OrdinalIgnoreCase)) return true;
        var name = Convert.ToString(((dynamic)workbookObject).Name, CultureInfo.InvariantCulture) ?? string.Empty;
        return string.Equals(name, workbookId, StringComparison.OrdinalIgnoreCase);
    }

    private static AuditCellClassification Classify(FormulaBlockSnapshot block)
    {
        AuditCellClassification? classification = null;
        for (var row = 0; row < block.Contents.RowCount; row++)
        {
            for (var column = 0; column < block.Contents.ColumnCount; column++)
            {
                var current = ClassifyCell(block, row, column);
                if (classification is null) classification = current;
                else if (classification.Value != current) return AuditCellClassification.Mixed;
            }
        }
        return classification ?? AuditCellClassification.Unknown;
    }

    private static AuditCellClassification ClassifyCell(FormulaBlockSnapshot block, int row, int column)
    {
        if (block.UnderlyingValues?[row, column].Kind == UnderlyingValueKind.Error)
            return AuditCellClassification.Error;
        var cell = block.Contents[row, column];
        if (cell.IsFormula) return AuditCellClassification.Formula;
        if (cell.IsBlank) return AuditCellClassification.Blank;
        return AuditCellClassification.Value;
    }

    private AuditNameBinding? TryResolveName(DirectPrecedentCapturePlan plan, string candidate)
    {
        _verifyExcelThread();
        try
        {
            return ExcelComRetry.Execute(() => ResolveNameOnce(plan, candidate));
        }
        catch (COMException) { return null; }
        catch (CommandRefusedException) { return null; }
    }

    private AuditNameBinding? ResolveNameOnce(DirectPrecedentCapturePlan plan, string candidate)
    {
        object? workbookObject = null;
        object? worksheetsObject = null;
        object? worksheetObject = null;
        object? namesObject = null;
        object? nameObject = null;
        try
        {
            var applicationObject = _getApplication();
            ExcelCommandReadiness.RequireReady(applicationObject);
            workbookObject = ((dynamic)applicationObject).ActiveWorkbook;
            if (workbookObject is null) return null;
            var workbookId = Convert.ToString(((dynamic)workbookObject).FullName, CultureInfo.InvariantCulture) ?? string.Empty;
            if (string.IsNullOrWhiteSpace(workbookId)) workbookId = Convert.ToString(((dynamic)workbookObject).Name, CultureInfo.InvariantCulture) ?? string.Empty;
            if (!string.Equals(workbookId, plan.Target.WorkbookId, StringComparison.OrdinalIgnoreCase)) return null;
            worksheetsObject = ((dynamic)workbookObject).Worksheets;
            worksheetObject = ((dynamic)worksheetsObject)[plan.Target.WorksheetName];

            var scope = AuditNameScope.Worksheet;
            try
            {
                namesObject = ((dynamic)worksheetObject).Names;
                nameObject = ((dynamic)namesObject).Item(candidate);
            }
            catch (COMException)
            {
                ComRelease.Owned(nameObject); nameObject = null;
                ComRelease.Owned(namesObject); namesObject = null;
                scope = AuditNameScope.Workbook;
                namesObject = ((dynamic)workbookObject).Names;
                nameObject = ((dynamic)namesObject).Item(candidate);
            }

            var refersTo = Convert.ToString(((dynamic)nameObject).RefersTo, CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(refersTo)) return null;
            var namePlan = new DirectPrecedentAnalyzer().CreateCapturePlan(plan.Target, refersTo!);
            if (namePlan.LocalTargets.Count != 1 || namePlan.NameCandidates.Count != 0) return null;
            return new AuditNameBinding(candidate, scope, namePlan.LocalTargets[0],
                scope == AuditNameScope.Worksheet ? plan.Target.WorksheetName : null);
        }
        finally
        {
            ComRelease.Owned(nameObject);
            ComRelease.Owned(namesObject);
            ComRelease.Owned(worksheetObject);
            ComRelease.Owned(worksheetsObject);
            ComRelease.Owned(workbookObject);
        }
    }
}
