using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Operations;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.Formulas;

namespace ExcelAccel.Application.Auditing;

/// <summary>
/// The Excel-side surface a dependent scan needs. Every method is read-only. The
/// port reports what the workbook says; all bounding, banding, progress, and
/// cancellation policy lives in the coordinator so it stays testable without
/// Excel.
/// </summary>
public interface IDependentScanPort
{
    /// <summary>The single cell or rectangular range the user selected.</summary>
    AuditCellIdentity CaptureTarget();

    /// <summary>Every worksheet in the workbook, for a workbook-scope plan.</summary>
    IReadOnlyList<string> CaptureWorksheetNames();

    /// <summary>The region a worksheet reports as used. Untrusted.</summary>
    UsedRegionBounds CaptureUsedRegion(string worksheetName);

    /// <summary>Reads the formula cells of exactly one banded rectangle.</summary>
    IReadOnlyList<AuditFormulaCell> CaptureBlock(string worksheetName, AuditRectangle band);

    /// <summary>Resolvable defined names in scope. Unresolvable ones are omitted.</summary>
    IReadOnlyList<AuditNameBinding> CaptureNames(DependentScanScope scope);
}

/// <summary>
/// What the user confirms before a scan reads anything. It describes the planned
/// read, including the per-worksheet inventory for a workbook scan; nothing has
/// been read yet.
/// </summary>
public sealed class DependentScanPreview
{
    public DependentScanPreview(
        string scopeLabel,
        string targetDisplay,
        long cellCount,
        int blockCount,
        IReadOnlyList<string>? inventoryLines = null)
    {
        ScopeLabel = scopeLabel;
        TargetDisplay = targetDisplay;
        CellCount = cellCount;
        BlockCount = blockCount;
        InventoryLines = inventoryLines ?? Array.Empty<string>();
    }

    /// <summary>"worksheet" or "workbook".</summary>
    public string ScopeLabel { get; }

    public string TargetDisplay { get; }

    public long CellCount { get; }

    public int BlockCount { get; }

    /// <summary>The per-worksheet inventory a workbook scan must show.</summary>
    public IReadOnlyList<string> InventoryLines { get; }
}

public sealed class DirectDependentCoordinator
{
    /// <summary>
    /// Planned region size above which a worksheet scan must be confirmed before
    /// any block is read. A workbook scan always confirms, whatever its size.
    /// </summary>
    public const long PreviewThresholdCells = 25_000;

    public DirectDependentResult Execute(
        IDependentScanPort port,
        OperationProgressTracker? tracker = null,
        Func<DependentScanPreview, bool>? confirmScan = null,
        DependentScanScopeKind scopeKind = DependentScanScopeKind.Worksheet)
    {
        if (port is null) throw new ArgumentNullException(nameof(port));
        var progress = tracker ?? new OperationProgressTracker();
        var target = port.CaptureTarget();
        var scope = scopeKind == DependentScanScopeKind.Workbook
            ? DependentScanScope.Workbook(target.WorkbookId)
            : DependentScanScope.Worksheet(target.WorkbookId, target.WorksheetName);

        var worksheets = scopeKind == DependentScanScopeKind.Workbook
            ? port.CaptureWorksheetNames()
            : new[] { target.WorksheetName };

        if (!WorkbookScanPlan.TryCreate(
                target.WorkbookId,
                worksheets.Select(port.CaptureUsedRegion),
                out var plan,
                out var refusalCode,
                out var message))
        {
            return DirectDependentResult.Refused(target, scope, refusalCode!, message!);
        }

        // A workbook scan always confirms and always shows its sheet inventory.
        // A worksheet scan confirms only above the threshold.
        if (scopeKind == DependentScanScopeKind.Workbook || plan!.TotalCellCount > PreviewThresholdCells)
        {
            var preview = new DependentScanPreview(
                scope.Label,
                AuditPresentationLabels.Location(target),
                plan!.TotalCellCount,
                plan.TotalBlockCount,
                plan.InventoryLines());
            if (confirmScan is null || !confirmScan(preview))
            {
                return DirectDependentResult.Refused(
                    target,
                    scope,
                    AuditRefusalCodes.PreviewRequired,
                    $"Scanning {plan.TotalCellCount:N0} cells across {plan.Included.Count:N0} worksheet(s) was not confirmed, so nothing was read.");
            }
        }

        var names = port.CaptureNames(scope);
        var formulas = new List<AuditFormulaCell>();
        var totalBlocks = Math.Max(plan!.TotalBlockCount, 1);
        var readBlocks = 0;
        progress.Report(new OperationProgress(OperationPhase.Snapshot, 0, totalBlocks, "Reading formulas."));

        foreach (var entry in plan.Included)
        {
            for (var index = 0; index < entry.BlockCount; index++)
            {
                if (progress.CancellationRequested) return Cancelled(progress, target, scope, readBlocks, totalBlocks);
                formulas.AddRange(port.CaptureBlock(entry.WorksheetName, entry.Region!.Block(index)));
                readBlocks++;
                progress.Report(new OperationProgress(OperationPhase.Snapshot, readBlocks, totalBlocks, "Reading formulas."));
            }
        }

        if (progress.CancellationRequested) return Cancelled(progress, target, scope, readBlocks, totalBlocks);

        progress.Report(new OperationProgress(OperationPhase.Analyze, 0, 1, "Building the bounded reverse index."));
        var result = ReverseReferenceIndex
            .Build(scope, formulas, names, FormulaDialect.InvariantA1, plan.Excluded.Count)
            .FindDirectDependents(target);
        progress.Report(new OperationProgress(OperationPhase.Completed, 1, 1, "Dependent scan complete."));
        return result;
    }

    private static DirectDependentResult Cancelled(
        OperationProgressTracker progress,
        AuditCellIdentity target,
        DependentScanScope scope,
        int readBlocks,
        int totalBlocks)
    {
        progress.Report(new OperationProgress(OperationPhase.Cancelled, readBlocks, totalBlocks, "Dependent scan cancelled."));
        return DirectDependentResult.Refused(
            target,
            scope,
            AuditRefusalCodes.ScanCancelled,
            "The dependent scan was cancelled; no partial scan is reported as a result.");
    }
}
