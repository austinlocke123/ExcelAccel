using System;
using System.Collections.Generic;
using ExcelAccel.Application.Operations;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.Formulas;

namespace ExcelAccel.Application.Auditing;

/// <summary>
/// The Excel-side surface a worksheet dependent scan needs. Every method is
/// read-only. The port reports what the worksheet says; all bounding, banding,
/// progress, and cancellation policy lives in the coordinator so it stays
/// testable without Excel.
/// </summary>
public interface IDependentScanPort
{
    /// <summary>The single cell or rectangular range the user selected.</summary>
    AuditCellIdentity CaptureTarget();

    /// <summary>The region the worksheet reports as used. Untrusted.</summary>
    UsedRegionBounds CaptureUsedRegion(DependentScanScope scope);

    /// <summary>Reads the formula cells of exactly one banded rectangle.</summary>
    IReadOnlyList<AuditFormulaCell> CaptureBlock(DependentScanScope scope, AuditRectangle band);

    /// <summary>Resolvable defined names in scope. Unresolvable ones are omitted.</summary>
    IReadOnlyList<AuditNameBinding> CaptureNames(DependentScanScope scope);
}

/// <summary>
/// What the user is asked to confirm before a scan large enough to be worth
/// noticing. It describes the planned read; nothing has been read yet.
/// </summary>
public sealed class DependentScanPreview
{
    public DependentScanPreview(string worksheetName, string targetDisplay, long cellCount, int blockCount)
    {
        WorksheetName = worksheetName;
        TargetDisplay = targetDisplay;
        CellCount = cellCount;
        BlockCount = blockCount;
    }

    public string WorksheetName { get; }

    public string TargetDisplay { get; }

    public long CellCount { get; }

    public int BlockCount { get; }
}

public sealed class DirectDependentCoordinator
{
    /// <summary>
    /// Planned region size above which the scan must be confirmed before any
    /// block is read. The AUDITING contract requires a preview before a scan
    /// above the scan threshold.
    /// </summary>
    public const long PreviewThresholdCells = 25_000;

    public DirectDependentResult Execute(
        IDependentScanPort port,
        OperationProgressTracker? tracker = null,
        Func<DependentScanPreview, bool>? confirmScan = null)
    {
        if (port is null) throw new ArgumentNullException(nameof(port));
        var progress = tracker ?? new OperationProgressTracker();
        var target = port.CaptureTarget();
        var scope = DependentScanScope.Worksheet(target.WorkbookId, target.WorksheetName);

        if (!DependentScanRegion.TryCreate(port.CaptureUsedRegion(scope), out var region, out var refusalCode, out var message))
        {
            return DirectDependentResult.Refused(target, scope, refusalCode!, message!);
        }

        if (region!.CellCount > PreviewThresholdCells)
        {
            var preview = new DependentScanPreview(
                region.WorksheetName, AuditPresentationLabels.Location(target), region.CellCount, region.BlockCount);
            // The confirmation gate runs before any work, so it reports no
            // progress: OperationPhase.AwaitingConfirmation sorts after Snapshot
            // for a mutation flow, and the tracker requires monotonic phases.
            if (confirmScan is null || !confirmScan(preview))
            {
                return DirectDependentResult.Refused(
                    target,
                    scope,
                    AuditRefusalCodes.PreviewRequired,
                    $"Scanning {AuditPresentationLabels.Count(region.CellCount)} cells of worksheet '{region.WorksheetName}' was not confirmed, so nothing was read.");
            }
        }

        var names = port.CaptureNames(scope);
        var formulas = new List<AuditFormulaCell>();
        progress.Report(new OperationProgress(OperationPhase.Snapshot, 0, region.BlockCount, "Reading worksheet formulas."));

        for (var index = 0; index < region.BlockCount; index++)
        {
            if (progress.CancellationRequested)
            {
                progress.Report(new OperationProgress(OperationPhase.Cancelled, index, region.BlockCount, "Dependent scan cancelled."));
                return DirectDependentResult.Refused(
                    target,
                    scope,
                    AuditRefusalCodes.ScanCancelled,
                    "The dependent scan was cancelled; no partial scan is reported as a result.");
            }

            formulas.AddRange(port.CaptureBlock(scope, region.Block(index)));
            progress.Report(new OperationProgress(OperationPhase.Snapshot, index + 1, region.BlockCount, "Reading worksheet formulas."));
        }

        if (progress.CancellationRequested)
        {
            progress.Report(new OperationProgress(OperationPhase.Cancelled, region.BlockCount, region.BlockCount, "Dependent scan cancelled."));
            return DirectDependentResult.Refused(
                target,
                scope,
                AuditRefusalCodes.ScanCancelled,
                "The dependent scan was cancelled; no partial scan is reported as a result.");
        }

        progress.Report(new OperationProgress(OperationPhase.Analyze, 0, 1, "Building the bounded reverse index."));
        var result = ReverseReferenceIndex
            .Build(scope, formulas, names, FormulaDialect.InvariantA1)
            .FindDirectDependents(target);
        progress.Report(new OperationProgress(OperationPhase.Completed, 1, 1, "Dependent scan complete."));
        return result;
    }
}
