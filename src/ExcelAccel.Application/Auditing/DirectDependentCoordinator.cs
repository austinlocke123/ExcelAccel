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

public sealed class DirectDependentCoordinator
{
    public DirectDependentResult Execute(IDependentScanPort port, OperationProgressTracker? tracker = null)
    {
        if (port is null) throw new ArgumentNullException(nameof(port));
        var progress = tracker ?? new OperationProgressTracker();
        var target = port.CaptureTarget();
        var scope = DependentScanScope.Worksheet(target.WorkbookId, target.WorksheetName);

        if (!DependentScanRegion.TryCreate(port.CaptureUsedRegion(scope), out var region, out var refusalCode, out var message))
        {
            return DirectDependentResult.Refused(target, scope, refusalCode!, message!);
        }

        var names = port.CaptureNames(scope);
        var formulas = new List<AuditFormulaCell>();
        progress.Report(new OperationProgress(OperationPhase.Snapshot, 0, region!.BlockCount, "Reading worksheet formulas."));

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
