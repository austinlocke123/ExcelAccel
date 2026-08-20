using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Auditing;
using ExcelAccel.Application.Operations;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.ModelCheck;

namespace ExcelAccel.Application.ModelCheck;

/// <summary>
/// Captures the cells a scan runs against. Read-only: it reports what the
/// worksheet holds and never writes, selects, or recalculates.
/// </summary>
public interface IModelCheckSnapshotPort
{
    /// <summary>The workbook and worksheet the active selection sits in.</summary>
    AuditCellIdentity CaptureTarget();

    /// <summary>The region the worksheet reports as used. Untrusted.</summary>
    UsedRegionBounds CaptureUsedRegion(string worksheetName);

    /// <summary>Reads the cells of exactly one banded rectangle.</summary>
    IReadOnlyList<ModelCheckCell> CaptureBlock(string worksheetName, AuditRectangle band);
}

/// <summary>The exact scope and configuration a scan ran with, so it can be repeated.</summary>
public sealed class ModelCheckScanRequest
{
    public ModelCheckScanRequest(
        ModelCheckScopeKind scope,
        AuditCellIdentity target,
        IReadOnlyList<string> ruleIds,
        ModelCheckConfiguration configuration)
    {
        Scope = scope;
        Target = target ?? throw new ArgumentNullException(nameof(target));
        RuleIds = ruleIds ?? throw new ArgumentNullException(nameof(ruleIds));
        Configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
    }

    public ModelCheckScopeKind Scope { get; }

    public AuditCellIdentity Target { get; }

    public IReadOnlyList<string> RuleIds { get; }

    public ModelCheckConfiguration Configuration { get; }

    /// <summary>
    /// Rebuilds the same request with a refreshed configuration, so a rescan
    /// repeats the exact prior scope and rule set against a new snapshot.
    /// </summary>
    public ModelCheckScanRequest WithConfiguration(ModelCheckConfiguration configuration) =>
        new ModelCheckScanRequest(Scope, Target, RuleIds, configuration);
}

public sealed class ModelCheckCoordinator
{
    /// <summary>Scanned-cell count above which the scan must be confirmed first.</summary>
    public const long PreviewThresholdCells = 25_000;

    public ModelCheckScanResult Execute(
        IModelCheckSnapshotPort port,
        ModelCheckScanRequest request,
        OperationProgressTracker? tracker = null,
        Func<ModelCheckScanPreview, bool>? confirmScan = null)
    {
        if (port is null) throw new ArgumentNullException(nameof(port));
        if (request is null) throw new ArgumentNullException(nameof(request));
        var progress = tracker ?? new OperationProgressTracker();
        var rules = ModelCheckRuleCatalog.Select(request.RuleIds);

        if (request.Scope == ModelCheckScopeKind.Workbook)
        {
            return Refused(
                request,
                ModelCheckRefusalCodes.ScopeTooLarge,
                "Workbook-scope scanning is not qualified yet; scan a selection or a worksheet.");
        }

        var worksheetName = request.Target.WorksheetName;
        IReadOnlyList<ModelCheckCell> cells;

        if (request.Scope == ModelCheckScopeKind.Selection)
        {
            if (!AuditAddress.TryParse(request.Target.Address, out var selection))
            {
                return Refused(request, ModelCheckRefusalCodes.ScopeTooLarge, "The selection is not a single A1 rectangle.");
            }

            cells = port.CaptureBlock(worksheetName, selection);
        }
        else
        {
            if (!DependentScanRegion.TryCreate(port.CaptureUsedRegion(worksheetName), out var region, out var code, out var message))
            {
                return Refused(request, code!, message!);
            }

            if (region!.CellCount > PreviewThresholdCells)
            {
                var preview = new ModelCheckScanPreview(worksheetName, region.CellCount, region.BlockCount, rules.Count);
                if (confirmScan is null || !confirmScan(preview))
                {
                    return Refused(
                        request,
                        ModelCheckRefusalCodes.PreviewRequired,
                        $"Scanning {region.CellCount:N0} cells of worksheet '{worksheetName}' was not confirmed, so nothing was read.");
                }
            }

            var collected = new List<ModelCheckCell>();
            progress.Report(new OperationProgress(OperationPhase.Snapshot, 0, region.BlockCount, "Reading worksheet cells."));
            for (var index = 0; index < region.BlockCount; index++)
            {
                if (progress.CancellationRequested)
                {
                    progress.Report(new OperationProgress(OperationPhase.Cancelled, index, region.BlockCount, "Scan cancelled."));
                    return Refused(request, ModelCheckRefusalCodes.ScanCancelled,
                        "The scan was cancelled; prior results remain and no partial scan is reported as complete.");
                }

                collected.AddRange(port.CaptureBlock(worksheetName, region.Block(index)));
                progress.Report(new OperationProgress(OperationPhase.Snapshot, index + 1, region.BlockCount, "Reading worksheet cells."));
            }

            cells = collected;
        }

        var snapshot = new ModelCheckSnapshot(request.Scope, request.Target.WorkbookId, cells);
        progress.Report(new OperationProgress(OperationPhase.Analyze, 0, Math.Max(rules.Count, 1), "Running rules."));
        var completedRules = 0;
        var result = new ModelCheckEngine().Run(
            snapshot,
            rules,
            request.Configuration,
            () => progress.CancellationRequested,
            _ =>
            {
                completedRules = Math.Min(completedRules + 1, Math.Max(rules.Count, 1));
                progress.Report(new OperationProgress(OperationPhase.Analyze, completedRules, Math.Max(rules.Count, 1), "Running rules."));
            });

        progress.Report(result.Status == AuditTraceStatus.Refused
            ? new OperationProgress(OperationPhase.Cancelled, completedRules, Math.Max(rules.Count, 1), "Scan cancelled.")
            : new OperationProgress(OperationPhase.Completed, 1, 1, "Scan complete."));
        return result;
    }

    private static ModelCheckScanResult Refused(ModelCheckScanRequest request, string code, string message) =>
        ModelCheckScanResult.Refused(
            new ModelCheckSnapshot(request.Scope, request.Target.WorkbookId, Array.Empty<ModelCheckCell>()),
            code,
            message);
}

public sealed class ModelCheckScanPreview
{
    public ModelCheckScanPreview(string worksheetName, long cellCount, int blockCount, int ruleCount)
    {
        WorksheetName = worksheetName;
        CellCount = cellCount;
        BlockCount = blockCount;
        RuleCount = ruleCount;
    }

    public string WorksheetName { get; }

    public long CellCount { get; }

    public int BlockCount { get; }

    public int RuleCount { get; }
}
