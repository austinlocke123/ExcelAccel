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

    /// <summary>Every worksheet in the workbook, for a workbook-scope plan.</summary>
    IReadOnlyList<string> CaptureWorksheetNames();

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

        var worksheetName = request.Target.WorksheetName;
        IReadOnlyList<ModelCheckCell> cells;
        var excludedWorksheets = 0;

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
            var worksheets = request.Scope == ModelCheckScopeKind.Workbook
                ? port.CaptureWorksheetNames()
                : new[] { worksheetName };

            if (!WorkbookScanPlan.TryCreate(
                    request.Target.WorkbookId,
                    worksheets.Select(port.CaptureUsedRegion),
                    out var plan,
                    out var code,
                    out var message))
            {
                return Refused(request, code!, message!);
            }

            excludedWorksheets = plan!.Excluded.Count;

            // A workbook scan always confirms and always shows its sheet
            // inventory; a worksheet scan confirms only above the threshold.
            if (request.Scope == ModelCheckScopeKind.Workbook || plan.TotalCellCount > PreviewThresholdCells)
            {
                var preview = new ModelCheckScanPreview(
                    request.Scope == ModelCheckScopeKind.Workbook ? request.Target.WorkbookId : worksheetName,
                    plan.TotalCellCount,
                    plan.TotalBlockCount,
                    rules.Count,
                    plan.InventoryLines());
                if (confirmScan is null || !confirmScan(preview))
                {
                    return Refused(
                        request,
                        ModelCheckRefusalCodes.PreviewRequired,
                        "Scanning " + plan.TotalCellCount.ToString("N0") + " cells across " +
                            plan.Included.Count.ToString("N0") + " worksheets was not confirmed, so nothing was read.");
                }
            }

            var collected = new List<ModelCheckCell>();
            var totalBlocks = Math.Max(plan.TotalBlockCount, 1);
            var readBlocks = 0;
            progress.Report(new OperationProgress(OperationPhase.Snapshot, 0, totalBlocks, "Reading cells."));
            foreach (var entry in plan.Included)
            {
                for (var index = 0; index < entry.BlockCount; index++)
                {
                    if (progress.CancellationRequested)
                    {
                        progress.Report(new OperationProgress(OperationPhase.Cancelled, readBlocks, totalBlocks, "Scan cancelled."));
                        return Refused(request, ModelCheckRefusalCodes.ScanCancelled,
                            "The scan was cancelled; prior results remain and no partial scan is reported as complete.");
                    }

                    collected.AddRange(port.CaptureBlock(entry.WorksheetName, entry.Region!.Block(index)));
                    readBlocks++;
                    progress.Report(new OperationProgress(OperationPhase.Snapshot, readBlocks, totalBlocks, "Reading cells."));
                }
            }

            cells = collected;
        }

        var snapshot = new ModelCheckSnapshot(request.Scope, request.Target.WorkbookId, cells, excludedWorksheets);
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
    public ModelCheckScanPreview(
        string scopeName,
        long cellCount,
        int blockCount,
        int ruleCount,
        IReadOnlyList<string>? inventoryLines = null)
    {
        WorksheetName = scopeName;
        CellCount = cellCount;
        BlockCount = blockCount;
        RuleCount = ruleCount;
        InventoryLines = inventoryLines ?? Array.Empty<string>();
    }

    /// <summary>The per-worksheet inventory a workbook scan must show.</summary>
    public IReadOnlyList<string> InventoryLines { get; }

    public string WorksheetName { get; }

    public long CellCount { get; }

    public int BlockCount { get; }

    public int RuleCount { get; }
}
