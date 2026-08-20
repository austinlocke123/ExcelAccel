using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ExcelAccel.Application.Auditing;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.ModelCheck;
using ExcelAccel.Application.Operations;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.ModelCheck;
using ExcelAccel.ExcelAddIn.Reliability;
using ExcelAccel.ExcelInterop;
using ExcelAccel.Persistence.ModelCheck;
using ExcelDna.Integration;

namespace ExcelAccel.ExcelAddIn;

/// <summary>
/// Session state for Model Check: the last scan's exact scope so a rescan can
/// repeat it, the last result so findings can be ignored or exported, and the
/// local ignore set.
/// </summary>
internal static class ModelCheckRuntime
{
    private static readonly object Sync = new object();
    private static readonly TraceViewRuntime View = new TraceViewRuntime(
        ModelCheckCommandCatalog.RunSelectionId,
        "Read-only Model Check findings. This view never changes the workbook.",
        CommandDispatcher.NavigateToTraceTarget);

    private static ModelCheckScanRequest? _lastRequest;
    private static ModelCheckScanResult? _lastResult;

    public static bool IsOpen => View.IsOpen;

    public static string IgnorePath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "ExcelAccel",
        "model-check-ignores.tsv");

    public static ModelCheckScanResult? LastResult
    {
        get { lock (Sync) return _lastResult; }
    }

    public static ModelCheckScanRequest? LastRequest
    {
        get { lock (Sync) return _lastRequest; }
    }

    public static IReadOnlyList<ModelCheckIgnoreEntry> LoadIgnores()
    {
        try
        {
            return new ModelCheckIgnoreStore().Load(IgnorePath);
        }
        catch (IOException exception)
        {
            DiagnosticLog.Error("model_check.ignores.load", exception);
            return Array.Empty<ModelCheckIgnoreEntry>();
        }
    }

    public static ModelCheckConfiguration CurrentConfiguration() =>
        ModelCheckConfiguration.Default.WithIgnoredFingerprints(LoadIgnores().Select(entry => entry.Fingerprint));

    public static CommandResult Run(ModelCheckScopeKind scope)
    {
        var port = new ExcelModelCheckSnapshotAdapter(() => ExcelDnaUtil.Application, RuntimeState.VerifyExcelThread);
        var request = new ModelCheckScanRequest(
            scope,
            port.CaptureTarget(),
            ModelCheckRuleCatalog.AllRuleIds,
            CurrentConfiguration());
        return Execute(port, request);
    }

    /// <summary>
    /// Repeats the exact prior scope and rule configuration against a newly
    /// captured snapshot. Prior findings are never reused as current evidence.
    /// </summary>
    public static CommandResult Rescan()
    {
        ModelCheckScanRequest? previous;
        lock (Sync) previous = _lastRequest;
        if (previous is null)
        {
            return CommandResult.Refused(
                ModelCheckCommandCatalog.RescanId,
                "There is no prior Model Check scan to repeat in this session.",
                RefusalCodes.CommandUnavailable);
        }

        var port = new ExcelModelCheckSnapshotAdapter(() => ExcelDnaUtil.Application, RuntimeState.VerifyExcelThread);
        return Execute(port, previous.WithConfiguration(CurrentConfiguration()));
    }

    private static CommandResult Execute(IModelCheckSnapshotPort port, ModelCheckScanRequest request)
    {
        var tracker = new OperationProgressTracker();
        var result = new ModelCheckCoordinator().Execute(port, request, tracker, ConfirmScan);
        var commandId = request.Scope == ModelCheckScopeKind.Selection
            ? ModelCheckCommandCatalog.RunSelectionId
            : ModelCheckCommandCatalog.RunWorksheetId;
        DiagnosticLog.Info(
            commandId,
            $"status:{result.Status};findings:{result.Findings.Count};cells:{result.Snapshot.Cells.Count};" +
            $"suppressed:{result.SuppressedFindingCount};failures:{result.RuleFailures.Count};phase:{tracker.Current.Phase}");

        // A prior result is replaced only when a scan completes. A refused scan
        // leaves the previous findings on screen and marked as prior.
        if (result.Status != AuditTraceStatus.Refused)
        {
            lock (Sync)
            {
                _lastRequest = request;
                _lastResult = result;
            }
        }

        var presence = new ExcelReferenceSnapshotAdapter(() => ExcelDnaUtil.Application, RuntimeState.VerifyExcelThread);
        return View.Present(ModelCheckReport.Create(result).ToPresentation(), result.Snapshot.WorkbookId, presence);
    }

    private static bool ConfirmScan(ModelCheckScanPreview preview)
    {
        var inventory = preview.InventoryLines.Count == 0
            ? string.Empty
            : Environment.NewLine + Environment.NewLine + string.Join(Environment.NewLine, preview.InventoryLines);
        var message =
            $"Run {preview.RuleCount:N0} Model Check rules over '{preview.WorksheetName}'?" +
            Environment.NewLine + Environment.NewLine +
            $"The scan reads {preview.CellCount:N0} cells in {preview.BlockCount:N0} bounded blocks. It is read-only and changes nothing." +
            inventory;
        var owner = ExcelWindowOwner.TryCreate();
        var answer = owner is null
            ? MessageBox.Show(message, "ExcelAccel", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
            : MessageBox.Show(owner, message, "ExcelAccel", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
        return answer == DialogResult.OK;
    }

    public static CommandResult IgnoreSelected()
    {
        var result = LastResult;
        if (result is null || result.Findings.Count == 0)
        {
            return CommandResult.Refused(
                ModelCheckCommandCatalog.IgnoreLocalId,
                "There are no current Model Check findings to ignore.",
                RefusalCodes.CommandUnavailable);
        }

        using (var dialog = new ModelCheckIgnoreDialog(result.Findings, LoadIgnores()))
        {
            var owner = ExcelWindowOwner.TryCreate();
            if ((owner is null ? dialog.ShowDialog() : dialog.ShowDialog(owner)) != DialogResult.OK)
            {
                return CommandResult.Refused(
                    ModelCheckCommandCatalog.IgnoreLocalId, "The ignore was not confirmed.", RefusalCodes.PreviewRequired);
            }

            var entries = LoadIgnores().ToList();
            entries.AddRange(dialog.SelectedFindings.Select(finding => new ModelCheckIgnoreEntry(
                finding.RuleId,
                finding.RuleVersion,
                finding.Fingerprint,
                AuditPresentationLabels.Location(finding.Target))));
            new ModelCheckIgnoreStore().SaveAtomic(IgnorePath, entries);
            return CommandResult.Success(
                ModelCheckCommandCatalog.IgnoreLocalId,
                $"Added {dialog.SelectedFindings.Count:N0} local ignore(s). Rescan for them to take effect.");
        }
    }

    public static CommandResult ManageIgnores()
    {
        var ignores = LoadIgnores();
        if (ignores.Count == 0)
        {
            return CommandResult.Success(ModelCheckCommandCatalog.UnignoreLocalId, "There are no active local ignores.");
        }

        using (var dialog = new ModelCheckUnignoreDialog(ignores))
        {
            var owner = ExcelWindowOwner.TryCreate();
            if ((owner is null ? dialog.ShowDialog() : dialog.ShowDialog(owner)) != DialogResult.OK)
            {
                return CommandResult.Refused(
                    ModelCheckCommandCatalog.UnignoreLocalId, "The removal was not confirmed.", RefusalCodes.PreviewRequired);
            }

            var remaining = ignores.Where(entry => !dialog.RemovedFingerprints.Contains(entry.Fingerprint)).ToArray();
            new ModelCheckIgnoreStore().SaveAtomic(IgnorePath, remaining);
            return CommandResult.Success(
                ModelCheckCommandCatalog.UnignoreLocalId,
                $"Removed {dialog.RemovedFingerprints.Count:N0} local ignore(s). Rescan for the change to take effect.");
        }
    }

    public static CommandResult Export()
    {
        var result = LastResult;
        if (result is null || result.Findings.Count == 0)
        {
            return CommandResult.Refused(
                ModelCheckCommandCatalog.ExportId,
                "There are no current Model Check findings to export.",
                RefusalCodes.CommandUnavailable);
        }

        using (var save = new SaveFileDialog
        {
            Title = "Export Model Check findings",
            Filter = "CSV file (*.csv)|*.csv",
            DefaultExt = "csv",
        })
        {
            if (save.ShowDialog() != DialogResult.OK)
            {
                return CommandResult.Refused(ModelCheckCommandCatalog.ExportId, "No destination was chosen.", RefusalCodes.PreviewRequired);
            }

            var exporter = new ModelCheckResultExporter();
            var findings = result.Findings.Select(finding => new ExportableFinding(
                finding.RuleId,
                finding.RuleVersion,
                finding.Severity.ToString(),
                finding.Target.WorksheetName,
                finding.Target.Address,
                finding.Statement,
                finding.Coverage.ToString(),
                finding.Fingerprint,
                finding.Evidence)).ToArray();

            using (var manifestDialog = new ModelCheckExportManifestDialog(exporter.Plan(save.FileName, findings, false)))
            {
                var owner = ExcelWindowOwner.TryCreate();
                if ((owner is null ? manifestDialog.ShowDialog() : manifestDialog.ShowDialog(owner)) != DialogResult.OK)
                {
                    return CommandResult.Refused(
                        ModelCheckCommandCatalog.ExportId, "The export manifest was not confirmed.", RefusalCodes.PreviewRequired);
                }

                var manifest = exporter.Plan(save.FileName, findings, manifestDialog.IncludeEvidence);
                exporter.Export(manifest, findings);
                return CommandResult.Success(
                    ModelCheckCommandCatalog.ExportId,
                    $"Exported {findings.Length:N0} findings to {manifest.Destination}.");
            }
        }
    }

    public static void Reset()
    {
        View.Reset();
        lock (Sync)
        {
            _lastRequest = null;
            _lastResult = null;
        }
    }
}
