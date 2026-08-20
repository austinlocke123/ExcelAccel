#if DEBUG
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using ExcelAccel.Application.Auditing;
using ExcelAccel.Application.ModelCheck;
using ExcelAccel.Application.Operations;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.ModelCheck;
using ExcelAccel.ExcelAddIn.Reliability;
using ExcelAccel.ExcelInterop;
using ExcelAccel.Persistence.Diagnostics;
using ExcelDna.Integration;

namespace ExcelAccel.ExcelAddIn;

/// <summary>
/// Debug-only measurement hooks for WP-2-09 qualification.
///
/// Each hook times one Phase 2 operation inside the add-in, so the reported
/// milliseconds exclude the harness's own COM setup. They call the coordinators
/// directly rather than the registered routes, so no result view is created and
/// the measurement is of the analysis alone. Every operation remains read-only.
/// </summary>
public static class Phase2QualificationHooks
{
    [ExcelCommand(Name = "ExcelAccel.Perf.DirectPrecedents", Description = "Debug-only timed direct-precedent capture.")]
    public static string DirectPrecedents()
    {
        return Measure("perf.audit.precedents.direct", () =>
        {
            var port = new ExcelReferenceSnapshotAdapter(() => ExcelDnaUtil.Application, RuntimeState.VerifyExcelThread);
            var result = new DirectPrecedentCoordinator().Execute(port);
            return result.Status + ";" + result.Precedents.Count;
        });
    }

    [ExcelCommand(Name = "ExcelAccel.Perf.DependentScan", Description = "Debug-only timed worksheet dependent scan.")]
    public static string DependentScan()
    {
        return Measure("perf.audit.dependents.direct", () =>
        {
            var port = new ExcelDependentScanAdapter(() => ExcelDnaUtil.Application, RuntimeState.VerifyExcelThread);
            var result = new DirectDependentCoordinator().Execute(port, new OperationProgressTracker(), _ => true);
            return result.Status + ";" + result.Dependents.Count + ";" + result.ScannedFormulaCount;
        });
    }

    [ExcelCommand(Name = "ExcelAccel.Perf.IndirectPrecedents", Description = "Debug-only timed indirect precedent traversal.")]
    public static string IndirectPrecedents()
    {
        return Measure("perf.audit.precedents.indirect", () =>
        {
            var snapshot = new ExcelReferenceSnapshotAdapter(() => ExcelDnaUtil.Application, RuntimeState.VerifyExcelThread);
            var root = snapshot.CaptureTarget().Target;
            var result = new IndirectTraceCoordinator().Execute(
                root, new PrecedentTraceExpansion(snapshot), TraceDirection.Precedents, IndirectTraceOptions.Default);
            return result.Status + ";" + result.Nodes.Count + ";" + result.DeepestDepthReached;
        });
    }

    [ExcelCommand(Name = "ExcelAccel.Perf.ModelCheckWorksheet", Description = "Debug-only timed worksheet Model Check scan.")]
    public static string ModelCheckWorksheet()
    {
        return Measure("perf.model_check.worksheet", () =>
        {
            var result = RunModelCheck(null);
            return result.Status + ";" + result.Findings.Count + ";" + result.Snapshot.Cells.Count + ";" + result.RuleFailures.Count;
        });
    }

    [ExcelCommand(Name = "ExcelAccel.Perf.ModelCheckCancelled", Description = "Debug-only timed cancelled Model Check scan.")]
    public static string ModelCheckCancelled()
    {
        return Measure("perf.model_check.cancelled", () =>
        {
            var tracker = new OperationProgressTracker();
            tracker.RequestCancellation();
            var result = RunModelCheck(tracker);
            return result.Status + ";" + result.RefusalCode + ";" + result.Findings.Count;
        });
    }

    /// <summary>
    /// Exports the sanitized diagnostics and reports whether the seeded marker
    /// survived anywhere in the file. The marker is written into a formula, a
    /// value, a defined name, and the worksheet name by the harness.
    /// </summary>
    [ExcelCommand(Name = "ExcelAccel.Perf.PrivacyProbe", Description = "Debug-only diagnostics privacy probe.")]
    public static string PrivacyProbe(string needle)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(needle)) throw new ArgumentException("A needle is required.", nameof(needle));
            var destination = Path.Combine(Path.GetTempPath(), "excelaccel-privacy-" + Guid.NewGuid().ToString("N") + ".txt");
            var log = File.Exists(DiagnosticLog.LogPath) ? File.ReadAllBytes(DiagnosticLog.LogPath) : new byte[0];
            var exporter = new DiagnosticExporter();
            var plan = exporter.Plan(destination, log);
            exporter.Export(plan, plan.PlanHash, log);

            var text = File.ReadAllText(destination, Encoding.UTF8);
            var contains = text.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
            var summary = (contains ? "leaked" : "clean") + "|" + plan.ByteCount.ToString(CultureInfo.InvariantCulture);
            try { File.Delete(destination); } catch (IOException) { }
            DiagnosticLog.Info("perf.privacy.probe", contains ? "leaked" : "clean");
            return summary;
        }
        catch (Exception exception)
        {
            DiagnosticLog.Error("perf.privacy.probe", exception);
            throw;
        }
    }

    private static ModelCheckScanResult RunModelCheck(OperationProgressTracker? tracker)
    {
        var port = new ExcelModelCheckSnapshotAdapter(() => ExcelDnaUtil.Application, RuntimeState.VerifyExcelThread);
        var request = new ModelCheckScanRequest(
            ModelCheckScopeKind.Worksheet,
            port.CaptureTarget(),
            ModelCheckRuleCatalog.AllRuleIds,
            ModelCheckConfiguration.Default);
        return new ModelCheckCoordinator().Execute(port, request, tracker, _ => true);
    }

    /// <summary>Returns "elapsedMs|detail" with the elapsed time measured in-process.</summary>
    private static string Measure(string operationId, Func<string> operation)
    {
        try
        {
            var stopwatch = Stopwatch.StartNew();
            var detail = operation();
            stopwatch.Stop();
            var summary = stopwatch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + "|" + detail;
            DiagnosticLog.Info(operationId, summary);
            return summary;
        }
        catch (Exception exception)
        {
            DiagnosticLog.Error(operationId, exception);
            throw;
        }
    }
}
#endif
