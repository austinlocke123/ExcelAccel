using System;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.Formulas;

namespace ExcelAccel.Application.Auditing;

public sealed class FormulaTargetCapture
{
    public FormulaTargetCapture(AuditCellIdentity target, string formula, FormulaDialect? dialect = null)
    {
        Target = target ?? throw new ArgumentNullException(nameof(target));
        Formula = formula ?? throw new ArgumentNullException(nameof(formula));
        Dialect = dialect ?? FormulaDialect.InvariantA1;
    }

    public AuditCellIdentity Target { get; }
    public string Formula { get; }
    public FormulaDialect Dialect { get; }
}

public interface IDirectPrecedentSnapshotPort
{
    FormulaTargetCapture CaptureTarget();
    ReferenceSnapshotIndex CaptureIndex(DirectPrecedentCapturePlan plan);
    bool SourceMatches(FormulaTargetCapture capture);
}

public enum WorkbookPresence
{
    Open,
    Closed,
    Unknown,
}

/// <summary>
/// Read-only probe used by result presentation to discard a captured analysis
/// once its source workbook is no longer open. It never opens a workbook.
/// </summary>
public interface IWorkbookPresencePort
{
    WorkbookPresence Probe(string workbookId);
}

public sealed class DirectPrecedentCoordinator
{
    private readonly DirectPrecedentAnalyzer _analyzer = new DirectPrecedentAnalyzer();

    public DirectPrecedentResult Execute(IDirectPrecedentSnapshotPort port)
    {
        if (port is null) throw new ArgumentNullException(nameof(port));
        var capture = port.CaptureTarget();
        var plan = _analyzer.CreateCapturePlan(capture.Target, capture.Formula, capture.Dialect);
        var index = port.CaptureIndex(plan);
        if (!port.SourceMatches(capture))
        {
            return DirectPrecedentResult.Refused(
                capture.Target,
                AuditRefusalCodes.StaleTarget,
                "The source formula changed during precedent capture; retry the command.");
        }
        return _analyzer.Analyze(new FormulaReferenceSnapshot(capture.Target, capture.Formula, index, capture.Dialect));
    }
}
