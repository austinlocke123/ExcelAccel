using System;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Formulas;
using ExcelAccel.Application.Formatting;
using ExcelAccel.ExcelInterop;
using ExcelDna.Integration;

namespace ExcelAccel.ExcelAddIn;

internal static class FormulaSourceRuntime
{
    private static readonly TimeSpan MaximumAge = TimeSpan.FromMinutes(30);
    public static FormulaSourceStore Store { get; } = new FormulaSourceStore();
    private static FormatBlockSnapshot? _formatSnapshot;
    private static DateTimeOffset _capturedAt;
    private static string _formatReason = "No internal format source has been captured in this Excel session.";

    public static CommandResult Capture()
    {
        var port = new ExcelSelectionAdapter(() => ExcelDnaUtil.Application, Reliability.RuntimeState.VerifyExcelThread);
        var snapshot = port.CaptureFormulaBlock();
        var now = DateTimeOffset.UtcNow;
        Store.Capture(snapshot, now);
        try
        {
            _formatSnapshot = port.CaptureFormatBlock();
            _capturedAt = now;
            _formatReason = string.Empty;
        }
        catch (CommandRefusedException exception)
        {
            _formatSnapshot = null;
            _formatReason = "Formula/value source captured, but formats-only paste is unavailable: " + exception.Message;
        }
        return CommandResult.Success("formula.source.capture",
            $"Captured {snapshot.Contents.RowCount:N0}x{snapshot.Contents.ColumnCount:N0} internal formula/value source for this Excel session." +
            (_formatSnapshot is null ? " " + _formatReason : " The bounded formatting source was also captured."));
    }

    public static bool TryGet(out FormulaBlockSnapshot? snapshot, out string reason) =>
        Store.TryGet(DateTimeOffset.UtcNow, MaximumAge, out snapshot, out reason);

    public static bool TryGetFormat(out FormatBlockSnapshot? snapshot, out string reason)
    {
        snapshot = null;
        if (_formatSnapshot is null) { reason = _formatReason; return false; }
        if (DateTimeOffset.UtcNow - _capturedAt > MaximumAge)
        {
            _formatSnapshot = null;
            reason = _formatReason = "The captured formatting source expired; capture it again.";
            return false;
        }
        snapshot = _formatSnapshot;
        reason = string.Empty;
        return true;
    }

    public static void Reset()
    {
        Store.Clear();
        _formatSnapshot = null;
        _capturedAt = default;
        _formatReason = "No internal format source has been captured in this Excel session.";
    }
}
