using System;

namespace ExcelAccel.Application.Formulas;

public sealed class FormulaSourceStore
{
    private readonly object _sync = new object();
    private FormulaBlockSnapshot? _snapshot;
    private DateTimeOffset _capturedUtc;

    public void Capture(FormulaBlockSnapshot snapshot, DateTimeOffset capturedUtc)
    {
        lock (_sync)
        {
            _snapshot = snapshot ?? throw new ArgumentNullException(nameof(snapshot));
            _capturedUtc = capturedUtc;
        }
    }

    public bool TryGet(DateTimeOffset now, TimeSpan maximumAge, out FormulaBlockSnapshot? snapshot, out string reason)
    {
        if (maximumAge <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(maximumAge));
        lock (_sync)
        {
            if (_snapshot is null)
            {
                snapshot = null;
                reason = "No formula source range is captured for this session.";
                return false;
            }
            if (now < _capturedUtc || now - _capturedUtc > maximumAge)
            {
                _snapshot = null;
                snapshot = null;
                reason = "The captured formula source expired; capture it again.";
                return false;
            }
            snapshot = _snapshot;
            reason = string.Empty;
            return true;
        }
    }

    public void Clear()
    {
        lock (_sync) _snapshot = null;
    }
}
