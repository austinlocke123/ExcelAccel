using System;

namespace ExcelAccel.Application.Operations;

public enum OperationPhase { Snapshot, Analyze, AwaitingConfirmation, Commit, Verify, Completed, Failed, Cancelled }

public sealed class OperationProgress
{
    public OperationProgress(OperationPhase phase, long completed, long total, string status)
    {
        if (completed < 0 || total < 0 || completed > total) throw new ArgumentOutOfRangeException(nameof(completed));
        Phase = phase;
        Completed = completed;
        Total = total;
        Status = status ?? string.Empty;
    }
    public OperationPhase Phase { get; }
    public long Completed { get; }
    public long Total { get; }
    public string Status { get; }
}

public sealed class OperationProgressTracker
{
    private readonly object _sync = new object();
    private OperationProgress _current = new OperationProgress(OperationPhase.Snapshot, 0, 0, string.Empty);
    private bool _cancellationRequested;
    public event Action<OperationProgress>? Changed;
    public OperationProgress Current { get { lock (_sync) return _current; } }
    public bool CancellationRequested { get { lock (_sync) return _cancellationRequested; } }

    public bool RequestCancellation()
    {
        lock (_sync)
        {
            if (_current.Phase >= OperationPhase.Commit && _current.Phase <= OperationPhase.Completed) return false;
            _cancellationRequested = true;
            return true;
        }
    }

    public void Report(OperationProgress next)
    {
        if (next is null) throw new ArgumentNullException(nameof(next));
        Action<OperationProgress>? changed;
        lock (_sync)
        {
            if (next.Phase < _current.Phase || (next.Phase == _current.Phase && next.Completed < _current.Completed))
                throw new InvalidOperationException("Operation progress must be monotonic.");
            _current = next;
            changed = Changed;
        }
        changed?.Invoke(next);
    }

    public void ThrowIfCancellationRequested()
    {
        lock (_sync)
        {
            if (_cancellationRequested && _current.Phase < OperationPhase.Commit) throw new OperationCanceledException("The operation was cancelled before commit.");
        }
    }
}
