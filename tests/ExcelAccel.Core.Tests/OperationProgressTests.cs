using System;
using ExcelAccel.Application.Operations;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class OperationProgressTests
{
    [Fact]
    public void CancellationIsHonoredBeforeCommitAndRefusedDuringCommit()
    {
        var tracker = new OperationProgressTracker();
        tracker.Report(new OperationProgress(OperationPhase.Analyze, 5, 10, "Analyzing"));
        Assert.True(tracker.RequestCancellation());
        Assert.Throws<OperationCanceledException>(tracker.ThrowIfCancellationRequested);

        var committing = new OperationProgressTracker();
        committing.Report(new OperationProgress(OperationPhase.Commit, 0, 1, "Committing"));
        Assert.False(committing.RequestCancellation());
        committing.ThrowIfCancellationRequested();
    }

    [Fact]
    public void ProgressCannotMoveBackwardOrExceedBound()
    {
        var tracker = new OperationProgressTracker();
        tracker.Report(new OperationProgress(OperationPhase.Analyze, 5, 10, "Analyzing"));
        Assert.Throws<InvalidOperationException>(() => tracker.Report(new OperationProgress(OperationPhase.Analyze, 4, 10, "Backward")));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OperationProgress(OperationPhase.Analyze, 11, 10, "Invalid"));
    }
}
