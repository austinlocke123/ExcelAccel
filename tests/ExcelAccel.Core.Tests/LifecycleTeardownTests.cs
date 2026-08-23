using System;
using System.Collections.Generic;
using ExcelAccel.Application.Reliability;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class LifecycleTeardownTests
{
    private static KeyValuePair<string, Action> Step(string name, Action action) =>
        new KeyValuePair<string, Action>(name, action);

    [Fact]
    public void EveryStepRunsAndTheFinalStepRunsOnASilentPath()
    {
        var order = new List<string>();

        var failures = LifecycleTeardown.Run(
            new[] { Step("a", () => order.Add("a")), Step("b", () => order.Add("b")) },
            () => order.Add("always"),
            (_, __) => throw new InvalidOperationException("No failure was expected."));

        Assert.Equal(0, failures);
        Assert.Equal(new[] { "a", "b", "always" }, order);
    }

    /// <summary>
    /// The defect this class exists for. A reset that threw used to skip every
    /// later reset and the marker cleanup with it, leaving a stale session marker
    /// that put the next Excel session into safe mode.
    /// </summary>
    [Fact]
    public void AThrowingStepDoesNotStopTheOnesAfterItOrTheFinalStep()
    {
        var order = new List<string>();
        var reported = new List<string>();

        var failures = LifecycleTeardown.Run(
            new[]
            {
                Step("first", () => order.Add("first")),
                Step("boom", () => throw new InvalidOperationException("Injected reset failure.")),
                Step("last", () => order.Add("last")),
            },
            () => order.Add("always"),
            (name, _) => reported.Add(name));

        Assert.Equal(1, failures);
        Assert.Equal(new[] { "first", "last", "always" }, order);
        Assert.Equal(new[] { "boom" }, reported);
    }

    [Fact]
    public void EveryStepFailingStillLeavesTheFinalStepRun()
    {
        var always = 0;

        var failures = LifecycleTeardown.Run(
            new[]
            {
                Step("one", () => throw new InvalidOperationException("one")),
                Step("two", () => throw new InvalidOperationException("two")),
            },
            () => always++,
            (_, __) => { });

        Assert.Equal(2, failures);
        Assert.Equal(1, always);
    }

    [Fact]
    public void AFailureInTheFinalStepIsCountedRatherThanThrown()
    {
        var reported = new List<string>();

        var failures = LifecycleTeardown.Run(
            new[] { Step("ok", () => { }) },
            () => throw new InvalidOperationException("Injected cleanup failure."),
            (name, _) => reported.Add(name));

        Assert.Equal(1, failures);
        Assert.Equal(new[] { "always" }, reported);
    }

    /// <summary>
    /// A reporter that throws must not defeat the sequence it is reporting on.
    /// </summary>
    [Fact]
    public void AThrowingReporterCannotAbandonTheTeardown()
    {
        var always = 0;

        var failures = LifecycleTeardown.Run(
            new[]
            {
                Step("boom", () => throw new InvalidOperationException("Injected.")),
                Step("after", () => { }),
            },
            () => always++,
            (_, __) => throw new InvalidOperationException("Injected reporter failure."));

        Assert.Equal(1, failures);
        Assert.Equal(1, always);
    }

    [Fact]
    public void ANullStepIsSkippedRatherThanThrowing()
    {
        var always = 0;

        var failures = LifecycleTeardown.Run(
            new[] { Step("null", null!) },
            () => always++,
            (_, __) => { });

        Assert.Equal(0, failures);
        Assert.Equal(1, always);
    }
}
