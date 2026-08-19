using System;

namespace ExcelAccel.Core.Performance;

public sealed class PerformanceRegressionResult
{
    internal PerformanceRegressionResult(bool passed, double changeRatio, string reason)
    {
        Passed = passed;
        ChangeRatio = changeRatio;
        Reason = reason;
    }

    public bool Passed { get; }

    public double ChangeRatio { get; }

    public string Reason { get; }
}

public static class PerformanceRegressionGate
{
    public static PerformanceRegressionResult Evaluate(
        BaselineDistribution baseline,
        BaselineDistribution candidate,
        double allowedP95RegressionRatio)
    {
        if (baseline is null)
        {
            throw new ArgumentNullException(nameof(baseline));
        }

        if (candidate is null)
        {
            throw new ArgumentNullException(nameof(candidate));
        }

        if (allowedP95RegressionRatio < 0 ||
            double.IsNaN(allowedP95RegressionRatio) ||
            double.IsInfinity(allowedP95RegressionRatio))
        {
            throw new ArgumentOutOfRangeException(nameof(allowedP95RegressionRatio));
        }

        if (baseline.P95 == 0)
        {
            bool unchanged = candidate.P95 == 0;
            return new PerformanceRegressionResult(
                unchanged,
                unchanged ? 0 : double.PositiveInfinity,
                unchanged
                    ? "Both P95 values are zero."
                    : "A zero-duration baseline cannot establish a finite regression ratio.");
        }

        double changeRatio = (candidate.P95 - baseline.P95) / baseline.P95;
        bool passed = changeRatio <= allowedP95RegressionRatio;
        return new PerformanceRegressionResult(
            passed,
            changeRatio,
            passed
                ? "Candidate P95 is within the allowed regression tolerance."
                : "Candidate P95 exceeds the allowed regression tolerance.");
    }
}
