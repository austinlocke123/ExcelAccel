using System;
using ExcelAccel.Core.Performance;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class PerformanceBaselineTests
{
    [Fact]
    public void DistributionUsesNearestRankPercentilesAndSampleDeviation()
    {
        var distribution = new BaselineDistribution(new double[] { 9, 1, 7, 5, 3 });

        Assert.Equal(5, distribution.Count);
        Assert.Equal(1, distribution.Minimum);
        Assert.Equal(9, distribution.Maximum);
        Assert.Equal(5, distribution.Mean);
        Assert.Equal(5, distribution.Median);
        Assert.Equal(9, distribution.P95);
        Assert.Equal(Math.Sqrt(10), distribution.StandardDeviation, 10);
        Assert.Equal(new double[] { 1, 3, 5, 7, 9 }, distribution.Samples);
    }

    [Theory]
    [InlineData(new double[] { })]
    [InlineData(new double[] { 1 })]
    [InlineData(new double[] { -1, 1 })]
    [InlineData(new double[] { double.NaN, 1 })]
    [InlineData(new double[] { double.PositiveInfinity, 1 })]
    public void DistributionRejectsInvalidSamples(double[] samples)
    {
        Assert.ThrowsAny<ArgumentException>(() => new BaselineDistribution(samples));
    }

    [Fact]
    public void RegressionGatePassesAtExactTolerance()
    {
        var baseline = new BaselineDistribution(new double[] { 100, 100 });
        var candidate = new BaselineDistribution(new double[] { 115, 115 });

        var result = PerformanceRegressionGate.Evaluate(baseline, candidate, 0.15);

        Assert.True(result.Passed);
        Assert.Equal(0.15, result.ChangeRatio, 10);
    }

    [Fact]
    public void RegressionGateFailsAboveTolerance()
    {
        var baseline = new BaselineDistribution(new double[] { 100, 100 });
        var candidate = new BaselineDistribution(new double[] { 116, 116 });

        var result = PerformanceRegressionGate.Evaluate(baseline, candidate, 0.15);

        Assert.False(result.Passed);
        Assert.Contains("exceeds", result.Reason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RegressionGateHandlesZeroBaselineExplicitly()
    {
        var baseline = new BaselineDistribution(new double[] { 0, 0 });
        var candidate = new BaselineDistribution(new double[] { 0, 1 });

        var result = PerformanceRegressionGate.Evaluate(baseline, candidate, 0.15);

        Assert.False(result.Passed);
        Assert.True(double.IsPositiveInfinity(result.ChangeRatio));
    }

    [Theory]
    [InlineData(-0.01)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void RegressionGateRejectsInvalidTolerance(double tolerance)
    {
        var distribution = new BaselineDistribution(new double[] { 1, 2 });

        Assert.Throws<ArgumentOutOfRangeException>(
            () => PerformanceRegressionGate.Evaluate(distribution, distribution, tolerance));
    }
}
