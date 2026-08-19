using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAccel.Core.Performance;

public sealed class BaselineDistribution
{
    private readonly double[] _sortedSamples;

    public BaselineDistribution(IEnumerable<double> samples)
    {
        if (samples is null)
        {
            throw new ArgumentNullException(nameof(samples));
        }

        _sortedSamples = samples.ToArray();
        if (_sortedSamples.Length < 2)
        {
            throw new ArgumentException("A distribution requires at least two samples.", nameof(samples));
        }

        if (_sortedSamples.Any(sample => sample < 0 || double.IsNaN(sample) || double.IsInfinity(sample)))
        {
            throw new ArgumentException("Samples must be finite, non-negative values.", nameof(samples));
        }

        Array.Sort(_sortedSamples);
        Minimum = _sortedSamples[0];
        Maximum = _sortedSamples[_sortedSamples.Length - 1];
        Mean = _sortedSamples.Average();
        Median = Percentile(0.50);
        P95 = Percentile(0.95);

        double squaredDeviationTotal = 0;
        foreach (double sample in _sortedSamples)
        {
            double deviation = sample - Mean;
            squaredDeviationTotal += deviation * deviation;
        }

        StandardDeviation = Math.Sqrt(squaredDeviationTotal / (_sortedSamples.Length - 1));
        CoefficientOfVariation = Mean == 0 ? 0 : StandardDeviation / Mean;
    }

    public int Count => _sortedSamples.Length;

    public double Minimum { get; }

    public double Maximum { get; }

    public double Mean { get; }

    public double Median { get; }

    public double P95 { get; }

    public double StandardDeviation { get; }

    public double CoefficientOfVariation { get; }

    public IReadOnlyList<double> Samples => Array.AsReadOnly(_sortedSamples);

    public double Percentile(double probability)
    {
        if (probability <= 0 || probability > 1 || double.IsNaN(probability))
        {
            throw new ArgumentOutOfRangeException(nameof(probability));
        }

        int nearestRank = (int)Math.Ceiling(probability * _sortedSamples.Length);
        return _sortedSamples[Math.Max(0, nearestRank - 1)];
    }
}
