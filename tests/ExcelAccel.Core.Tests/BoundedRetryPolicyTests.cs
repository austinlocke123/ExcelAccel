using System;
using ExcelAccel.Core.Reliability;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class BoundedRetryPolicyTests
{
    [Fact]
    public void TransientFailuresRetryOnlyToTheConfiguredLimit()
    {
        var attempts = 0;
        var delays = 0;
        var policy = CreatePolicy(() => delays++);

        var result = policy.Execute(() =>
        {
            attempts++;
            if (attempts < 3)
            {
                throw new TransientTestException();
            }

            return 42;
        });

        Assert.Equal(42, result);
        Assert.Equal(3, attempts);
        Assert.Equal(2, delays);
    }

    [Fact]
    public void ExhaustedTransientFailureStopsDeterministically()
    {
        var attempts = 0;
        var delays = 0;
        var policy = CreatePolicy(() => delays++);

        Assert.Throws<TransientTestException>(() => policy.Execute<int>(() =>
        {
            attempts++;
            throw new TransientTestException();
        }));

        Assert.Equal(3, attempts);
        Assert.Equal(2, delays);
    }

    [Fact]
    public void NonTransientFailureIsNeverRetried()
    {
        var attempts = 0;
        var delays = 0;
        var policy = CreatePolicy(() => delays++);

        Assert.Throws<InvalidOperationException>(() => policy.Execute<int>(() =>
        {
            attempts++;
            throw new InvalidOperationException();
        }));

        Assert.Equal(1, attempts);
        Assert.Equal(0, delays);
    }

    private static BoundedRetryPolicy CreatePolicy(Action delay) =>
        new BoundedRetryPolicy(
            3,
            TimeSpan.FromMilliseconds(25),
            exception => exception is TransientTestException,
            _ => delay());

    private sealed class TransientTestException : Exception
    {
    }
}
