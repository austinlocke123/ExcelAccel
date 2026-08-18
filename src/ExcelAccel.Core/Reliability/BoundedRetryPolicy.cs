using System;

namespace ExcelAccel.Core.Reliability;

public sealed class BoundedRetryPolicy
{
    private readonly int _maximumAttempts;
    private readonly TimeSpan _delay;
    private readonly Func<Exception, bool> _isTransient;
    private readonly Action<TimeSpan> _delayAction;

    public BoundedRetryPolicy(
        int maximumAttempts,
        TimeSpan delay,
        Func<Exception, bool> isTransient,
        Action<TimeSpan> delayAction)
    {
        if (maximumAttempts < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(maximumAttempts));
        }

        if (delay < TimeSpan.Zero)
        {
            throw new ArgumentOutOfRangeException(nameof(delay));
        }

        _maximumAttempts = maximumAttempts;
        _delay = delay;
        _isTransient = isTransient ?? throw new ArgumentNullException(nameof(isTransient));
        _delayAction = delayAction ?? throw new ArgumentNullException(nameof(delayAction));
    }

    public T Execute<T>(Func<T> operation)
    {
        if (operation is null)
        {
            throw new ArgumentNullException(nameof(operation));
        }

        for (var attempt = 1; ; attempt++)
        {
            try
            {
                return operation();
            }
            catch (Exception exception) when (attempt < _maximumAttempts && _isTransient(exception))
            {
                _delayAction(_delay);
            }
        }
    }

    public void Execute(Action operation)
    {
        Execute(() =>
        {
            operation();
            return true;
        });
    }
}
