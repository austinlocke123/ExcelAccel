using System;
using System.Collections.Generic;

namespace ExcelAccel.Application.Reliability;

/// <summary>
/// Runs a teardown sequence so that no single failure can abandon the rest of it.
/// </summary>
/// <remarks>
/// This exists because the add-in's unload path previously ran its resets inline
/// inside one try block. A reset that threw skipped every later reset <b>and</b>
/// the final marker cleanup, which left a stale session marker behind. The next
/// Excel session then started in safe mode and refused every mutation command,
/// for a reason that looked nothing like the actual cause.
///
/// The logic lives in the application layer rather than the host so it can be
/// fault-injected in tests; the host has no test project.
/// </remarks>
public static class LifecycleTeardown
{
    /// <summary>
    /// Runs every step, then <paramref name="always"/>, whatever happens.
    /// Returns the number of steps that failed.
    /// </summary>
    /// <param name="steps">Named teardown steps, run in order.</param>
    /// <param name="always">
    /// Runs exactly once after every step has been attempted, including when
    /// steps threw. A failure here is reported but never hidden.
    /// </param>
    /// <param name="onFailure">Receives the name and exception of each failure.</param>
    public static int Run(
        IEnumerable<KeyValuePair<string, Action>> steps,
        Action always,
        Action<string, Exception> onFailure)
    {
        if (steps is null) throw new ArgumentNullException(nameof(steps));
        if (always is null) throw new ArgumentNullException(nameof(always));
        if (onFailure is null) throw new ArgumentNullException(nameof(onFailure));

        var failures = 0;
        try
        {
            foreach (var step in steps)
            {
                try
                {
                    step.Value?.Invoke();
                }
                catch (Exception exception)
                {
                    failures++;
                    Report(onFailure, step.Key, exception);
                }
            }
        }
        finally
        {
            try
            {
                always();
            }
            catch (Exception exception)
            {
                failures++;
                Report(onFailure, "always", exception);
            }
        }

        return failures;
    }

    /// <summary>
    /// A reporter that throws would defeat the whole point, so it is contained.
    /// </summary>
    private static void Report(Action<string, Exception> onFailure, string name, Exception exception)
    {
        try { onFailure(name, exception); }
        catch { }
    }
}
