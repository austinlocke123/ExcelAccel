using System;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelAccel.Core.Reliability;

namespace ExcelAccel.ExcelAddIn.Interop;

internal static class ExcelComRetry
{
    private static readonly BoundedRetryPolicy Policy = new BoundedRetryPolicy(
        maximumAttempts: 3,
        delay: TimeSpan.FromMilliseconds(25),
        isTransient: IsTransient,
        delayAction: Thread.Sleep);

    public static T Execute<T>(Func<T> operation) => Policy.Execute(operation);

    public static void Execute(Action operation) => Policy.Execute(operation);

    private static bool IsTransient(Exception exception)
    {
        if (!(exception is COMException comException))
        {
            return false;
        }

        return comException.HResult == unchecked((int)0x80010001) ||
               comException.HResult == unchecked((int)0x8001010A) ||
               comException.HResult == unchecked((int)0x800AC472);
    }
}
