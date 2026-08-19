using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAccel.Core.Reliability;

public sealed class StateRestoreException : Exception
{
    public StateRestoreException(string message, IEnumerable<string> failedProperties, Exception? innerException = null)
        : base(message, innerException)
    {
        FailedProperties = (failedProperties ?? throw new ArgumentNullException(nameof(failedProperties))).ToArray();
    }

    public IReadOnlyList<string> FailedProperties { get; }
}
