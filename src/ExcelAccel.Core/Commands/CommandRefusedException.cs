using System;

namespace ExcelAccel.Core.Commands;

public sealed class CommandRefusedException : Exception
{
    public CommandRefusedException(string message)
        : base(message)
    {
    }
}
