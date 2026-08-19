using System;

namespace ExcelAccel.Core.Commands;

public sealed class CommandRefusedException : Exception
{
    public CommandRefusedException(string message)
        : this(RefusalCodes.SelectionUnsupported, message, string.Empty)
    {
    }

    public CommandRefusedException(string refusalCode, string message, string remediation)
        : base(message)
    {
        RefusalCode = refusalCode;
        Remediation = remediation;
    }

    public string RefusalCode { get; }

    public string Remediation { get; }
}
