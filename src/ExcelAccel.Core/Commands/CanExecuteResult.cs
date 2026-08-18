using System;

namespace ExcelAccel.Core.Commands;

public sealed class CanExecuteResult
{
    private CanExecuteResult(bool allowed, string refusalCode, string message, string remediation)
    {
        Allowed = allowed;
        RefusalCode = refusalCode;
        Message = message;
        Remediation = remediation;
    }

    public bool Allowed { get; }

    public string RefusalCode { get; }

    public string Message { get; }

    public string Remediation { get; }

    public static CanExecuteResult Permit() => new CanExecuteResult(true, string.Empty, string.Empty, string.Empty);

    public static CanExecuteResult Refuse(string code, string message, string remediation)
    {
        if (string.IsNullOrWhiteSpace(code))
        {
            throw new ArgumentException("A stable refusal code is required.", nameof(code));
        }

        return new CanExecuteResult(false, code, message ?? string.Empty, remediation ?? string.Empty);
    }
}
