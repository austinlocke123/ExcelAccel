using System;

namespace ExcelAccel.Core.Collaboration;

public sealed class WorkbookConcurrencyStamp
{
    public WorkbookConcurrencyStamp(
        string workbookIdentity,
        long revision,
        string preconditionFingerprint,
        WorkbookCollaborationState collaboration,
        bool workbookClosed)
    {
        if (string.IsNullOrWhiteSpace(workbookIdentity))
        {
            throw new ArgumentException("A workbook identity is required.", nameof(workbookIdentity));
        }

        if (revision < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(revision));
        }

        WorkbookIdentity = workbookIdentity;
        Revision = revision;
        PreconditionFingerprint = preconditionFingerprint ?? throw new ArgumentNullException(nameof(preconditionFingerprint));
        Collaboration = collaboration ?? throw new ArgumentNullException(nameof(collaboration));
        WorkbookClosed = workbookClosed;
    }

    public string WorkbookIdentity { get; }

    public long Revision { get; }

    public string PreconditionFingerprint { get; }

    public WorkbookCollaborationState Collaboration { get; }

    public bool WorkbookClosed { get; }
}
