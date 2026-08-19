using System;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Core.Collaboration;

public static class CollaborationRefusalCodes
{
    public const string AutoSaveUnknown = "AUTOSAVE_STATE_UNKNOWN";
    public const string CoauthoringUnknown = "COAUTHORING_STATE_UNKNOWN";
    public const string RemoteChangeInProgress = "REMOTE_CHANGE_IN_PROGRESS";
    public const string LegacySharedUnqualified = "LEGACY_SHARED_UNQUALIFIED";
    public const string CoauthoringUnqualified = "COAUTHORING_UNQUALIFIED";
    public const string ReceiptRequired = "COLLABORATION_RECEIPT_REQUIRED";
    public const string RemoteEventsRequired = "REMOTE_EVENTS_REQUIRED";
    public const string StaleWorkbookVersion = "STALE_WORKBOOK_VERSION";
    public const string StalePrecondition = "STALE_PRECONDITION";
    public const string PlanExpired = "PLAN_EXPIRED";
    public const string WorkbookClosed = "WORKBOOK_CLOSED";
    public const string ClockInvalid = "PLAN_CLOCK_INVALID";
}

public sealed class CollaborationPolicyLimits
{
    public CollaborationPolicyLimits(TimeSpan lowImpactLifetime, TimeSpan mediumImpactLifetime, TimeSpan highImpactLifetime)
    {
        LowImpactLifetime = RequirePositive(lowImpactLifetime, nameof(lowImpactLifetime));
        MediumImpactLifetime = RequirePositive(mediumImpactLifetime, nameof(mediumImpactLifetime));
        HighImpactLifetime = RequirePositive(highImpactLifetime, nameof(highImpactLifetime));
    }

    public TimeSpan LowImpactLifetime { get; }

    public TimeSpan MediumImpactLifetime { get; }

    public TimeSpan HighImpactLifetime { get; }

    public TimeSpan For(CommandImpact impact) => impact switch
    {
        CommandImpact.Low => LowImpactLifetime,
        CommandImpact.Medium => MediumImpactLifetime,
        CommandImpact.High => HighImpactLifetime,
        _ => TimeSpan.MaxValue
    };

    private static TimeSpan RequirePositive(TimeSpan value, string parameterName)
    {
        if (value <= TimeSpan.Zero || value > TimeSpan.FromMinutes(5))
        {
            throw new ArgumentOutOfRangeException(parameterName, "A plan lifetime must be positive and no longer than five minutes.");
        }

        return value;
    }
}

public sealed class CollaborationPlanLease
{
    public CollaborationPlanLease(
        string commandId,
        CommandImpact impact,
        WorkbookConcurrencyStamp plannedStamp,
        DateTimeOffset createdAtUtc)
    {
        if (string.IsNullOrWhiteSpace(commandId))
        {
            throw new ArgumentException("A command ID is required.", nameof(commandId));
        }

        if (!Enum.IsDefined(typeof(CommandImpact), impact))
        {
            throw new ArgumentOutOfRangeException(nameof(impact));
        }

        CommandId = commandId;
        Impact = impact;
        PlannedStamp = plannedStamp ?? throw new ArgumentNullException(nameof(plannedStamp));
        CreatedAtUtc = createdAtUtc.ToUniversalTime();
    }

    public string CommandId { get; }

    public CommandImpact Impact { get; }

    public WorkbookConcurrencyStamp PlannedStamp { get; }

    public DateTimeOffset CreatedAtUtc { get; }
}

public sealed class CollaborationDecision
{
    private CollaborationDecision(bool allowed, string? refusalCode, string message)
    {
        Allowed = allowed;
        RefusalCode = refusalCode;
        Message = message;
    }

    public bool Allowed { get; }

    public string? RefusalCode { get; }

    public string Message { get; }

    public static CollaborationDecision Permit() => new CollaborationDecision(true, null, string.Empty);

    public static CollaborationDecision Refuse(string code, string message) =>
        new CollaborationDecision(
            false,
            !string.IsNullOrWhiteSpace(code) ? code : throw new ArgumentException("A refusal code is required.", nameof(code)),
            !string.IsNullOrWhiteSpace(message) ? message : throw new ArgumentException("A refusal message is required.", nameof(message)));
}

public sealed class CollaborationPolicy
{
    private readonly CollaborationPolicyLimits _limits;

    public CollaborationPolicy(CollaborationPolicyLimits limits)
    {
        _limits = limits ?? throw new ArgumentNullException(nameof(limits));
    }

    public CollaborationDecision CanPlan(
        CommandImpact impact,
        WorkbookCollaborationState state,
        bool receiptEligible)
    {
        if (state is null)
        {
            throw new ArgumentNullException(nameof(state));
        }

        if (!Enum.IsDefined(typeof(CommandImpact), impact))
        {
            throw new ArgumentOutOfRangeException(nameof(impact));
        }

        if (impact == CommandImpact.ReadOnly)
        {
            return CollaborationDecision.Permit();
        }

        if (state.AutoSave == AutoSaveState.Unknown)
        {
            return CollaborationDecision.Refuse(
                CollaborationRefusalCodes.AutoSaveUnknown,
                "AutoSave state could not be determined without changing the workbook setting.");
        }

        if (state.Coauthoring == CoauthoringState.Unknown)
        {
            return CollaborationDecision.Refuse(
                CollaborationRefusalCodes.CoauthoringUnknown,
                "Modern coauthoring risk could not be determined from qualified signals.");
        }

        if (state.Coauthoring == CoauthoringState.RemoteChangeInProgress)
        {
            return CollaborationDecision.Refuse(
                CollaborationRefusalCodes.RemoteChangeInProgress,
                "A remote change is currently being merged.");
        }

        if (state.Coauthoring == CoauthoringState.LegacyShared)
        {
            return CollaborationDecision.Refuse(
                CollaborationRefusalCodes.LegacySharedUnqualified,
                "Legacy shared-workbook mutation is not qualified.");
        }

        var collaborativeRisk =
            state.AutoSave == AutoSaveState.On ||
            state.Coauthoring == CoauthoringState.PotentialModern ||
            state.Coauthoring == CoauthoringState.RemoteChangeObserved;

        if (impact == CommandImpact.High && collaborativeRisk)
        {
            return CollaborationDecision.Refuse(
                CollaborationRefusalCodes.CoauthoringUnqualified,
                "High-impact mutation is unavailable while AutoSave or coauthoring risk is present.");
        }

        if (impact == CommandImpact.Medium)
        {
            if (!receiptEligible)
            {
                return CollaborationDecision.Refuse(
                    CollaborationRefusalCodes.ReceiptRequired,
                    "Medium-impact mutation requires an eligible session receipt.");
            }

            if (collaborativeRisk && !state.RemoteChangeEventsSupported)
            {
                return CollaborationDecision.Refuse(
                    CollaborationRefusalCodes.RemoteEventsRequired,
                    "Medium-impact mutation requires qualified remote-change invalidation signals.");
            }
        }

        return CollaborationDecision.Permit();
    }

    public CollaborationDecision ValidateForExecution(
        CollaborationPlanLease lease,
        WorkbookConcurrencyStamp currentStamp,
        DateTimeOffset nowUtc,
        bool receiptEligible)
    {
        if (lease is null)
        {
            throw new ArgumentNullException(nameof(lease));
        }

        if (currentStamp is null)
        {
            throw new ArgumentNullException(nameof(currentStamp));
        }

        if (currentStamp.WorkbookClosed)
        {
            return CollaborationDecision.Refuse(
                CollaborationRefusalCodes.WorkbookClosed,
                "The planned workbook has closed.");
        }

        if (!string.Equals(
                lease.PlannedStamp.WorkbookIdentity,
                currentStamp.WorkbookIdentity,
                StringComparison.Ordinal))
        {
            return CollaborationDecision.Refuse(
                CollaborationRefusalCodes.StaleWorkbookVersion,
                "The active workbook identity no longer matches the plan.");
        }

        if (lease.Impact == CommandImpact.ReadOnly)
        {
            return CollaborationDecision.Permit();
        }

        if (lease.PlannedStamp.Revision != currentStamp.Revision)
        {
            return CollaborationDecision.Refuse(
                CollaborationRefusalCodes.StaleWorkbookVersion,
                "A workbook invalidation event occurred after planning.");
        }

        if (!string.Equals(
                lease.PlannedStamp.PreconditionFingerprint,
                currentStamp.PreconditionFingerprint,
                StringComparison.Ordinal))
        {
            return CollaborationDecision.Refuse(
                CollaborationRefusalCodes.StalePrecondition,
                "A planned property changed after planning.");
        }

        var normalizedNow = nowUtc.ToUniversalTime();
        if (normalizedNow < lease.CreatedAtUtc)
        {
            return CollaborationDecision.Refuse(
                CollaborationRefusalCodes.ClockInvalid,
                "The plan clock moved backwards; refresh the plan.");
        }

        if (normalizedNow - lease.CreatedAtUtc > _limits.For(lease.Impact))
        {
            return CollaborationDecision.Refuse(
                CollaborationRefusalCodes.PlanExpired,
                "The collaboration-sensitive plan expired; refresh it before mutation.");
        }

        return CanPlan(lease.Impact, currentStamp.Collaboration, receiptEligible);
    }
}
