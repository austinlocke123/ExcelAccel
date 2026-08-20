using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Profiles;
using ExcelAccel.Core.Commands;
using ExcelAccel.Application.Undo;

namespace ExcelAccel.Application.Formatting;

public interface IFormattingPort : ISelectionPort
{
    string ReadFormattingProperty(string propertyId);
    void WriteFormattingProperty(string propertyId, string invariantValue);
}

public sealed class ProfileFormattingCommand
{
    public const long MaximumCellCount = 50_000;
    private readonly CommandDescriptor _descriptor;
    private readonly Func<ProfileDefinition, string, string> _resolveValue;
    private readonly Func<string, string, bool> _verifyPostcondition;
    private readonly string? _noChangeMessage;

    public ProfileFormattingCommand(
        CommandDescriptor descriptor,
        Func<ProfileDefinition, string, string> resolveValue,
        Func<string, string, bool>? verifyPostcondition = null,
        string? noChangeMessage = null)
    {
        _noChangeMessage = noChangeMessage;
        _descriptor = descriptor ?? throw new ArgumentNullException(nameof(descriptor));
        _resolveValue = resolveValue ?? throw new ArgumentNullException(nameof(resolveValue));
        _verifyPostcondition = verifyPostcondition ?? ((desired, observed) =>
            string.Equals(desired, observed, StringComparison.OrdinalIgnoreCase));
        if ((_descriptor.Impact != CommandImpact.Low && _descriptor.Impact != CommandImpact.Medium) || _descriptor.ChangedProperties.Count != 1)
        {
            throw new ArgumentException("Profile formatting commands require one low/medium-impact changed property.", nameof(descriptor));
        }
    }

    public CommandPlan Plan(ProfileDefinition profile, IFormattingPort port)
    {
        if (profile is null)
        {
            throw new ArgumentNullException(nameof(profile));
        }

        if (port is null)
        {
            throw new ArgumentNullException(nameof(port));
        }

        var snapshot = port.CaptureSelection();
        RequireSafeSelection(snapshot);
        var propertyId = _descriptor.ChangedProperties[0];
        var current = port.ReadFormattingProperty(propertyId);
        var desired = _resolveValue(profile, current);
        if (string.IsNullOrWhiteSpace(desired) && _noChangeMessage is not null)
        {
            // The resolver derives its value from the cell's current format
            // rather than from the profile, so an empty result means there was
            // nothing to change, not that the profile is broken.
            throw new CommandRefusedException(
                RefusalCodes.NoChangeRequired,
                _noChangeMessage,
                "No change was needed, so nothing was written.");
        }

        if (string.IsNullOrWhiteSpace(desired))
        {
            throw new CommandRefusedException(
                RefusalCodes.CommandUnavailable,
                $"The active profile does not define a usable value for '{propertyId}'.",
                "Repair or reset the active profile and try again.");
        }

        return new CommandPlan(
            _descriptor.Id,
            _descriptor.Impact,
            snapshot.Context,
            _descriptor.ChangedProperties,
            snapshot.CellCount,
            $"Change only {_descriptor.ChangedProperties[0]} on {snapshot.CellCount:N0} cell(s).",
            PreconditionFingerprint.Create(current),
            _descriptor.ContractVersion,
            requiresPreview: _descriptor.PreviewPolicy == PreviewPolicy.Mandatory,
            arguments: new[]
            {
                new KeyValuePair<string, string>("property_id", propertyId),
                new KeyValuePair<string, string>("value", desired),
            });
    }

    public CommandResult Execute(CommandPlan plan, IFormattingPort port, string? confirmedPlanHash = null, IPropertyReceiptSink? receiptSink = null)
    {
        if (plan is null)
        {
            throw new ArgumentNullException(nameof(plan));
        }

        if (port is null)
        {
            throw new ArgumentNullException(nameof(port));
        }

        var authorization = CommandExecutionGate.Authorize(_descriptor, plan, confirmedPlanHash);
        if (!authorization.Allowed)
        {
            return CommandResult.Refused(plan, authorization.Message, authorization.RefusalCode);
        }

        var snapshot = port.CaptureSelection();
        if (!plan.Context.Equals(snapshot.Context) || snapshot.CellCount != plan.AffectedCellCount)
        {
            return CommandResult.Refused(plan, "The selection changed after planning.", RefusalCodes.StaleContext);
        }

        try
        {
            RequireSafeSelection(snapshot);
        }
        catch (CommandRefusedException exception)
        {
            return CommandResult.Refused(plan, exception.Message, exception.RefusalCode);
        }

        var propertyId = plan.Arguments["property_id"];
        var desired = plan.Arguments["value"];
        var current = port.ReadFormattingProperty(propertyId);
        if (!string.Equals(plan.PreconditionFingerprint, PreconditionFingerprint.Create(current), StringComparison.Ordinal))
        {
            return CommandResult.Refused(plan, "The formatting property changed after planning.", RefusalCodes.StaleContext);
        }

        port.WriteFormattingProperty(propertyId, desired);
        var observed = port.ReadFormattingProperty(propertyId);
        if (!_verifyPostcondition(desired, observed))
        {
            return CommandResult.Failed(plan.CommandId, "Excel did not report the planned formatting postcondition.", "POSTCONDITION_MISMATCH");
        }

        var receiptId = string.Empty;
        if (receiptSink is not null)
        {
            receiptId = Guid.NewGuid().ToString("N");
            var now = DateTimeOffset.UtcNow;
            receiptSink.Add(new PropertyReceipt(receiptId, plan.CommandId, plan.ContractVersion, plan.Context,
                propertyId, current, desired, plan.PlanHash, now, now.AddHours(8)));
        }
        return CommandResult.Success(plan, $"Applied {_descriptor.DisplayName} to {plan.AffectedCellCount:N0} cell(s).", receiptId);
    }

    public static string Next(IReadOnlyList<string> cycle, string current)
    {
        if (cycle is null || cycle.Count == 0)
        {
            return string.Empty;
        }

        var currentIndex = cycle
            .Select((value, index) => new { value, index })
            .FirstOrDefault(item => string.Equals(item.value, current, StringComparison.OrdinalIgnoreCase))?.index ?? -1;
        return cycle[(currentIndex + 1) % cycle.Count];
    }

    private static void RequireSafeSelection(SelectionSnapshot snapshot)
    {
        if (snapshot.Safety.AreaCount != 1 || snapshot.Safety.HasMergedCells)
        {
            throw new CommandRefusedException(RefusalCodes.SelectionUnsupported, "Formatting requires one unmerged rectangular selection.", "Select one unmerged range.");
        }

        if (snapshot.Safety.WorksheetProtected || snapshot.Safety.WorkbookReadOnly)
        {
            throw new CommandRefusedException(RefusalCodes.ProtectedTarget, "The target is protected or read-only.", "Use an editable unprotected target.");
        }

        if (!snapshot.Safety.DynamicArraySpillCheckSupported || snapshot.Safety.HasLegacyArray || snapshot.Safety.HasDynamicArraySpill)
        {
            throw new CommandRefusedException(RefusalCodes.ArrayOrSpillUnsafe, "The target intersects an unqualified array or spill state.", "Select cells outside array/spill ranges.");
        }

        if (snapshot.CellCount > MaximumCellCount)
        {
            throw new CommandRefusedException(RefusalCodes.ResourceLimit, "The selection exceeds the immediate formatting limit.", "Select a smaller range.");
        }
    }
}
