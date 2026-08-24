using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Profiles;
using ExcelAccel.Application.Undo;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.AutoColor;

/// <summary>
/// Reads the cells a recolour will consider and writes their font colours.
/// </summary>
/// <remarks>
/// Extends <see cref="IPropertyReceiptPort"/> because undo restores the whole
/// block through the same port, under
/// <see cref="FontColorBlock.ReceiptPropertyId"/>.
/// </remarks>
public interface IAutoColorPort : IReferenceValuePropertyReceiptPort
{
    SelectionSnapshot CaptureSelection();

    /// <summary>Every cell in scope, with its kind, formula, and current colour.</summary>
    IReadOnlyList<AutoColorCellSnapshot> CaptureCells(AutoColorScope scope);

    /// <summary>Writes font colour only, on the given cells only.</summary>
    void WriteFontColors(IEnumerable<AutoColorChange> changes);
}

public sealed class AutoColorExecution
{
    public AutoColorExecution(CommandPlan commandPlan, AutoColorPlan plan)
    {
        CommandPlan = commandPlan;
        Plan = plan;
    }

    public CommandPlan CommandPlan { get; }
    public AutoColorPlan Plan { get; }
}

/// <summary>
/// Plans and executes a recolour. Writes font colour and nothing else, verifies
/// what it wrote, restores the exact before-state on failure, and records the
/// whole change as one undo receipt.
/// </summary>
public sealed class AutoColorCommand
{
    private readonly CommandDescriptor _descriptor;
    private readonly AutoColorScope _scope;

    public AutoColorCommand(CommandDescriptor descriptor, AutoColorScope scope)
    {
        _descriptor = descriptor ?? throw new ArgumentNullException(nameof(descriptor));
        _scope = scope;
    }

    public AutoColorExecution Plan(ProfileDefinition profile, IAutoColorPort port)
    {
        if (profile is null) throw new ArgumentNullException(nameof(profile));
        if (port is null) throw new ArgumentNullException(nameof(port));

        var gate = AutoColorPlanner.ExecutionGate(_scope);
        if (!gate.Allowed)
        {
            throw new CommandRefusedException(gate.RefusalCode!, gate.Message!, gate.Remediation!);
        }

        var selection = port.CaptureSelection();
        var cells = port.CaptureCells(_scope);
        var plan = AutoColorPlanner.Plan(profile, selection, cells, _scope);

        // A recolour has to fit in one undo value, or it could not be reversed.
        // Refusing here is better than writing something the user cannot undo.
        if (plan.Changes.Count > FontColorBlock.MaximumCells)
        {
            throw new CommandRefusedException(
                RefusalCodes.ResourceLimit,
                string.Format(
                    CultureInfo.InvariantCulture,
                    "This scope would recolour {0:N0} cells, more than the {1:N0} a single undo can hold.",
                    plan.Changes.Count,
                    FontColorBlock.MaximumCells),
                "Select a smaller range and run it again.");
        }

        var commandPlan = new CommandPlan(
            _descriptor.Id,
            _descriptor.Impact,
            plan.Context,
            new[] { FontColorBlock.ReceiptPropertyId },
            plan.Changes.Count,
            Summarize(plan),
            plan.Fingerprint,
            _descriptor.ContractVersion,
            requiresPreview: plan.RequiresPreview,
            arguments: new[]
            {
                new KeyValuePair<string, string>("scope", _scope.ToString().ToLowerInvariant()),
                new KeyValuePair<string, string>("unsupported", plan.UnsupportedCount.ToString(CultureInfo.InvariantCulture)),
            });

        return new AutoColorExecution(commandPlan, plan);
    }

    public CommandResult Execute(
        AutoColorExecution execution,
        ProfileDefinition profile,
        IAutoColorPort port,
        string? confirmedPlanHash = null,
        IPropertyBatchReceiptSink? receiptSink = null)
    {
        if (execution is null) throw new ArgumentNullException(nameof(execution));
        if (profile is null) throw new ArgumentNullException(nameof(profile));
        if (port is null) throw new ArgumentNullException(nameof(port));

        var authorization = CommandExecutionGate.Authorize(_descriptor, execution.CommandPlan, confirmedPlanHash);
        if (!authorization.Allowed)
        {
            return CommandResult.Refused(execution.CommandPlan, authorization.Message, authorization.RefusalCode);
        }

        if (receiptSink is null)
        {
            return CommandResult.Refused(
                execution.CommandPlan,
                "Recolouring requires an available bounded undo store.",
                RefusalCodes.CommandUnavailable);
        }

        if (execution.Plan.Changes.Count == 0)
        {
            return CommandResult.Success(_descriptor.Id, "Every cell already carries its classified colour.");
        }

        // Re-read and compare the complete fingerprint: it covers every cell's
        // kind, formula, and current colour, so any edit since planning shows up.
        var current = port.CaptureCells(_scope);
        var currentPlan = AutoColorPlanner.Plan(profile, port.CaptureSelection(), current, _scope);
        if (!string.Equals(currentPlan.Fingerprint, execution.Plan.Fingerprint, StringComparison.Ordinal))
        {
            return CommandResult.Refused(
                execution.CommandPlan,
                "The target changed after planning.",
                RefusalCodes.StaleContext);
        }

        var before = FontColorBlock.Serialize(
            execution.Plan.Changes.Select(change => new KeyValuePair<string, string>(change.Address, change.BeforeColor)));
        var after = FontColorBlock.Serialize(
            execution.Plan.Changes.Select(change => new KeyValuePair<string, string>(change.Address, change.AfterColor)));

        try
        {
            port.WriteFontColors(execution.Plan.Changes);
            var observed = Observed(port, execution);
            if (!string.Equals(observed, after, StringComparison.Ordinal))
            {
                throw new InvalidOperationException("AutoColor postcondition mismatch.");
            }
        }
        catch (Exception exception)
        {
            return Rollback(execution, port, before, exception);
        }

        var receiptId = Guid.NewGuid().ToString("N");
        var now = DateTimeOffset.UtcNow;
        try
        {
            receiptSink.Add(new PropertyBatchReceipt(
                receiptId,
                _descriptor.Id,
                _descriptor.ContractVersion,
                execution.CommandPlan.Context,
                new[] { new PropertyChange(FontColorBlock.ReceiptPropertyId, before, after) },
                execution.CommandPlan.PlanHash,
                now,
                now.AddHours(8)));
        }
        catch (Exception exception)
        {
            return Rollback(execution, port, before, exception, receiptFailure: true);
        }

        return CommandResult.Success(
            _descriptor.Id,
            string.Format(
                CultureInfo.InvariantCulture,
                "Recoloured {0:N0} cell(s); {1:N0} left unchanged.",
                execution.Plan.Changes.Count,
                execution.Plan.UnsupportedCount),
            execution.Plan.Changes.Count,
            receiptId);
    }

    /// <summary>
    /// Restores the exact before-state. A rollback that cannot be verified is
    /// reported as partial rather than as a failure that implies nothing changed.
    /// </summary>
    private CommandResult Rollback(
        AutoColorExecution execution,
        IAutoColorPort port,
        string before,
        Exception cause,
        bool receiptFailure = false)
    {
        var restored = false;
        try
        {
            port.WriteFontColors(execution.Plan.Changes
                .Select(change => new AutoColorChange(change.Address, change.Category, change.AfterColor, change.BeforeColor)));
            restored = string.Equals(Observed(port, execution), before, StringComparison.Ordinal);
        }
        catch (Exception)
        {
            restored = false;
        }

        var reason = receiptFailure ? "The undo receipt could not be stored" : "Recolouring failed";
        return restored
            ? CommandResult.Failed(
                _descriptor.Id,
                $"{reason} ({cause.GetType().Name}); every cell was restored to its exact prior colour.",
                receiptFailure ? "RECEIPT_STORE_ROLLED_BACK" : "AUTO_COLOR_ROLLED_BACK")
            : CommandResult.Partial(
                _descriptor.Id,
                $"{reason} ({cause.GetType().Name}) and rollback could not be verified; inspect the target's font colours.",
                execution.Plan.Changes.Count,
                execution.Plan.UnsupportedCount,
                receiptFailure ? "RECEIPT_STORE_ROLLBACK_INCOMPLETE" : "AUTO_COLOR_ROLLBACK_INCOMPLETE");
    }

    /// <summary>Reads back only the cells the plan touched, in block form.</summary>
    private string Observed(IAutoColorPort port, AutoColorExecution execution)
    {
        var touched = new HashSet<string>(
            execution.Plan.Changes.Select(change => change.Address), StringComparer.OrdinalIgnoreCase);
        return FontColorBlock.Serialize(port
            .CaptureCells(_scope)
            .Where(cell => touched.Contains(cell.Address))
            .Select(cell => new KeyValuePair<string, string>(cell.Address, cell.FontColor)));
    }

    private string Summarize(AutoColorPlan plan) =>
        string.Format(
            CultureInfo.InvariantCulture,
            "Change only font colour on {0:N0} of {1:N0} cell(s) in the {2}; {3:N0} unsupported cell(s) are left alone.",
            plan.Changes.Count,
            plan.Counts.Values.Sum(),
            _scope.ToString().ToLowerInvariant(),
            plan.UnsupportedCount);
}
