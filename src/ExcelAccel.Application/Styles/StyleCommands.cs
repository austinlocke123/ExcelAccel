using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Formatting;
using ExcelAccel.Application.Undo;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.Styles;

public interface IStylePort : IFormattingPort, IPropertyReceiptPort { }

public sealed class StyleCaptureCommand
{
    public StyleRecipe Capture(string styleId, string displayName, IEnumerable<string> requestedProperties, IStylePort port)
    {
        if (port is null) throw new ArgumentNullException(nameof(port));
        var snapshot = port.CaptureSelection();
        if (snapshot.CellCount != 1 || snapshot.Safety.AreaCount != 1 || snapshot.Safety.HasMergedCells)
            throw new CommandRefusedException(RefusalCodes.SelectionUnsupported, "Style capture requires exactly one unmerged source cell.", "Select one unmerged cell and retry.");
        var requested = (requestedProperties ?? throw new ArgumentNullException(nameof(requestedProperties)))
            .Select(value => value?.Trim().ToLowerInvariant() ?? string.Empty).Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal).ToArray();
        if (requested.Length == 0) throw new CommandRefusedException(RefusalCodes.SelectionUnsupported, "Select at least one supported formatting property.", "Choose one or more style properties.");
        var unsupported = requested.Where(value => !StylePropertyCatalog.Supported.Contains(value, StringComparer.Ordinal)).ToArray();
        if (unsupported.Length > 0)
            throw new CommandRefusedException(RefusalCodes.CommandUnavailable, $"Unsupported capture properties: {string.Join(", ", unsupported)}.", "Remove unsupported properties and retry.");
        var properties = requested.Select(value => new KeyValuePair<string, string>(value, port.ReadFormattingProperty(value))).ToArray();
        return new StyleRecipe(styleId, StyleRecipe.CurrentVersion, displayName, StyleOrigin.Local, UnsupportedStylePropertyPolicy.Refuse, properties);
    }
}

public sealed class StylePropertyChange
{
    public StylePropertyChange(string propertyId, string beforeValue, string afterValue)
    { PropertyId = propertyId; BeforeValue = beforeValue; AfterValue = afterValue; }
    public string PropertyId { get; }
    public string BeforeValue { get; }
    public string AfterValue { get; }
}

public sealed class StyleApplyPlan
{
    public StyleApplyPlan(CommandPlan commandPlan, StyleRecipe recipe, IEnumerable<StylePropertyChange> changes,
        IEnumerable<KeyValuePair<string, string>> skipped)
    { CommandPlan = commandPlan; Recipe = recipe; Changes = changes.ToArray(); Skipped = skipped.ToArray(); }
    public CommandPlan CommandPlan { get; }
    public StyleRecipe Recipe { get; }
    public IReadOnlyList<StylePropertyChange> Changes { get; }
    public IReadOnlyList<KeyValuePair<string, string>> Skipped { get; }
}

public sealed class StyleApplyCommand
{
    public const long MaximumCellCount = 50_000;
    private readonly CommandDescriptor _descriptor;
    public StyleApplyCommand(CommandDescriptor descriptor) => _descriptor = descriptor ?? throw new ArgumentNullException(nameof(descriptor));

    public StyleApplyPlan Plan(StyleRecipe recipe, IStylePort port, long immediatePreviewCellLimit)
    {
        if (recipe is null) throw new ArgumentNullException(nameof(recipe));
        if (port is null) throw new ArgumentNullException(nameof(port));
        var snapshot = port.CaptureSelection();
        RequireSafe(snapshot);
        var changes = new List<StylePropertyChange>();
        var skipped = new List<KeyValuePair<string, string>>();
        foreach (var property in recipe.Properties)
        {
            string current;
            try { current = port.ReadFormattingProperty(property.Key); }
            catch (CommandRefusedException exception)
            {
                if (recipe.UnsupportedPropertyPolicy == UnsupportedStylePropertyPolicy.Skip)
                { skipped.Add(new KeyValuePair<string, string>(property.Key, exception.Message)); continue; }
                throw;
            }
            if (string.Equals(current, "(mixed)", StringComparison.OrdinalIgnoreCase))
            {
                if (recipe.UnsupportedPropertyPolicy == UnsupportedStylePropertyPolicy.Skip)
                { skipped.Add(new KeyValuePair<string, string>(property.Key, "The selected cells have mixed before-state.")); continue; }
                throw new CommandRefusedException(RefusalCodes.SelectionUnsupported,
                    $"Property '{property.Key}' has mixed before-state and cannot be rolled back exactly.", "Use a uniformly formatted selection or a skip-safe recipe.");
            }
            if (!string.Equals(current, property.Value, StringComparison.OrdinalIgnoreCase))
                changes.Add(new StylePropertyChange(property.Key, current, property.Value));
        }
        var plannedProperties = changes.Count > 0 ? changes.Select(value => value.PropertyId) : recipe.Properties.Keys;
        var fingerprint = PreconditionFingerprint.Create(string.Join("\n", changes.Select(value => value.PropertyId + "=" + value.BeforeValue)));
        var commandPlan = new CommandPlan(_descriptor.Id, _descriptor.Impact, snapshot.Context, plannedProperties,
            snapshot.CellCount, $"Apply style '{recipe.DisplayName}' to {snapshot.CellCount:N0} cell(s): {changes.Count} properties changed, {skipped.Count} skipped.",
            fingerprint, _descriptor.ContractVersion, snapshot.CellCount > immediatePreviewCellLimit,
            new[] { new KeyValuePair<string, string>("style_id", recipe.StyleId), new KeyValuePair<string, string>("style_version", recipe.Version.ToString()) });
        return new StyleApplyPlan(commandPlan, recipe, changes, skipped);
    }

    public CommandResult Execute(StyleApplyPlan plan, IStylePort port, string? confirmedPlanHash = null, IPropertyBatchReceiptSink? receiptSink = null)
    {
        if (plan is null) throw new ArgumentNullException(nameof(plan));
        if (port is null) throw new ArgumentNullException(nameof(port));
        var authorization = CommandExecutionGate.Authorize(_descriptor, plan.CommandPlan, confirmedPlanHash);
        if (!authorization.Allowed) return CommandResult.Refused(plan.CommandPlan, authorization.Message, authorization.RefusalCode);
        var snapshot = port.CaptureSelection();
        if (!plan.CommandPlan.Context.Equals(snapshot.Context) || snapshot.CellCount != plan.CommandPlan.AffectedCellCount)
            return CommandResult.Refused(plan.CommandPlan, "The target selection changed after preview.", RefusalCodes.StaleContext);
        foreach (var change in plan.Changes)
        {
            var current = port.ReadFormattingProperty(change.PropertyId);
            if (!string.Equals(current, change.BeforeValue, StringComparison.OrdinalIgnoreCase))
                return CommandResult.Refused(plan.CommandPlan, $"Style property '{change.PropertyId}' changed after preview.", RefusalCodes.StaleContext);
        }
        if (plan.Changes.Count == 0)
            return CommandResult.Success(_descriptor.Id, $"Style '{plan.Recipe.DisplayName}' already matches the target; {plan.Skipped.Count} properties skipped.", 0);

        var attempted = new List<StylePropertyChange>();
        foreach (var change in plan.Changes)
        {
            attempted.Add(change);
            try
            {
                port.WriteFormattingProperty(change.PropertyId, change.AfterValue);
                var observed = port.ReadFormattingProperty(change.PropertyId);
                if (!string.Equals(observed, change.AfterValue, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Postcondition mismatch.");
            }
            catch (Exception)
            {
                var remaining = RollBack(attempted, port);
                if (remaining.Count == 0)
                    return CommandResult.Failed(_descriptor.Id, "Style apply failed; every attempted property was restored to its exact before-state.", "STYLE_APPLY_ROLLED_BACK");
                return CommandResult.Partial(_descriptor.Id,
                    $"Style apply and rollback were incomplete. Remaining changed properties: {string.Join(", ", remaining)}.",
                    remaining.Count, plan.Skipped.Count, "STYLE_ROLLBACK_INCOMPLETE");
            }
        }

        var receiptId = string.Empty;
        if (receiptSink is not null)
        {
            receiptId = Guid.NewGuid().ToString("N");
            var now = DateTimeOffset.UtcNow;
            receiptSink.Add(new PropertyBatchReceipt(receiptId, _descriptor.Id, _descriptor.ContractVersion, plan.CommandPlan.Context,
                plan.Changes.Select(value => new PropertyChange(value.PropertyId, value.BeforeValue, value.AfterValue)),
                plan.CommandPlan.PlanHash, now, now.AddHours(8)));
        }
        return CommandResult.Success(plan.CommandPlan,
            $"Applied style '{plan.Recipe.DisplayName}' with {plan.Changes.Count} verified properties; {plan.Skipped.Count} skipped.", receiptId);
    }

    private static IReadOnlyList<string> RollBack(IEnumerable<StylePropertyChange> attempted, IStylePort port)
    {
        var changes = attempted.Reverse().ToArray();
        foreach (var change in changes)
        {
            try { port.WriteFormattingProperty(change.PropertyId, change.BeforeValue); } catch { }
        }
        var remaining = new List<string>();
        foreach (var change in changes)
        {
            try
            {
                if (!string.Equals(port.ReadFormattingProperty(change.PropertyId), change.BeforeValue, StringComparison.OrdinalIgnoreCase)) remaining.Add(change.PropertyId);
            }
            catch { remaining.Add(change.PropertyId); }
        }
        return remaining.OrderBy(value => value, StringComparer.Ordinal).ToArray();
    }

    private static void RequireSafe(SelectionSnapshot snapshot)
    {
        if (snapshot.Safety.AreaCount != 1 || snapshot.Safety.HasMergedCells)
            throw new CommandRefusedException(RefusalCodes.SelectionUnsupported, "Style apply requires one unmerged rectangular selection.", "Select one unmerged range.");
        if (snapshot.Safety.WorksheetProtected || snapshot.Safety.WorkbookReadOnly)
            throw new CommandRefusedException(RefusalCodes.ProtectedTarget, "The target is protected or read-only.", "Use an editable unprotected target.");
        if (!snapshot.Safety.DynamicArraySpillCheckSupported || snapshot.Safety.HasLegacyArray || snapshot.Safety.HasDynamicArraySpill)
            throw new CommandRefusedException(RefusalCodes.ArrayOrSpillUnsafe, "The target intersects an unqualified array or spill state.", "Select cells outside array/spill ranges.");
        if (snapshot.CellCount > MaximumCellCount)
            throw new CommandRefusedException(RefusalCodes.ResourceLimit, $"The selection exceeds the {MaximumCellCount:N0}-cell style limit.", "Select a smaller range.");
    }
}
