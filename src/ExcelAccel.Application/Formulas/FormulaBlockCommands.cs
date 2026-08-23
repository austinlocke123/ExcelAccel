using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Formatting;
using ExcelAccel.Application.Undo;
using ExcelAccel.Core.Commands;
using ExcelAccel.Core.Formulas;

namespace ExcelAccel.Application.Formulas;

public interface IFormulaBlockPort : ISelectionPort, IPropertyReceiptPort
{
    FormulaBlockSnapshot CaptureFormulaBlock();
    FormulaBlockSnapshot CaptureFormulaBlock(SelectionContext target);
    void WriteFormulaBlock(FormulaCellBlock contents);
    void WriteFormulaBlock(SelectionContext target, FormulaCellBlock contents);

    /// <summary>
    /// Applies one number format across the whole target. Capturing the prior
    /// formats and restoring them needs no new member: <see cref="IPropertyReceiptPort"/>
    /// already reads and writes the whole format block under
    /// <c>cell_format_block_v1</c>, which is what undo uses.
    /// </summary>
    void ApplyNumberFormat(SelectionContext target, string numberFormat);
}

public sealed class FormulaBlockPlan
{
    public FormulaBlockPlan(CommandPlan commandPlan, FormulaBlockSnapshot before, FormulaCellBlock after,
        int changedCount, int skippedCount, IEnumerable<string> samples, FormulaBlockSnapshot? externalSource = null,
        string? numberFormat = null)
    {
        NumberFormat = string.IsNullOrWhiteSpace(numberFormat) ? null : numberFormat;
        CommandPlan = commandPlan ?? throw new ArgumentNullException(nameof(commandPlan));
        Before = before ?? throw new ArgumentNullException(nameof(before));
        After = after ?? throw new ArgumentNullException(nameof(after));
        if (changedCount < 0) throw new ArgumentOutOfRangeException(nameof(changedCount));
        if (skippedCount < 0) throw new ArgumentOutOfRangeException(nameof(skippedCount));
        if (before.Contents.RowCount != after.RowCount || before.Contents.ColumnCount != after.ColumnCount)
            throw new ArgumentException("Formula block plan dimensions must remain stable.", nameof(after));
        ChangedCount = changedCount;
        SkippedCount = skippedCount;
        Samples = Array.AsReadOnly((samples ?? throw new ArgumentNullException(nameof(samples))).Take(10).ToArray());
        ExternalSource = externalSource;
    }
    public CommandPlan CommandPlan { get; }

    /// <summary>An optional number format applied across the whole target.</summary>
    public string? NumberFormat { get; }
    public FormulaBlockSnapshot Before { get; }
    public FormulaCellBlock After { get; }
    public int ChangedCount { get; }
    public int SkippedCount { get; }
    public IReadOnlyList<string> Samples { get; }
    public FormulaBlockSnapshot? ExternalSource { get; }
    public bool RequiresExternalSourceRevalidation => ExternalSource is not null;
}

public enum FormulaCopyDirection
{
    Down,
    Right,
}

public sealed class FormulaBlockCommand
{
    public const string ReceiptPropertyId = "cell_contents_v1";
    public const int DefaultImmediatePreviewLimit = 100;
    private readonly CommandDescriptor _descriptor;
    private readonly A1FormulaTransformer _references;
    private readonly FormulaWrapperTransformer _wrappers;

    public FormulaBlockCommand(
        CommandDescriptor descriptor,
        A1FormulaTransformer? references = null,
        FormulaWrapperTransformer? wrappers = null)
    {
        _descriptor = descriptor ?? throw new ArgumentNullException(nameof(descriptor));
        _references = references ?? new A1FormulaTransformer();
        _wrappers = wrappers ?? new FormulaWrapperTransformer();
    }

    public FormulaBlockPlan PlanCopy(FormulaBlockSnapshot snapshot, FormulaCopyDirection direction, int immediatePreviewLimit = DefaultImmediatePreviewLimit)
    {
        RequireSafe(snapshot);
        if ((direction == FormulaCopyDirection.Down && snapshot.Contents.RowCount < 2) ||
            (direction == FormulaCopyDirection.Right && snapshot.Contents.ColumnCount < 2))
            throw Refuse(RefusalCodes.SelectionUnsupported, "Copy requires a source edge plus at least one destination row or column.");

        var changed = 0;
        var nonblankOverwrite = 0;
        var samples = new List<string>();
        var after = snapshot.Contents.Map((row, column, current) =>
        {
            var destination = direction == FormulaCopyDirection.Down ? row > 0 : column > 0;
            if (!destination) return current;
            var source = direction == FormulaCopyDirection.Down ? snapshot.Contents[0, column] : snapshot.Contents[row, 0];
            if (!source.IsFormula)
                throw Refuse(RefusalCodes.SelectionUnsupported, "Every source-edge cell must contain a qualified formula.");
            var transformed = _references.Translate(source.InvariantValue,
                direction == FormulaCopyDirection.Down ? row : 0,
                direction == FormulaCopyDirection.Right ? column : 0);
            if (!transformed.IsSuccess) throw Refuse(transformed.RefusalCode!, transformed.Message!);
            var proposed = FormulaCellValue.Formula(transformed.Formula!);
            if (current.Equals(proposed)) return current;
            changed++;
            if (!current.IsBlank) nonblankOverwrite++;
            AddSample(samples, row, column, current, proposed);
            return proposed;
        });
        return BuildPlan(snapshot, after, changed, 0, samples,
            nonblankOverwrite > 0 || changed > immediatePreviewLimit,
            $"Translate formulas {direction.ToString().ToLowerInvariant()} into {changed:N0} destination cell(s).",
            new[] { "formula" }, new[] { new KeyValuePair<string, string>("direction", direction.ToString().ToLowerInvariant()) });
    }

    public FormulaBlockPlan PlanIfError(FormulaBlockSnapshot snapshot, string fallbackExpression, int immediatePreviewLimit = DefaultImmediatePreviewLimit) =>
        PlanFormulaMap(snapshot, "IFERROR toggle", formula => _wrappers.ToggleIfError(formula, fallbackExpression),
            immediatePreviewLimit, new[] { "formula" }, new[] { new KeyValuePair<string, string>("fallback", fallbackExpression ?? string.Empty) });

    public FormulaBlockPlan PlanReverseSign(FormulaBlockSnapshot snapshot, bool includeNumericConstants, int immediatePreviewLimit = DefaultImmediatePreviewLimit)
    {
        RequireSafe(snapshot);
        var changed = 0;
        var skipped = 0;
        var samples = new List<string>();
        var changedProperties = new HashSet<string>(StringComparer.Ordinal);
        var after = snapshot.Contents.Map((row, column, current) =>
        {
            FormulaCellValue proposed;
            if (current.IsFormula)
            {
                var transformed = _wrappers.ReverseSign(current.InvariantValue);
                if (!transformed.IsSuccess) throw Refuse(transformed.RefusalCode!, transformed.Message!);
                proposed = FormulaCellValue.Formula(transformed.Formula!);
                changedProperties.Add("formula");
            }
            else if (includeNumericConstants && current.Kind == FormulaCellKind.Number)
            {
                proposed = FormulaCellValue.Number(-current.AsNumber());
                changedProperties.Add("value");
            }
            else { skipped++; return current; }
            if (current.Equals(proposed)) return current;
            changed++;
            AddSample(samples, row, column, current, proposed);
            return proposed;
        });
        if (changedProperties.Count == 0) changedProperties.Add("formula");
        return BuildPlan(snapshot, after, changed, skipped, samples, changed > immediatePreviewLimit,
            $"Reverse sign for {changed:N0} cell(s); skip {skipped:N0}.", changedProperties,
            new[] { new KeyValuePair<string, string>("include_numeric_constants", includeNumericConstants ? "true" : "false") });
    }

    public FormulaBlockPlan PlanScale(FormulaBlockSnapshot snapshot, long scale, bool divide, bool includeNumericConstants,
        int immediatePreviewLimit = DefaultImmediatePreviewLimit, string? numberFormat = null)
    {
        RequireSafe(snapshot);
        var changed = 0;
        var skipped = 0;
        var samples = new List<string>();
        var changedProperties = new HashSet<string>(StringComparer.Ordinal);
        var after = snapshot.Contents.Map((row, column, current) =>
        {
            FormulaCellValue proposed;
            if (current.IsFormula)
            {
                var transformed = _wrappers.Scale(current.InvariantValue, scale, divide);
                if (!transformed.IsSuccess) throw Refuse(transformed.RefusalCode!, transformed.Message!);
                proposed = FormulaCellValue.Formula(transformed.Formula!);
                changedProperties.Add("formula");
            }
            else if (includeNumericConstants && current.Kind == FormulaCellKind.Number)
            {
                var output = divide ? current.AsNumber() / scale : current.AsNumber() * scale;
                if (double.IsNaN(output) || double.IsInfinity(output))
                    throw Refuse(FormulaTransformRefusalCodes.InvalidTransformArgument, "A scaled numeric constant would overflow.");
                proposed = FormulaCellValue.Number(output);
                changedProperties.Add("value");
            }
            else { skipped++; return current; }
            if (current.Equals(proposed)) return current;
            changed++;
            AddSample(samples, row, column, current, proposed);
            return proposed;
        });
        if (changedProperties.Count == 0) changedProperties.Add("formula");
        if (!string.IsNullOrWhiteSpace(numberFormat)) changedProperties.Add(FormatPasteCommand.ReceiptPropertyId);
        var operation = divide ? "divide" : "multiply";
        return BuildPlan(snapshot, after, changed, skipped, samples, changed > immediatePreviewLimit,
            $"{operation} {changed:N0} cell(s) by {scale.ToString("N0", CultureInfo.InvariantCulture)}; skip {skipped:N0}.",
            changedProperties, new[]
            {
                new KeyValuePair<string, string>("operator", divide ? "/" : "*"),
                new KeyValuePair<string, string>("scale", scale.ToString(CultureInfo.InvariantCulture)),
                new KeyValuePair<string, string>("include_numeric_constants", includeNumericConstants ? "true" : "false"),
            }, numberFormat);
    }

    public CommandResult Execute(FormulaBlockPlan plan, IFormulaBlockPort port, string? confirmedPlanHash, IPropertyReceiptSink? receiptSink)
    {
        if (plan is null) throw new ArgumentNullException(nameof(plan));
        if (port is null) throw new ArgumentNullException(nameof(port));
        var authorization = CommandExecutionGate.Authorize(_descriptor, plan.CommandPlan, confirmedPlanHash);
        if (!authorization.Allowed) return CommandResult.Refused(plan.CommandPlan, authorization.Message, authorization.RefusalCode);
        if (receiptSink is null) return CommandResult.Refused(plan.CommandPlan, "Formula mutation requires an available bounded undo receipt store.", RefusalCodes.CommandUnavailable);
        if (plan.ExternalSource is not null)
        {
            var source = port.CaptureFormulaBlock(plan.ExternalSource.Selection.Context);
            if (!plan.ExternalSource.SourceContentEquals(source))
                return CommandResult.Refused(plan.CommandPlan, "The external source range changed after planning.", RefusalCodes.StaleContext);
        }
        var current = port.CaptureFormulaBlock(plan.Before.Selection.Context);
        try { RequireSafe(current); }
        catch (CommandRefusedException exception) { return CommandResult.Refused(plan.CommandPlan, exception.Message, exception.RefusalCode); }
        if (!plan.Before.Selection.Context.Equals(current.Selection.Context) ||
            plan.Before.FirstRow != current.FirstRow || plan.Before.FirstColumn != current.FirstColumn ||
            !plan.Before.Contents.ContentEquals(current.Contents))
            return CommandResult.Refused(plan.CommandPlan, "The target contents or shape changed after planning.", RefusalCodes.StaleContext);
        if (plan.ChangedCount == 0) return CommandResult.Success(_descriptor.Id, $"Nothing changed; {plan.SkippedCount:N0} cell(s) skipped.", 0);

        var beforeSerialized = plan.Before.Contents.Serialize();
        var afterSerialized = plan.After.Serialize();

        // A unit transform that also restates the display format has to do both
        // inside one transaction and record both on one receipt, or a single
        // Ctrl+Z would reverse the format and leave the values scaled.
        var formatBefore = string.Empty;
        var formatAfter = string.Empty;
        if (plan.NumberFormat is not null
            && !port.TryRead(plan.Before.Selection.Context, FormatPasteCommand.ReceiptPropertyId, out formatBefore))
        {
            return CommandResult.Refused(plan.CommandPlan,
                "The target's current number formats could not be captured, so the change could not be made undoable.",
                RefusalCodes.CommandUnavailable);
        }

        try
        {
            port.WriteFormulaBlock(plan.Before.Selection.Context, plan.After);
            var observed = port.CaptureFormulaBlock(plan.Before.Selection.Context);
            if (!plan.Before.Selection.Context.Equals(observed.Selection.Context) || !plan.After.ContentEquals(observed.Contents))
                throw new InvalidOperationException("Formula block postcondition mismatch.");
            if (plan.NumberFormat is not null)
            {
                port.ApplyNumberFormat(plan.Before.Selection.Context, plan.NumberFormat);
                if (!port.TryRead(plan.Before.Selection.Context, FormatPasteCommand.ReceiptPropertyId, out formatAfter))
                    throw new InvalidOperationException("The applied number format could not be read back.");
            }
        }
        catch (Exception exception)
        {
            var restored = TryRestore(plan.Before, port);
            if (restored && plan.NumberFormat is not null)
            {
                try { restored = port.TryWrite(plan.Before.Selection.Context, FormatPasteCommand.ReceiptPropertyId, formatBefore); }
                catch { restored = false; }
            }

            if (restored) return CommandResult.Failed(_descriptor.Id,
                $"Formula mutation failed ({exception.GetType().Name}); the entire target was restored to its exact before-state.",
                "FORMULA_WRITE_ROLLED_BACK");
            return CommandResult.Partial(_descriptor.Id,
                $"Formula mutation failed ({exception.GetType().Name}) and exact rollback could not be verified; inspect the target.",
                plan.ChangedCount, plan.SkippedCount, "FORMULA_ROLLBACK_INCOMPLETE");
        }

        var now = DateTimeOffset.UtcNow;
        var receiptId = Guid.NewGuid().ToString("N");
        try
        {
            if (plan.NumberFormat is null)
            {
                receiptSink.Add(new PropertyReceipt(receiptId, _descriptor.Id, _descriptor.ContractVersion,
                    plan.Before.Selection.Context, ReceiptPropertyId, beforeSerialized, afterSerialized,
                    plan.CommandPlan.PlanHash, now, now.AddHours(8)));
            }
            else if (receiptSink is IPropertyBatchReceiptSink batchSink)
            {
                batchSink.Add(new PropertyBatchReceipt(receiptId, _descriptor.Id, _descriptor.ContractVersion,
                    plan.Before.Selection.Context,
                    new[]
                    {
                        new PropertyChange(ReceiptPropertyId, beforeSerialized, afterSerialized),
                        new PropertyChange(FormatPasteCommand.ReceiptPropertyId, formatBefore, formatAfter),
                    },
                    plan.CommandPlan.PlanHash, now, now.AddHours(8)));
            }
            else
            {
                throw new InvalidOperationException(
                    "A transform that also sets a number format needs a batch receipt store so one undo reverses both.");
            }
        }
        catch (Exception exception)
        {
            if (TryRestore(plan.Before, port)) return CommandResult.Failed(_descriptor.Id,
                $"The undo receipt could not be stored ({exception.GetType().Name}); the complete mutation was rolled back.",
                "RECEIPT_STORE_ROLLED_BACK");
            return CommandResult.Partial(_descriptor.Id,
                $"The undo receipt could not be stored ({exception.GetType().Name}) and rollback could not be verified; inspect the target.",
                plan.ChangedCount, plan.SkippedCount, "RECEIPT_STORE_ROLLBACK_INCOMPLETE");
        }
        return CommandResult.Success(_descriptor.Id,
            $"Changed {plan.ChangedCount:N0} cell(s); skipped {plan.SkippedCount:N0}; verified the complete target.",
            plan.ChangedCount, receiptId);
    }

    private FormulaBlockPlan PlanFormulaMap(FormulaBlockSnapshot snapshot, string label,
        Func<string, FormulaTransformResult> transform, int immediatePreviewLimit,
        IEnumerable<string> changedProperties, IEnumerable<KeyValuePair<string, string>> arguments)
    {
        RequireSafe(snapshot);
        var changed = 0;
        var skipped = 0;
        var samples = new List<string>();
        var after = snapshot.Contents.Map((row, column, current) =>
        {
            if (!current.IsFormula) { skipped++; return current; }
            var transformed = transform(current.InvariantValue);
            if (!transformed.IsSuccess) throw Refuse(transformed.RefusalCode!, transformed.Message!);
            var proposed = FormulaCellValue.Formula(transformed.Formula!);
            if (current.Equals(proposed)) return current;
            changed++;
            AddSample(samples, row, column, current, proposed);
            return proposed;
        });
        return BuildPlan(snapshot, after, changed, skipped, samples, changed > immediatePreviewLimit,
            $"{label}: change {changed:N0} formula cell(s); skip {skipped:N0} non-formula cell(s).", changedProperties, arguments);
    }

    private FormulaBlockPlan BuildPlan(FormulaBlockSnapshot snapshot, FormulaCellBlock after,
        int changed, int skipped, IEnumerable<string> samples, bool requiresPreview,
        string summary, IEnumerable<string> changedProperties, IEnumerable<KeyValuePair<string, string>> arguments,
        string? numberFormat = null)
    {
        var fullArguments = arguments.Concat(new[]
        {
            new KeyValuePair<string, string>("before_sha256", snapshot.Contents.Fingerprint),
            new KeyValuePair<string, string>("after_sha256", after.Fingerprint),
            new KeyValuePair<string, string>("rows", after.RowCount.ToString(CultureInfo.InvariantCulture)),
            new KeyValuePair<string, string>("columns", after.ColumnCount.ToString(CultureInfo.InvariantCulture)),
        });
        var commandPlan = new CommandPlan(_descriptor.Id, _descriptor.Impact, snapshot.Selection.Context,
            changedProperties, changed, summary, snapshot.Contents.Fingerprint, _descriptor.ContractVersion,
            requiresPreview, fullArguments);
        return new FormulaBlockPlan(commandPlan, snapshot, after, changed, skipped, samples, numberFormat: numberFormat);
    }

    private static bool TryRestore(FormulaBlockSnapshot before, IFormulaBlockPort port)
    {
        try
        {
            port.WriteFormulaBlock(before.Selection.Context, before.Contents);
            var observed = port.CaptureFormulaBlock(before.Selection.Context);
            return before.Selection.Context.Equals(observed.Selection.Context) && before.Contents.ContentEquals(observed.Contents);
        }
        catch { return false; }
    }

    private static void RequireSafe(FormulaBlockSnapshot snapshot)
    {
        var safety = snapshot.Selection.Safety;
        if (safety.AreaCount != 1 || safety.HasMergedCells)
            throw Refuse(RefusalCodes.SelectionUnsupported, "Formula commands require one unmerged rectangular selection.");
        if (safety.WorksheetProtected || safety.WorkbookReadOnly)
            throw Refuse(RefusalCodes.ProtectedTarget, "The target is protected or read-only.");
        if (!safety.DynamicArraySpillCheckSupported || safety.HasLegacyArray || safety.HasDynamicArraySpill)
            throw Refuse(RefusalCodes.ArrayOrSpillUnsafe, "The target intersects an unqualified array or spill state.");
        if (snapshot.Contents.CellCount > FormulaCellBlock.MaximumCells)
            throw Refuse(RefusalCodes.ResourceLimit, "The formula command selection exceeds its bounded cell limit.");
    }

    private static void AddSample(ICollection<string> samples, int row, int column, FormulaCellValue before, FormulaCellValue after)
    {
        if (samples.Count >= 10) return;
        samples.Add($"R{row + 1}C{column + 1}: {before.Kind} {before.InvariantValue} -> {after.Kind} {after.InvariantValue}");
    }

    private static CommandRefusedException Refuse(string code, string message) =>
        new CommandRefusedException(code, message, "Use a smaller, unmerged, editable selection containing only qualified A1 formulas and supported constants.");
}
