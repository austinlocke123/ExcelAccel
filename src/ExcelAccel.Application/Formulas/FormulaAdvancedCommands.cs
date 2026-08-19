using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Commands;
using ExcelAccel.Core.Formulas;

namespace ExcelAccel.Application.Formulas;

public enum FormulaSpacingDirection { Rows, Columns }
public enum SequenceFillDirection { Down, Right }
public enum ExcelDateSystem { Excel1900, Excel1904 }

public sealed class FormulaAdvancedCommand
{
    private readonly CommandDescriptor _descriptor;
    private readonly A1FormulaTransformer _transformer;

    public FormulaAdvancedCommand(CommandDescriptor descriptor, A1FormulaTransformer? transformer = null)
    {
        _descriptor = descriptor ?? throw new ArgumentNullException(nameof(descriptor));
        _transformer = transformer ?? new A1FormulaTransformer();
    }

    public FormulaBlockPlan PlanSpacing(FormulaBlockSnapshot snapshot, FormulaSpacingDirection direction,
        int interval, int immediatePreviewLimit = FormulaBlockCommand.DefaultImmediatePreviewLimit)
    {
        RequireSafe(snapshot);
        if (interval < 1 || interval > 10000) throw Refuse(FormulaTransformRefusalCodes.InvalidTransformArgument, "Spacing interval must be from 1 through 10,000.");
        if ((direction == FormulaSpacingDirection.Rows && snapshot.Contents.RowCount <= interval) ||
            (direction == FormulaSpacingDirection.Columns && snapshot.Contents.ColumnCount <= interval))
            throw Refuse(RefusalCodes.SelectionUnsupported, "The selection has no destination at the requested spacing interval.");
        var changed = 0;
        var overwrites = 0;
        var samples = new List<string>();
        var after = snapshot.Contents.Map((row, column, current) =>
        {
            var isDestination = direction == FormulaSpacingDirection.Rows
                ? row > 0 && row % interval == 0
                : column > 0 && column % interval == 0;
            if (!isDestination) return current;
            var source = direction == FormulaSpacingDirection.Rows ? snapshot.Contents[0, column] : snapshot.Contents[row, 0];
            if (!source.IsFormula) throw Refuse(RefusalCodes.SelectionUnsupported, "Every spacing source-edge cell must contain a qualified formula.");
            var transformed = _transformer.Translate(source.InvariantValue,
                direction == FormulaSpacingDirection.Rows ? row : 0,
                direction == FormulaSpacingDirection.Columns ? column : 0);
            if (!transformed.IsSuccess) throw Refuse(transformed.RefusalCode!, transformed.Message!);
            var proposed = FormulaCellValue.Formula(transformed.Formula!);
            if (current.Equals(proposed)) return current;
            changed++;
            if (!current.IsBlank) overwrites++;
            AddSample(samples, row, column, current, proposed);
            return proposed;
        });
        return Build(snapshot, after, changed, 0, samples, overwrites > 0 || changed > immediatePreviewLimit,
            $"Space formulas every {interval:N0} {direction.ToString().ToLowerInvariant()}: {changed:N0} destinations.",
            new[] { "formula" }, new[]
            {
                Pair("direction", direction.ToString().ToLowerInvariant()),
                Pair("interval", interval.ToString(CultureInfo.InvariantCulture)),
            }, snapshot.Contents.Fingerprint);
    }

    public FormulaBlockPlan PlanTranspose(FormulaBlockSnapshot source, FormulaBlockSnapshot destination)
    {
        RequireSafe(source);
        RequireSafe(destination);
        if (!string.Equals(source.Selection.Context.WorkbookId, destination.Selection.Context.WorkbookId, StringComparison.Ordinal) ||
            !string.Equals(source.Selection.Context.WorksheetName, destination.Selection.Context.WorksheetName, StringComparison.Ordinal))
            throw Refuse(RefusalCodes.SelectionUnsupported, "Initial transpose requires source and destination on the same worksheet.");
        if (destination.Contents.RowCount != source.Contents.ColumnCount || destination.Contents.ColumnCount != source.Contents.RowCount)
            throw Refuse(RefusalCodes.SelectionUnsupported, "The destination dimensions must be the exact transpose of the source.");
        var changed = 0;
        var samples = new List<string>();
        var afterCells = new FormulaCellValue[destination.Contents.CellCount];
        for (var row = 0; row < destination.Contents.RowCount; row++)
        {
            for (var column = 0; column < destination.Contents.ColumnCount; column++)
            {
                var sourceCell = source.Contents[column, row];
                FormulaCellValue proposed;
                if (sourceCell.IsFormula)
                {
                    var transformed = _transformer.Transpose(sourceCell.InvariantValue,
                        source.FirstRow + column, source.FirstColumn + row,
                        destination.FirstRow + row, destination.FirstColumn + column);
                    if (!transformed.IsSuccess) throw Refuse(transformed.RefusalCode!, transformed.Message!);
                    proposed = FormulaCellValue.Formula(transformed.Formula!);
                }
                else proposed = sourceCell;
                afterCells[(row * destination.Contents.ColumnCount) + column] = proposed;
                var current = destination.Contents[row, column];
                if (!current.Equals(proposed)) { changed++; AddSample(samples, row, column, current, proposed); }
            }
        }
        var after = new FormulaCellBlock(destination.Contents.RowCount, destination.Contents.ColumnCount, afterCells);
        return Build(destination, after, changed, 0, samples, requiresPreview: true,
            $"Transpose {source.Contents.RowCount:N0}×{source.Contents.ColumnCount:N0} source cells into the complete {destination.Contents.RowCount:N0}×{destination.Contents.ColumnCount:N0} destination.",
            new[] { "formula", "value" }, new[]
            {
                Pair("source_context", source.Selection.Context.ToString()),
                Pair("source_sha256", source.Contents.Fingerprint),
            }, PreconditionFingerprint.Create(source.Contents.Serialize(), destination.Contents.Serialize()), requiresExternalSourceRevalidation: true);
    }

    public FormulaBlockPlan PlanFormulaFromAbove(FormulaBlockSnapshot sourceRow, FormulaBlockSnapshot destination,
        int immediatePreviewLimit = FormulaBlockCommand.DefaultImmediatePreviewLimit)
    {
        RequireSafe(sourceRow);
        RequireSafe(destination);
        if (sourceRow.Contents.RowCount != 1 || sourceRow.Contents.ColumnCount != destination.Contents.ColumnCount ||
            sourceRow.FirstRow + 1 != destination.FirstRow || sourceRow.FirstColumn != destination.FirstColumn ||
            !SameSheet(sourceRow, destination))
            throw Refuse(RefusalCodes.SelectionUnsupported, "The source must be the single immediately adjacent row above the destination.");
        var changed = 0;
        var overwrites = 0;
        var samples = new List<string>();
        var after = destination.Contents.Map((row, column, current) =>
        {
            var source = sourceRow.Contents[0, column];
            if (!source.IsFormula) throw Refuse(RefusalCodes.SelectionUnsupported, "Every immediately-above source cell must contain a qualified formula.");
            var transformed = _transformer.Translate(source.InvariantValue, row + 1, 0);
            if (!transformed.IsSuccess) throw Refuse(transformed.RefusalCode!, transformed.Message!);
            var proposed = FormulaCellValue.Formula(transformed.Formula!);
            if (current.Equals(proposed)) return current;
            changed++;
            if (!current.IsBlank) overwrites++;
            AddSample(samples, row, column, current, proposed);
            return proposed;
        });
        return Build(destination, after, changed, 0, samples, overwrites > 0 || changed > immediatePreviewLimit,
            $"Fill {changed:N0} destination cell(s) from the immediately adjacent formula row above.",
            new[] { "formula" }, new[] { Pair("source_sha256", sourceRow.Contents.Fingerprint) },
            PreconditionFingerprint.Create(sourceRow.Contents.Serialize(), destination.Contents.Serialize()), requiresExternalSourceRevalidation: true);
    }

    public FormulaBlockPlan PlanNumericSequence(FormulaBlockSnapshot destination, double start, double step,
        SequenceFillDirection direction)
    {
        RequireSafe(destination);
        if (double.IsNaN(start) || double.IsInfinity(start) || double.IsNaN(step) || double.IsInfinity(step))
            throw Refuse(FormulaTransformRefusalCodes.InvalidTransformArgument, "Numeric sequence start and step must be finite.");
        var changed = 0;
        var samples = new List<string>();
        var cells = new FormulaCellValue[destination.Contents.CellCount];
        for (var row = 0; row < destination.Contents.RowCount; row++)
        {
            for (var column = 0; column < destination.Contents.ColumnCount; column++)
            {
                var index = direction == SequenceFillDirection.Down
                    ? (column * destination.Contents.RowCount) + row
                    : (row * destination.Contents.ColumnCount) + column;
                var output = start + (step * index);
                if (double.IsNaN(output) || double.IsInfinity(output))
                    throw Refuse(FormulaTransformRefusalCodes.InvalidTransformArgument, "The numeric sequence would overflow.");
                var proposed = FormulaCellValue.Number(output);
                cells[(row * destination.Contents.ColumnCount) + column] = proposed;
                var current = destination.Contents[row, column];
                if (!current.Equals(proposed)) { changed++; AddSample(samples, row, column, current, proposed); }
            }
        }
        var after = new FormulaCellBlock(destination.Contents.RowCount, destination.Contents.ColumnCount, cells);
        var overwrite = destination.Contents.Cells.Any(value => !value.IsBlank);
        return Build(destination, after, changed, 0, samples, overwrite || changed > FormulaBlockCommand.DefaultImmediatePreviewLimit,
            $"Fill {changed:N0} cell(s) with start {start:R}, step {step:R}, direction {direction.ToString().ToLowerInvariant()}.",
            new[] { "value" }, new[] { Pair("start", start.ToString("R", CultureInfo.InvariantCulture)), Pair("step", step.ToString("R", CultureInfo.InvariantCulture)), Pair("direction", direction.ToString().ToLowerInvariant()) },
            destination.Contents.Fingerprint);
    }

    public FormulaBlockPlan PlanDateSequence(FormulaBlockSnapshot destination, DateTime startDate, int stepDays,
        SequenceFillDirection direction, ExcelDateSystem dateSystem)
    {
        if (startDate.TimeOfDay != TimeSpan.Zero) throw Refuse(FormulaTransformRefusalCodes.InvalidTransformArgument, "Date fill requires a date without a time component.");
        var serialStart = ToExcelSerial(startDate, dateSystem);
        var plan = PlanNumericSequence(destination, serialStart, stepDays, direction);
        var arguments = plan.CommandPlan.Arguments
            .Where(value => value.Key != "start" && value.Key != "step")
            .Concat(new[]
            {
                Pair("start_date", startDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)),
                Pair("step_days", stepDays.ToString(CultureInfo.InvariantCulture)),
                Pair("date_system", dateSystem == ExcelDateSystem.Excel1900 ? "1900" : "1904"),
            });
        var commandPlan = new CommandPlan(_descriptor.Id, _descriptor.Impact, plan.CommandPlan.Context,
            new[] { "value" }, plan.ChangedCount,
            $"Fill {plan.ChangedCount:N0} date cell(s) from {startDate:yyyy-MM-dd} by {stepDays:N0} day(s) using the Excel {(dateSystem == ExcelDateSystem.Excel1900 ? "1900" : "1904")} date system.",
            plan.CommandPlan.PreconditionFingerprint, _descriptor.ContractVersion, plan.CommandPlan.RequiresPreview, arguments);
        return new FormulaBlockPlan(commandPlan, plan.Before, plan.After, plan.ChangedCount, 0, plan.Samples);
    }

    public static double ToExcelSerial(DateTime date, ExcelDateSystem dateSystem)
    {
        var value = date.Date;
        if (dateSystem == ExcelDateSystem.Excel1904)
        {
            var baseDate = new DateTime(1904, 1, 1);
            if (value < baseDate) throw Refuse(FormulaTransformRefusalCodes.InvalidTransformArgument, "A 1904-system date cannot precede 1904-01-01.");
            return (value - baseDate).TotalDays;
        }
        var excelBase = new DateTime(1899, 12, 31);
        if (value < excelBase) throw Refuse(FormulaTransformRefusalCodes.InvalidTransformArgument, "A 1900-system date cannot precede 1899-12-31.");
        var serial = (value - excelBase).TotalDays;
        if (value >= new DateTime(1900, 3, 1)) serial++;
        return serial;
    }

    private FormulaBlockPlan Build(FormulaBlockSnapshot before, FormulaCellBlock after, int changed, int skipped,
        IEnumerable<string> samples, bool requiresPreview, string summary, IEnumerable<string> changedProperties,
        IEnumerable<KeyValuePair<string, string>> arguments, string preconditionFingerprint,
        bool requiresExternalSourceRevalidation = false)
    {
        var fullArguments = arguments.Concat(new[]
        {
            Pair("destination_before_sha256", before.Contents.Fingerprint),
            Pair("destination_after_sha256", after.Fingerprint),
            Pair("rows", after.RowCount.ToString(CultureInfo.InvariantCulture)),
            Pair("columns", after.ColumnCount.ToString(CultureInfo.InvariantCulture)),
        });
        var commandPlan = new CommandPlan(_descriptor.Id, _descriptor.Impact, before.Selection.Context,
            changedProperties, changed, summary, preconditionFingerprint, _descriptor.ContractVersion, requiresPreview, fullArguments);
        return new FormulaBlockPlan(commandPlan, before, after, changed, skipped, samples, requiresExternalSourceRevalidation);
    }

    private static void RequireSafe(FormulaBlockSnapshot snapshot)
    {
        var safety = snapshot.Selection.Safety;
        if (safety.AreaCount != 1 || safety.HasMergedCells) throw Refuse(RefusalCodes.SelectionUnsupported, "The command requires one unmerged rectangle.");
        if (safety.WorksheetProtected || safety.WorkbookReadOnly) throw Refuse(RefusalCodes.ProtectedTarget, "The target is protected or read-only.");
        if (!safety.DynamicArraySpillCheckSupported || safety.HasLegacyArray || safety.HasDynamicArraySpill)
            throw Refuse(RefusalCodes.ArrayOrSpillUnsafe, "The range intersects an unqualified array or spill state.");
    }

    private static bool SameSheet(FormulaBlockSnapshot first, FormulaBlockSnapshot second) =>
        string.Equals(first.Selection.Context.WorkbookId, second.Selection.Context.WorkbookId, StringComparison.Ordinal) &&
        string.Equals(first.Selection.Context.WorksheetName, second.Selection.Context.WorksheetName, StringComparison.Ordinal);
    private static KeyValuePair<string, string> Pair(string key, string value) => new KeyValuePair<string, string>(key, value);
    private static void AddSample(ICollection<string> samples, int row, int column, FormulaCellValue before, FormulaCellValue after)
    {
        if (samples.Count < 10) samples.Add($"R{row + 1}C{column + 1}: {before.Kind} {before.InvariantValue} -> {after.Kind} {after.InvariantValue}");
    }
    private static CommandRefusedException Refuse(string code, string message) => new CommandRefusedException(code, message, "Use an exact bounded, editable rectangle and qualified A1 formulas.");
}
