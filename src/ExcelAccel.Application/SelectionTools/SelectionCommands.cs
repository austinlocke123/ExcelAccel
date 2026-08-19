using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Formulas;
using ExcelAccel.Core.Commands;
using ExcelAccel.Core.Formulas;

namespace ExcelAccel.Application.SelectionTools;

public enum SelectionPredicate
{
    Formulas,
    Constants,
    Blanks,
    NumericHardcodes,
    ExternalFormulas,
}

public sealed class SelectionArea
{
    public SelectionArea(int firstRow, int firstColumn, int lastRow, int lastColumn)
    {
        if (firstRow < 1 || firstColumn < 1 || lastRow < firstRow || lastColumn < firstColumn)
            throw new ArgumentOutOfRangeException(nameof(firstRow));
        FirstRow = firstRow; FirstColumn = firstColumn; LastRow = lastRow; LastColumn = lastColumn;
    }
    public int FirstRow { get; }
    public int FirstColumn { get; }
    public int LastRow { get; }
    public int LastColumn { get; }
    public long CellCount => checked((long)(LastRow - FirstRow + 1) * (LastColumn - FirstColumn + 1));
    public string Address => CellAddress(FirstRow, FirstColumn) +
        (FirstRow == LastRow && FirstColumn == LastColumn ? string.Empty : ":" + CellAddress(LastRow, LastColumn));

    private static string CellAddress(int row, int column)
    {
        var name = string.Empty;
        var remaining = column;
        while (remaining > 0) { remaining--; name = (char)('A' + remaining % 26) + name; remaining /= 26; }
        return name + row.ToString(CultureInfo.InvariantCulture);
    }
}

public sealed class SelectionMatchPlan
{
    public SelectionMatchPlan(CommandPlan commandPlan, FormulaBlockSnapshot source, IEnumerable<SelectionArea> areas, SelectionPredicate predicate)
    {
        CommandPlan = commandPlan ?? throw new ArgumentNullException(nameof(commandPlan));
        Source = source ?? throw new ArgumentNullException(nameof(source));
        Areas = Array.AsReadOnly((areas ?? throw new ArgumentNullException(nameof(areas))).ToArray());
        Predicate = predicate;
    }
    public CommandPlan CommandPlan { get; }
    public FormulaBlockSnapshot Source { get; }
    public IReadOnlyList<SelectionArea> Areas { get; }
    public SelectionPredicate Predicate { get; }
}

public interface ISelectionMatchPort : IFormulaBlockPort
{
    void SelectAreas(SelectionContext sourceContext, IReadOnlyList<SelectionArea> areas);
    IReadOnlyList<string> CaptureSelectedAreaAddresses();
}

public sealed class SelectionMatchCommand
{
    public const int MaximumAreas = 64;
    public const int MaximumAddressCharacters = 4096;
    private readonly CommandDescriptor _descriptor;
    private readonly FormulaParser _parser = new FormulaParser();

    public SelectionMatchCommand(CommandDescriptor descriptor) => _descriptor = descriptor ?? throw new ArgumentNullException(nameof(descriptor));

    public SelectionMatchPlan Plan(FormulaBlockSnapshot source, SelectionPredicate predicate)
    {
        if (source.Selection.Safety.AreaCount != 1 || source.Selection.Safety.HasMergedCells)
            throw Refuse(RefusalCodes.SelectionUnsupported, "Selection matching requires one unmerged rectangular source range.");
        var matches = new bool[source.Contents.RowCount, source.Contents.ColumnCount];
        var matched = 0;
        for (var row = 0; row < source.Contents.RowCount; row++)
        {
            for (var column = 0; column < source.Contents.ColumnCount; column++)
            {
                if (!Matches(source.Contents[row, column], predicate)) continue;
                matches[row, column] = true;
                matched++;
            }
        }
        if (matched == 0) throw Refuse(RefusalCodes.SelectionUnsupported, $"No cells match the {predicate.ToString().ToLowerInvariant()} predicate.");
        var areas = Compress(matches, source.FirstRow, source.FirstColumn);
        if (areas.Count > MaximumAreas)
            throw Refuse(RefusalCodes.ResourceLimit, $"The match set requires {areas.Count:N0} areas, above the {MaximumAreas:N0}-area limit.");
        var addressCharacters = areas.Sum(area => area.Address.Length + 1);
        if (addressCharacters > MaximumAddressCharacters)
            throw Refuse(RefusalCodes.ResourceLimit, "The exact multi-area address exceeds the qualified Excel address limit.");
        var plan = new CommandPlan(_descriptor.Id, CommandImpact.ReadOnly, source.Selection.Context,
            Array.Empty<string>(), matched,
            $"Select {matched:N0} {predicate.ToString().ToLowerInvariant()} cell(s) in {areas.Count:N0} deterministic area(s).",
            source.Contents.Fingerprint, _descriptor.ContractVersion, requiresPreview: false,
            new[]
            {
                Pair("predicate", predicate.ToString().ToLowerInvariant()),
                Pair("source_sha256", source.Contents.Fingerprint),
                Pair("area_count", areas.Count.ToString(CultureInfo.InvariantCulture)),
            });
        return new SelectionMatchPlan(plan, source, areas, predicate);
    }

    public CommandResult Execute(SelectionMatchPlan plan, ISelectionMatchPort port)
    {
        if (plan is null) throw new ArgumentNullException(nameof(plan));
        if (port is null) throw new ArgumentNullException(nameof(port));
        var authorization = CommandExecutionGate.Authorize(_descriptor, plan.CommandPlan);
        if (!authorization.Allowed) return CommandResult.Refused(plan.CommandPlan, authorization.Message, authorization.RefusalCode);
        var current = port.CaptureFormulaBlock(plan.Source.Selection.Context);
        if (current.FirstRow != plan.Source.FirstRow || current.FirstColumn != plan.Source.FirstColumn ||
            !current.Contents.ContentEquals(plan.Source.Contents))
            return CommandResult.Refused(plan.CommandPlan, "The source range changed after selection planning.", RefusalCodes.StaleContext);
        port.SelectAreas(plan.Source.Selection.Context, plan.Areas);
        var observed = port.CaptureSelectedAreaAddresses();
        var expected = plan.Areas.Select(area => area.Address).ToArray();
        if (!expected.SequenceEqual(observed, StringComparer.OrdinalIgnoreCase))
            return CommandResult.Failed(_descriptor.Id, "Excel did not report the exact planned selected areas.", "SELECTION_POSTCONDITION_MISMATCH");
        return CommandResult.Success(_descriptor.Id, plan.CommandPlan.Summary, plan.CommandPlan.AffectedCellCount);
    }

    private bool Matches(FormulaCellValue cell, SelectionPredicate predicate)
    {
        switch (predicate)
        {
            case SelectionPredicate.Formulas: return cell.IsFormula;
            case SelectionPredicate.Constants: return !cell.IsBlank && !cell.IsFormula;
            case SelectionPredicate.Blanks: return cell.IsBlank;
            case SelectionPredicate.NumericHardcodes: return cell.Kind == FormulaCellKind.Number;
            case SelectionPredicate.ExternalFormulas:
                if (!cell.IsFormula) return false;
                var parsed = _parser.Parse(cell.InvariantValue);
                return parsed.IsSuccess && parsed.Document!.References.Any(reference =>
                    reference.Qualifier is not null && reference.Qualifier.IndexOf('[') >= 0);
            default: return false;
        }
    }

    private static IReadOnlyList<SelectionArea> Compress(bool[,] matches, int firstRow, int firstColumn)
    {
        var rows = matches.GetLength(0);
        var columns = matches.GetLength(1);
        var completed = new List<SelectionArea>();
        var active = new Dictionary<string, SelectionArea>(StringComparer.Ordinal);
        for (var row = 0; row < rows; row++)
        {
            var runs = new List<(int Start, int End)>();
            var column = 0;
            while (column < columns)
            {
                if (!matches[row, column]) { column++; continue; }
                var start = column;
                while (column + 1 < columns && matches[row, column + 1]) column++;
                runs.Add((start, column));
                column++;
            }
            var next = new Dictionary<string, SelectionArea>(StringComparer.Ordinal);
            foreach (var run in runs)
            {
                var key = run.Start.ToString(CultureInfo.InvariantCulture) + ":" + run.End.ToString(CultureInfo.InvariantCulture);
                if (active.TryGetValue(key, out var prior))
                    next[key] = new SelectionArea(prior.FirstRow, prior.FirstColumn, firstRow + row, firstColumn + run.End);
                else
                    next[key] = new SelectionArea(firstRow + row, firstColumn + run.Start, firstRow + row, firstColumn + run.End);
            }
            foreach (var item in active) if (!next.ContainsKey(item.Key)) completed.Add(item.Value);
            active = next;
        }
        completed.AddRange(active.Values);
        return completed.OrderBy(area => area.FirstRow).ThenBy(area => area.FirstColumn).ToArray();
    }

    private static KeyValuePair<string, string> Pair(string key, string value) => new KeyValuePair<string, string>(key, value);
    private static CommandRefusedException Refuse(string code, string message) =>
        new CommandRefusedException(code, message, "Use a smaller single rectangle or narrow the source range.");
}
