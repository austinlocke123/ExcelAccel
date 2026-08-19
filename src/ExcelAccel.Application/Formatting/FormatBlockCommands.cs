using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Formulas;
using ExcelAccel.Application.Undo;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.Formatting;

public sealed class CellFormatValue : IEquatable<CellFormatValue>
{
    public CellFormatValue(string numberFormat, string fontName, double fontSize, bool fontBold, bool fontItalic,
        string underline, string horizontalAlignment, string verticalAlignment, int indentLevel)
    {
        NumberFormat = StyleValue("number_format", numberFormat);
        FontName = StyleValue("font_name", fontName);
        FontSize = double.Parse(StyleValue("font_size", fontSize.ToString("R", CultureInfo.InvariantCulture)), CultureInfo.InvariantCulture);
        FontBold = fontBold;
        FontItalic = fontItalic;
        Underline = StyleValue("underline", underline);
        HorizontalAlignment = StyleValue("horizontal_alignment", horizontalAlignment);
        VerticalAlignment = StyleValue("vertical_alignment", verticalAlignment);
        IndentLevel = int.Parse(StyleValue("indent_level", indentLevel.ToString(CultureInfo.InvariantCulture)), CultureInfo.InvariantCulture);
    }
    public string NumberFormat { get; }
    public string FontName { get; }
    public double FontSize { get; }
    public bool FontBold { get; }
    public bool FontItalic { get; }
    public string Underline { get; }
    public string HorizontalAlignment { get; }
    public string VerticalAlignment { get; }
    public int IndentLevel { get; }
    public string Serialize() => string.Join("|", new[]
    {
        Encode(NumberFormat), Encode(FontName), FontSize.ToString("R", CultureInfo.InvariantCulture),
        FontBold ? "1" : "0", FontItalic ? "1" : "0", Underline,
        HorizontalAlignment, VerticalAlignment, IndentLevel.ToString(CultureInfo.InvariantCulture),
    });
    public static CellFormatValue Deserialize(string value)
    {
        var parts = (value ?? throw new ArgumentNullException(nameof(value))).Split('|');
        if (parts.Length != 9 || !double.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out var size) ||
            (parts[3] != "0" && parts[3] != "1") || (parts[4] != "0" && parts[4] != "1") ||
            !int.TryParse(parts[8], NumberStyles.None, CultureInfo.InvariantCulture, out var indent))
            throw new FormatException("A cell format value is malformed.");
        try { return new CellFormatValue(Decode(parts[0]), Decode(parts[1]), size, parts[3] == "1", parts[4] == "1", parts[5], parts[6], parts[7], indent); }
        catch (Exception exception) when (exception is ArgumentException || exception is FormatException) { throw new FormatException("A cell format value is invalid.", exception); }
    }
    public bool Equals(CellFormatValue? other) => other is not null &&
        string.Equals(NumberFormat, other.NumberFormat, StringComparison.Ordinal) &&
        string.Equals(FontName, other.FontName, StringComparison.OrdinalIgnoreCase) && FontSize.Equals(other.FontSize) &&
        FontBold == other.FontBold && FontItalic == other.FontItalic &&
        string.Equals(Underline, other.Underline, StringComparison.Ordinal) &&
        string.Equals(HorizontalAlignment, other.HorizontalAlignment, StringComparison.Ordinal) &&
        string.Equals(VerticalAlignment, other.VerticalAlignment, StringComparison.Ordinal) && IndentLevel == other.IndentLevel;
    public override bool Equals(object? obj) => Equals(obj as CellFormatValue);
    public override int GetHashCode() => StringComparer.Ordinal.GetHashCode(Serialize());
    private static string StyleValue(string id, string value) => Styles.StylePropertyCatalog.Normalize(id, value);
    private static string Encode(string value) => Convert.ToBase64String(Encoding.UTF8.GetBytes(value));
    private static string Decode(string value) => Encoding.UTF8.GetString(Convert.FromBase64String(value));
}

public sealed class FormatBlock
{
    public const int MaximumCells = 100;
    public const int MaximumSerializedCharacters = 500_000;
    private readonly CellFormatValue[] _cells;
    public FormatBlock(int rows, int columns, IEnumerable<CellFormatValue> cells)
    {
        if (rows < 1 || columns < 1 || (long)rows * columns > MaximumCells) throw new ArgumentOutOfRangeException(nameof(rows));
        _cells = (cells ?? throw new ArgumentNullException(nameof(cells))).ToArray();
        if (_cells.Length != rows * columns || _cells.Any(value => value is null)) throw new ArgumentException("Format cells must match the block dimensions.", nameof(cells));
        RowCount = rows; ColumnCount = columns;
    }
    public int RowCount { get; }
    public int ColumnCount { get; }
    public int CellCount => _cells.Length;
    public CellFormatValue this[int row, int column] => _cells[checked((row * ColumnCount) + column)];
    public IReadOnlyList<CellFormatValue> Cells => Array.AsReadOnly(_cells);
    public string Serialize()
    {
        var builder = new StringBuilder("FMB1|").Append(RowCount.ToString(CultureInfo.InvariantCulture)).Append('|')
            .Append(ColumnCount.ToString(CultureInfo.InvariantCulture)).Append('|');
        foreach (var cell in _cells)
        {
            builder.Append(Convert.ToBase64String(Encoding.UTF8.GetBytes(cell.Serialize()))).Append('|');
            if (builder.Length > MaximumSerializedCharacters) throw new InvalidOperationException("The format block exceeds its receipt bound.");
        }
        return builder.ToString();
    }
    public static FormatBlock Deserialize(string value)
    {
        if (value is null || value.Length > MaximumSerializedCharacters) throw new ArgumentOutOfRangeException(nameof(value));
        var parts = value.Split('|');
        if (parts.Length < 5 || parts[0] != "FMB1" ||
            !int.TryParse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture, out var rows) ||
            !int.TryParse(parts[2], NumberStyles.None, CultureInfo.InvariantCulture, out var columns) ||
            rows < 1 || columns < 1 || (long)rows * columns > MaximumCells || parts.Length != 4 + rows * columns)
            throw new FormatException("The format block is malformed.");
        var cells = new CellFormatValue[rows * columns];
        for (var index = 0; index < cells.Length; index++)
            cells[index] = CellFormatValue.Deserialize(Encoding.UTF8.GetString(Convert.FromBase64String(parts[index + 3])));
        return new FormatBlock(rows, columns, cells);
    }
    public string Fingerprint => PreconditionFingerprint.Create(Serialize());
    public bool ContentEquals(FormatBlock? other) => other is not null && RowCount == other.RowCount && ColumnCount == other.ColumnCount && _cells.SequenceEqual(other._cells);
}

public sealed class FormatBlockSnapshot
{
    public FormatBlockSnapshot(SelectionSnapshot selection, int firstRow, int firstColumn, FormatBlock contents)
    {
        Selection = selection ?? throw new ArgumentNullException(nameof(selection));
        Contents = contents ?? throw new ArgumentNullException(nameof(contents));
        if (selection.CellCount != contents.CellCount || firstRow < 1 || firstColumn < 1) throw new ArgumentException("Format snapshot bounds are invalid.");
        FirstRow = firstRow; FirstColumn = firstColumn;
    }
    public SelectionSnapshot Selection { get; }
    public int FirstRow { get; }
    public int FirstColumn { get; }
    public FormatBlock Contents { get; }
    public bool ContentEquals(FormatBlockSnapshot? other) => other is not null && FirstRow == other.FirstRow && FirstColumn == other.FirstColumn && Contents.ContentEquals(other.Contents);
}

public interface IFormatBlockPort : ISelectionPort, IPropertyReceiptPort
{
    FormatBlockSnapshot CaptureFormatBlock();
    FormatBlockSnapshot CaptureFormatBlock(SelectionContext target);
    void WriteFormatBlock(SelectionContext target, FormatBlock contents);
}

public sealed class FormatPastePlan
{
    public FormatPastePlan(CommandPlan commandPlan, FormatBlockSnapshot source, FormatBlockSnapshot destination, FormatBlock after, int changedCells)
    { CommandPlan = commandPlan; Source = source; Destination = destination; After = after; ChangedCells = changedCells; }
    public CommandPlan CommandPlan { get; }
    public FormatBlockSnapshot Source { get; }
    public FormatBlockSnapshot Destination { get; }
    public FormatBlock After { get; }
    public int ChangedCells { get; }
}

public sealed class FormatPasteCommand
{
    public const string ReceiptPropertyId = "cell_format_block_v1";
    private readonly CommandDescriptor _descriptor;
    public FormatPasteCommand(CommandDescriptor descriptor) => _descriptor = descriptor ?? throw new ArgumentNullException(nameof(descriptor));

    public FormatPastePlan Plan(FormatBlockSnapshot source, FormatBlockSnapshot destination)
    {
        RequireSafe(source); RequireSafe(destination);
        if (!SameSheet(source, destination)) throw Refuse("Formats-only paste requires source and destination on the same worksheet.");
        if (Overlap(source, destination)) throw Refuse("Formats-only paste refuses overlapping source and destination ranges.");
        if (destination.Contents.RowCount % source.Contents.RowCount != 0 || destination.Contents.ColumnCount % source.Contents.ColumnCount != 0)
            throw Refuse("Destination dimensions must be exact whole multiples of the captured source dimensions.");
        var cells = new CellFormatValue[destination.Contents.CellCount];
        var changed = 0;
        for (var row = 0; row < destination.Contents.RowCount; row++)
            for (var column = 0; column < destination.Contents.ColumnCount; column++)
            {
                var proposed = source.Contents[row % source.Contents.RowCount, column % source.Contents.ColumnCount];
                cells[(row * destination.Contents.ColumnCount) + column] = proposed;
                if (!destination.Contents[row, column].Equals(proposed)) changed++;
            }
        var after = new FormatBlock(destination.Contents.RowCount, destination.Contents.ColumnCount, cells);
        var plan = new CommandPlan(_descriptor.Id, _descriptor.Impact, destination.Selection.Context,
            new[] { ReceiptPropertyId }, changed,
            $"Paste the approved nine-property format set from {source.Contents.RowCount:N0}x{source.Contents.ColumnCount:N0} source into {destination.Contents.RowCount:N0}x{destination.Contents.ColumnCount:N0} destination; {changed:N0} cell(s) differ.",
            PreconditionFingerprint.Create(source.Contents.Serialize(), destination.Contents.Serialize()), _descriptor.ContractVersion,
            requiresPreview: true, new[]
            {
                Pair("source_context", source.Selection.Context.ToString()), Pair("source_sha256", source.Contents.Fingerprint),
                Pair("destination_before_sha256", destination.Contents.Fingerprint), Pair("destination_after_sha256", after.Fingerprint),
                Pair("approved_properties", "number_format|font_name|font_size|font_bold|font_italic|underline|horizontal_alignment|vertical_alignment|indent_level"),
            });
        return new FormatPastePlan(plan, source, destination, after, changed);
    }

    public CommandResult Execute(FormatPastePlan plan, IFormatBlockPort port, string? confirmedPlanHash, IPropertyReceiptSink? receiptSink)
    {
        if (plan is null || port is null) throw new ArgumentNullException(plan is null ? nameof(plan) : nameof(port));
        var authorization = CommandExecutionGate.Authorize(_descriptor, plan.CommandPlan, confirmedPlanHash);
        if (!authorization.Allowed) return CommandResult.Refused(plan.CommandPlan, authorization.Message, authorization.RefusalCode);
        if (receiptSink is null) return CommandResult.Refused(plan.CommandPlan, "Formats-only paste requires an available bounded undo store.", RefusalCodes.CommandUnavailable);
        var source = port.CaptureFormatBlock(plan.Source.Selection.Context);
        if (!plan.Source.ContentEquals(source)) return CommandResult.Refused(plan.CommandPlan, "The format source changed after planning.", RefusalCodes.StaleContext);
        var before = port.CaptureFormatBlock(plan.Destination.Selection.Context);
        if (!plan.Destination.ContentEquals(before)) return CommandResult.Refused(plan.CommandPlan, "The destination format changed after planning.", RefusalCodes.StaleContext);
        if (plan.ChangedCells == 0) return CommandResult.Success(_descriptor.Id, "Destination formatting already matches the exact plan.");
        var beforeSerialized = plan.Destination.Contents.Serialize();
        var afterSerialized = plan.After.Serialize();
        try
        {
            port.WriteFormatBlock(plan.Destination.Selection.Context, plan.After);
            var observed = port.CaptureFormatBlock(plan.Destination.Selection.Context);
            if (!plan.After.ContentEquals(observed.Contents)) throw new InvalidOperationException("Format paste postcondition mismatch.");
        }
        catch (Exception)
        {
            try
            {
                port.WriteFormatBlock(plan.Destination.Selection.Context, plan.Destination.Contents);
                if (plan.Destination.Contents.ContentEquals(port.CaptureFormatBlock(plan.Destination.Selection.Context).Contents))
                    return CommandResult.Failed(_descriptor.Id, "Formats-only paste failed; the exact destination before-state was restored.", "FORMAT_PASTE_ROLLED_BACK");
            }
            catch { }
            return CommandResult.Partial(_descriptor.Id, "Formats-only paste failed and exact rollback could not be verified; inspect the bounded destination.", plan.ChangedCells, 0, "FORMAT_PASTE_ROLLBACK_INCOMPLETE");
        }
        var receiptId = Guid.NewGuid().ToString("N");
        var now = DateTimeOffset.UtcNow;
        try
        {
            receiptSink.Add(new PropertyReceipt(receiptId, _descriptor.Id, _descriptor.ContractVersion, plan.Destination.Selection.Context,
                ReceiptPropertyId, beforeSerialized, afterSerialized, plan.CommandPlan.PlanHash, now, now.AddHours(8)));
        }
        catch (Exception exception)
        {
            try
            {
                port.WriteFormatBlock(plan.Destination.Selection.Context, plan.Destination.Contents);
                if (plan.Destination.Contents.ContentEquals(port.CaptureFormatBlock(plan.Destination.Selection.Context).Contents))
                    return CommandResult.Failed(_descriptor.Id, $"The undo receipt could not be stored ({exception.GetType().Name}); the format paste was rolled back.", "RECEIPT_STORE_ROLLED_BACK");
            }
            catch { }
            return CommandResult.Partial(_descriptor.Id, $"The undo receipt could not be stored ({exception.GetType().Name}) and rollback could not be verified; inspect the target.",
                plan.ChangedCells, 0, "RECEIPT_STORE_ROLLBACK_INCOMPLETE");
        }
        return CommandResult.Success(plan.CommandPlan, $"Pasted and verified the approved formatting set in {plan.ChangedCells:N0} cell(s).", receiptId);
    }

    private static bool SameSheet(FormatBlockSnapshot left, FormatBlockSnapshot right) => left.Selection.Context.WorkbookId == right.Selection.Context.WorkbookId && left.Selection.Context.WorksheetName == right.Selection.Context.WorksheetName;
    private static bool Overlap(FormatBlockSnapshot left, FormatBlockSnapshot right) => left.FirstRow <= right.FirstRow + right.Contents.RowCount - 1 && right.FirstRow <= left.FirstRow + left.Contents.RowCount - 1 && left.FirstColumn <= right.FirstColumn + right.Contents.ColumnCount - 1 && right.FirstColumn <= left.FirstColumn + left.Contents.ColumnCount - 1;
    private static void RequireSafe(FormatBlockSnapshot snapshot)
    {
        var safety = snapshot.Selection.Safety;
        if (safety.AreaCount != 1 || safety.HasMergedCells || safety.WorksheetProtected || safety.WorkbookReadOnly ||
            !safety.DynamicArraySpillCheckSupported || safety.HasLegacyArray || safety.HasDynamicArraySpill)
            throw Refuse("Formats-only paste requires one editable, unmerged rectangle outside arrays and spills.");
    }
    private static KeyValuePair<string, string> Pair(string key, string value) => new KeyValuePair<string, string>(key, value);
    private static CommandRefusedException Refuse(string message) => new CommandRefusedException(RefusalCodes.SelectionUnsupported, message, $"Use nonoverlapping exact-multiple rectangles of at most {FormatBlock.MaximumCells:N0} cells.");
}
