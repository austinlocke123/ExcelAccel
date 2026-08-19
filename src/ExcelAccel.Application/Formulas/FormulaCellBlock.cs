using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.Formulas;

public enum FormulaCellKind
{
    Blank,
    Formula,
    Text,
    Number,
    Boolean,
}

public sealed class FormulaCellValue : IEquatable<FormulaCellValue>
{
    private FormulaCellValue(FormulaCellKind kind, string invariantValue)
    {
        Kind = kind;
        InvariantValue = invariantValue ?? throw new ArgumentNullException(nameof(invariantValue));
        if (kind == FormulaCellKind.Formula && (invariantValue.Length < 2 || invariantValue[0] != '='))
            throw new ArgumentException("Formula cell text must begin with '='.", nameof(invariantValue));
        if (kind == FormulaCellKind.Blank && invariantValue.Length != 0)
            throw new ArgumentException("A blank cell cannot carry a value.", nameof(invariantValue));
    }

    public FormulaCellKind Kind { get; }
    public string InvariantValue { get; }
    public bool IsBlank => Kind == FormulaCellKind.Blank;
    public bool IsFormula => Kind == FormulaCellKind.Formula;

    public static FormulaCellValue Blank() => new FormulaCellValue(FormulaCellKind.Blank, string.Empty);
    public static FormulaCellValue Formula(string formula) => new FormulaCellValue(FormulaCellKind.Formula, formula);
    public static FormulaCellValue Text(string value) => new FormulaCellValue(FormulaCellKind.Text, value ?? string.Empty);
    public static FormulaCellValue Number(double value)
    {
        if (double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(nameof(value));
        return new FormulaCellValue(FormulaCellKind.Number, value.ToString("R", CultureInfo.InvariantCulture));
    }
    public static FormulaCellValue Boolean(bool value) => new FormulaCellValue(FormulaCellKind.Boolean, value ? "true" : "false");

    public double AsNumber()
    {
        if (Kind != FormulaCellKind.Number) throw new InvalidOperationException("The cell is not numeric.");
        return double.Parse(InvariantValue, NumberStyles.Float, CultureInfo.InvariantCulture);
    }

    public bool Equals(FormulaCellValue? other) =>
        other is not null && Kind == other.Kind && string.Equals(InvariantValue, other.InvariantValue, StringComparison.Ordinal);
    public override bool Equals(object? obj) => Equals(obj as FormulaCellValue);
    public override int GetHashCode() => ((int)Kind * 397) ^ StringComparer.Ordinal.GetHashCode(InvariantValue);
}

public sealed class FormulaCellBlock
{
    public const int MaximumCells = 10_000;
    public const int MaximumSerializedCharacters = 1_000_000;
    private readonly FormulaCellValue[] _cells;

    public FormulaCellBlock(int rowCount, int columnCount, IEnumerable<FormulaCellValue> cells)
    {
        if (rowCount < 1) throw new ArgumentOutOfRangeException(nameof(rowCount));
        if (columnCount < 1) throw new ArgumentOutOfRangeException(nameof(columnCount));
        if ((long)rowCount * columnCount > MaximumCells) throw new ArgumentOutOfRangeException(nameof(rowCount), $"A formula block cannot exceed {MaximumCells:N0} cells.");
        _cells = (cells ?? throw new ArgumentNullException(nameof(cells))).ToArray();
        if (_cells.Length != rowCount * columnCount) throw new ArgumentException("The cell count does not match the block dimensions.", nameof(cells));
        if (_cells.Any(value => value is null)) throw new ArgumentException("Formula block cells cannot be null.", nameof(cells));
        RowCount = rowCount;
        ColumnCount = columnCount;
    }

    public int RowCount { get; }
    public int ColumnCount { get; }
    public int CellCount => _cells.Length;
    public IReadOnlyList<FormulaCellValue> Cells => Array.AsReadOnly(_cells);
    public FormulaCellValue this[int row, int column] => _cells[Offset(row, column)];

    public FormulaCellBlock Map(Func<int, int, FormulaCellValue, FormulaCellValue> transform)
    {
        if (transform is null) throw new ArgumentNullException(nameof(transform));
        var cells = new FormulaCellValue[_cells.Length];
        for (var row = 0; row < RowCount; row++)
            for (var column = 0; column < ColumnCount; column++)
                cells[Offset(row, column)] = transform(row, column, this[row, column]) ?? throw new InvalidOperationException("A cell transform returned null.");
        return new FormulaCellBlock(RowCount, ColumnCount, cells);
    }

    public string Serialize()
    {
        var builder = new StringBuilder();
        builder.Append("FCB1|").Append(RowCount.ToString(CultureInfo.InvariantCulture)).Append('|')
            .Append(ColumnCount.ToString(CultureInfo.InvariantCulture)).Append('|');
        foreach (var cell in _cells)
        {
            var bytes = Encoding.UTF8.GetBytes(cell.InvariantValue);
            builder.Append(((int)cell.Kind).ToString(CultureInfo.InvariantCulture)).Append(':')
                .Append(Convert.ToBase64String(bytes)).Append('|');
            if (builder.Length > MaximumSerializedCharacters)
                throw new InvalidOperationException("The formula block exceeds the serialized receipt limit.");
        }
        return builder.ToString();
    }

    public static FormulaCellBlock Deserialize(string serialized)
    {
        if (serialized is null) throw new ArgumentNullException(nameof(serialized));
        if (serialized.Length > MaximumSerializedCharacters) throw new ArgumentOutOfRangeException(nameof(serialized));
        var parts = serialized.Split('|');
        if (parts.Length < 5 || parts[0] != "FCB1" ||
            !int.TryParse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture, out var rows) ||
            !int.TryParse(parts[2], NumberStyles.None, CultureInfo.InvariantCulture, out var columns) ||
            rows < 1 || columns < 1 || (long)rows * columns > MaximumCells || parts.Length != 4 + (rows * columns))
            throw new FormatException("The formula block receipt is malformed or outside its bounds.");
        var cells = new List<FormulaCellValue>(rows * columns);
        for (var index = 0; index < rows * columns; index++)
        {
            var part = parts[index + 3];
            var separator = part.IndexOf(':');
            if (separator != 1 || !int.TryParse(part.Substring(0, separator), NumberStyles.None, CultureInfo.InvariantCulture, out var kindValue) ||
                !Enum.IsDefined(typeof(FormulaCellKind), kindValue)) throw new FormatException("A formula block cell kind is invalid.");
            string value;
            try { value = Encoding.UTF8.GetString(Convert.FromBase64String(part.Substring(separator + 1))); }
            catch (FormatException) { throw new FormatException("A formula block cell value is not valid base64."); }
            cells.Add(Create((FormulaCellKind)kindValue, value));
        }
        return new FormulaCellBlock(rows, columns, cells);
    }

    public string Fingerprint => PreconditionFingerprint.Create(Serialize());

    public bool ContentEquals(FormulaCellBlock? other) => other is not null && RowCount == other.RowCount && ColumnCount == other.ColumnCount && _cells.SequenceEqual(other._cells);

    private int Offset(int row, int column)
    {
        if (row < 0 || row >= RowCount) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 0 || column >= ColumnCount) throw new ArgumentOutOfRangeException(nameof(column));
        return checked((row * ColumnCount) + column);
    }

    private static FormulaCellValue Create(FormulaCellKind kind, string value)
    {
        switch (kind)
        {
            case FormulaCellKind.Blank: return FormulaCellValue.Blank();
            case FormulaCellKind.Formula: return FormulaCellValue.Formula(value);
            case FormulaCellKind.Text: return FormulaCellValue.Text(value);
            case FormulaCellKind.Number:
                if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) || double.IsNaN(number) || double.IsInfinity(number))
                    throw new FormatException("A numeric formula block cell is invalid.");
                return FormulaCellValue.Number(number);
            case FormulaCellKind.Boolean:
                if (value == "true") return FormulaCellValue.Boolean(true);
                if (value == "false") return FormulaCellValue.Boolean(false);
                throw new FormatException("A Boolean formula block cell is invalid.");
            default: throw new FormatException("A formula block cell kind is invalid.");
        }
    }
}

public sealed class FormulaBlockSnapshot
{
    public FormulaBlockSnapshot(SelectionSnapshot selection, int firstRow, int firstColumn, FormulaCellBlock contents)
    {
        Selection = selection ?? throw new ArgumentNullException(nameof(selection));
        if (firstRow < 1) throw new ArgumentOutOfRangeException(nameof(firstRow));
        if (firstColumn < 1) throw new ArgumentOutOfRangeException(nameof(firstColumn));
        Contents = contents ?? throw new ArgumentNullException(nameof(contents));
        if (selection.CellCount != contents.CellCount) throw new ArgumentException("Selection and content cell counts must match.", nameof(contents));
        FirstRow = firstRow;
        FirstColumn = firstColumn;
    }
    public SelectionSnapshot Selection { get; }
    public int FirstRow { get; }
    public int FirstColumn { get; }
    public FormulaCellBlock Contents { get; }
}
