using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelAccel.Core.Formulas;

namespace ExcelAccel.Core.Auditing;

/// <summary>
/// An A1 rectangle in one worksheet, normalized so the first endpoint is always
/// the top-left corner. Coordinates are one-based, matching Excel.
/// </summary>
public readonly struct AuditRectangle : IEquatable<AuditRectangle>
{
    public AuditRectangle(int firstRow, int firstColumn, int lastRow, int lastColumn)
    {
        if (firstRow < 1 || lastRow < 1 || firstColumn < 1 || lastColumn < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(firstRow), "Rectangle coordinates are one-based.");
        }

        FirstRow = Math.Min(firstRow, lastRow);
        LastRow = Math.Max(firstRow, lastRow);
        FirstColumn = Math.Min(firstColumn, lastColumn);
        LastColumn = Math.Max(firstColumn, lastColumn);
    }

    public int FirstRow { get; }

    public int FirstColumn { get; }

    public int LastRow { get; }

    public int LastColumn { get; }

    public bool IsSingleCell => FirstRow == LastRow && FirstColumn == LastColumn;

    /// <summary>
    /// Direct dependence is an intersection relation, not containment: a formula
    /// that reads any part of the target reads the target.
    /// </summary>
    public bool Intersects(AuditRectangle other) =>
        FirstRow <= other.LastRow && other.FirstRow <= LastRow &&
        FirstColumn <= other.LastColumn && other.FirstColumn <= LastColumn;

    public bool Equals(AuditRectangle other) =>
        FirstRow == other.FirstRow && LastRow == other.LastRow &&
        FirstColumn == other.FirstColumn && LastColumn == other.LastColumn;

    public override bool Equals(object? obj) => obj is AuditRectangle other && Equals(other);

    public override int GetHashCode()
    {
        unchecked
        {
            var hash = FirstRow;
            hash = (hash * 397) ^ FirstColumn;
            hash = (hash * 397) ^ LastRow;
            return (hash * 397) ^ LastColumn;
        }
    }

    public override string ToString() => IsSingleCell
        ? AuditAddress.Cell(FirstRow, FirstColumn)
        : AuditAddress.Cell(FirstRow, FirstColumn) + ":" + AuditAddress.Cell(LastRow, LastColumn);
}

/// <summary>
/// Invariant A1 address formatting and parsing shared by the auditing model.
/// Only the qualified A1 subset is accepted; anything else fails closed so a
/// caller must represent it as an explicit coverage gap.
/// </summary>
public static class AuditAddress
{
    public const int MaximumRow = 1_048_576;
    public const int MaximumColumn = 16_384;

    public static string Cell(int row, int column)
    {
        if (row < 1 || row > MaximumRow) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 1 || column > MaximumColumn) throw new ArgumentOutOfRangeException(nameof(column));
        var name = string.Empty;
        for (var remaining = column; remaining > 0; remaining /= 26)
        {
            var zeroBased = (remaining - 1) % 26;
            name = (char)('A' + zeroBased) + name;
            remaining -= zeroBased;
        }

        return name + row.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Parses "A1" or "A1:B5", with or without absolute markers. Whole-column,
    /// whole-row, multi-area, and sheet-qualified forms are deliberately
    /// rejected; the caller must treat them as an explicit coverage gap.
    /// </summary>
    public static bool TryParse(string? address, out AuditRectangle rectangle)
    {
        rectangle = default;
        if (string.IsNullOrWhiteSpace(address)) return false;
        var text = address!.Trim();
        var separator = text.IndexOf(':');
        if (separator < 0)
        {
            if (!TryParseEndpoint(text, out var row, out var column)) return false;
            rectangle = new AuditRectangle(row, column, row, column);
            return true;
        }

        if (text.IndexOf(':', separator + 1) >= 0) return false;
        if (!TryParseEndpoint(text.Substring(0, separator), out var firstRow, out var firstColumn)) return false;
        if (!TryParseEndpoint(text.Substring(separator + 1), out var lastRow, out var lastColumn)) return false;
        rectangle = new AuditRectangle(firstRow, firstColumn, lastRow, lastColumn);
        return true;
    }

    /// <summary>
    /// Converts a parsed A1 reference to its rectangle. Coordinates the parser
    /// could not pin to a concrete row and column fail closed.
    /// </summary>
    public static bool TryResolve(FormulaReference reference, out AuditRectangle rectangle)
    {
        rectangle = default;
        if (reference is null) throw new ArgumentNullException(nameof(reference));
        if (reference.Notation != FormulaNotation.A1) return false;
        if (reference.HasSpillOperator || reference.HasImplicitIntersection) return false;
        if (!TryEndpoint(reference.First, out var firstRow, out var firstColumn)) return false;
        if (reference.Second is null)
        {
            rectangle = new AuditRectangle(firstRow, firstColumn, firstRow, firstColumn);
            return true;
        }

        if (!TryEndpoint(reference.Second, out var lastRow, out var lastColumn)) return false;
        rectangle = new AuditRectangle(firstRow, firstColumn, lastRow, lastColumn);
        return true;
    }

    private static bool TryEndpoint(FormulaReferenceEndpoint endpoint, out int row, out int column)
    {
        row = endpoint.Row.Value;
        column = endpoint.Column.Value;
        if (endpoint.Row.Kind == FormulaCoordinateKind.Current || endpoint.Column.Kind == FormulaCoordinateKind.Current) return false;
        return row >= 1 && row <= MaximumRow && column >= 1 && column <= MaximumColumn;
    }

    private static bool TryParseEndpoint(string text, out int row, out int column)
    {
        row = 0;
        column = 0;
        if (string.IsNullOrWhiteSpace(text)) return false;
        var index = 0;
        if (text[index] == '$') index++;
        var columnStart = index;
        while (index < text.Length && IsLetter(text[index])) index++;
        if (index == columnStart) return false;
        var columnText = text.Substring(columnStart, index - columnStart);
        if (columnText.Length > 3) return false;
        var columnValue = 0;
        foreach (var character in columnText)
        {
            columnValue = (columnValue * 26) + (char.ToUpperInvariant(character) - 'A' + 1);
        }

        if (columnValue < 1 || columnValue > MaximumColumn) return false;
        if (index < text.Length && text[index] == '$') index++;
        var rowStart = index;
        while (index < text.Length && text[index] >= '0' && text[index] <= '9') index++;
        if (index != text.Length || index == rowStart) return false;
        if (!int.TryParse(text.Substring(rowStart, index - rowStart), NumberStyles.None, CultureInfo.InvariantCulture, out var rowValue)) return false;
        if (rowValue < 1 || rowValue > MaximumRow) return false;
        row = rowValue;
        column = columnValue;
        return true;
    }

    private static bool IsLetter(char value) => (value >= 'A' && value <= 'Z') || (value >= 'a' && value <= 'z');
}

/// <summary>
/// The single definition of "this identifier token is a defined-name reference",
/// shared so precedent and dependent analysis cannot drift apart.
/// </summary>
public static class AuditNameCandidates
{
    public static bool IsNameCandidate(IReadOnlyList<FormulaToken> tokens, int index)
    {
        if (tokens is null) throw new ArgumentNullException(nameof(tokens));
        var token = tokens[index];
        if (token.Kind != FormulaTokenKind.Identifier || IsBoolean(token.Text)) return false;
        var next = NextSignificant(tokens, index + 1);
        if (next?.Kind == FormulaTokenKind.OpenParenthesis) return false;
        return index + 1 >= tokens.Count || tokens[index + 1].Kind != FormulaTokenKind.BracketedIdentifier;
    }

    private static FormulaToken? NextSignificant(IReadOnlyList<FormulaToken> tokens, int start)
    {
        for (var index = start; index < tokens.Count; index++)
        {
            if (tokens[index].Kind != FormulaTokenKind.Whitespace) return tokens[index];
        }

        return null;
    }

    private static bool IsBoolean(string value) =>
        string.Equals(value, "TRUE", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "FALSE", StringComparison.OrdinalIgnoreCase);
}
