using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.Formulas;

namespace ExcelAccel.Core.ModelCheck;

/// <summary>
/// The normalized relative structure of a formula: the form two copied cells
/// share. Every reference is rewritten as an offset from its own cell, so
/// `=A1*2` in B1 and `=A2*2` in B2 normalize to the same shape while `=A1*3`
/// does not.
///
/// Normalization is purely structural. It never evaluates, and a formula the
/// qualified parser cannot cover has no shape at all rather than a guessed one.
/// </summary>
public static class FormulaShape
{
    /// <summary>
    /// Returns the normalized shape, or null when the formula is outside
    /// qualified parser coverage. A null shape is a coverage gap, never a
    /// mismatch.
    /// </summary>
    public static string? TryNormalize(string? formula, AuditCellIdentity origin)
    {
        if (origin is null) throw new ArgumentNullException(nameof(origin));
        if (string.IsNullOrEmpty(formula) || formula![0] != '=') return null;
        if (!AuditAddress.TryParse(origin.Address, out var anchor)) return null;

        var parse = new FormulaParser().Parse(formula, new FormulaParseOptions(FormulaDialect.InvariantA1));
        if (!parse.IsSuccess) return null;
        var document = parse.Document!;
        if (document.LimitationCodes.Any(code => code != FormulaRefusalCodes.ExternalReferenceInspectOnly)) return null;

        var builder = new StringBuilder(formula);
        foreach (var reference in document.References.OrderByDescending(value => value.Span.Start))
        {
            var normalized = Normalize(reference, anchor.FirstRow, anchor.FirstColumn);
            if (normalized is null) return null;
            builder.Remove(reference.Span.Start, reference.Span.Length);
            builder.Insert(reference.Span.Start, normalized);
        }

        return builder.ToString();
    }

    private static string? Normalize(FormulaReference reference, int anchorRow, int anchorColumn)
    {
        var qualifier = reference.Qualifier is null ? string.Empty : reference.Qualifier + "!";
        var first = NormalizeEndpoint(reference.First, anchorRow, anchorColumn);
        if (first is null) return null;
        if (reference.Second is null) return qualifier + first;
        var second = NormalizeEndpoint(reference.Second, anchorRow, anchorColumn);
        return second is null ? null : qualifier + first + ":" + second;
    }

    private static string? NormalizeEndpoint(FormulaReferenceEndpoint endpoint, int anchorRow, int anchorColumn)
    {
        var row = NormalizeCoordinate(endpoint.Row, anchorRow, 'R');
        var column = NormalizeCoordinate(endpoint.Column, anchorColumn, 'C');
        return row is null || column is null ? null : row + column;
    }

    private static string? NormalizeCoordinate(FormulaReferenceCoordinate coordinate, int anchor, char prefix)
    {
        switch (coordinate.Kind)
        {
            case FormulaCoordinateKind.Absolute:
                return prefix + "$" + coordinate.Value.ToString(CultureInfo.InvariantCulture);
            case FormulaCoordinateKind.Relative:
                var offset = coordinate.Value - anchor;
                return prefix + "[" + offset.ToString(CultureInfo.InvariantCulture) + "]";
            default:
                return null;
        }
    }

    /// <summary>
    /// The numeric literals a formula embeds, with their exact source spans and
    /// the function that encloses each one.
    /// </summary>
    public static IReadOnlyList<EmbeddedLiteral> ReadEmbeddedLiterals(string? formula)
    {
        if (string.IsNullOrEmpty(formula) || formula![0] != '=') return Array.Empty<EmbeddedLiteral>();
        var parse = new FormulaParser().Parse(formula, new FormulaParseOptions(FormulaDialect.InvariantA1));
        if (!parse.IsSuccess) return Array.Empty<EmbeddedLiteral>();

        var literals = new List<EmbeddedLiteral>();
        var enclosing = new Stack<string>();
        var tokens = parse.Document!.Tokens;
        for (var index = 0; index < tokens.Count; index++)
        {
            var token = tokens[index];
            switch (token.Kind)
            {
                case FormulaTokenKind.OpenParenthesis:
                    enclosing.Push(PrecedingFunctionName(tokens, index));
                    break;
                case FormulaTokenKind.CloseParenthesis:
                    if (enclosing.Count != 0) enclosing.Pop();
                    break;
                case FormulaTokenKind.Number:
                    if (double.TryParse(token.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
                    {
                        literals.Add(new EmbeddedLiteral(value, token.Text, token.Span, enclosing.Count == 0 ? string.Empty : enclosing.Peek()));
                    }

                    break;
            }
        }

        return literals;
    }

    private static string PrecedingFunctionName(IReadOnlyList<FormulaToken> tokens, int parenthesisIndex)
    {
        for (var index = parenthesisIndex - 1; index >= 0; index--)
        {
            if (tokens[index].Kind == FormulaTokenKind.Whitespace) continue;
            return tokens[index].Kind == FormulaTokenKind.Identifier ? tokens[index].Text.ToUpperInvariant() : string.Empty;
        }

        return string.Empty;
    }
}

public sealed class EmbeddedLiteral
{
    internal EmbeddedLiteral(double value, string sourceText, FormulaSourceSpan span, string enclosingFunction)
    {
        Value = value;
        SourceText = sourceText;
        Span = span;
        EnclosingFunction = enclosingFunction;
    }

    public double Value { get; }

    public string SourceText { get; }

    public FormulaSourceSpan Span { get; }

    /// <summary>The function directly enclosing the literal, or empty at top level.</summary>
    public string EnclosingFunction { get; }
}
