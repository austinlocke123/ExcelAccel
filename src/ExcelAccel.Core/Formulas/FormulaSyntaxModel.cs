using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAccel.Core.Formulas;

public readonly struct FormulaSourceSpan
{
    public FormulaSourceSpan(int start, int length)
    {
        if (start < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(start));
        }

        if (length < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(length));
        }

        Start = start;
        Length = length;
    }

    public int Start { get; }

    public int Length { get; }

    public int End => checked(Start + Length);

    public bool Contains(FormulaSourceSpan other) => Start <= other.Start && End >= other.End;
}

public enum FormulaTokenKind
{
    Prefix,
    Whitespace,
    Reference,
    Identifier,
    Number,
    ErrorLiteral,
    StringLiteral,
    QuotedIdentifier,
    BracketedIdentifier,
    Operator,
    OpenParenthesis,
    CloseParenthesis,
    Separator
}

public sealed class FormulaToken
{
    public FormulaToken(FormulaTokenKind kind, string text, FormulaSourceSpan span)
    {
        Kind = kind;
        Text = text ?? throw new ArgumentNullException(nameof(text));
        Span = span;
    }

    public FormulaTokenKind Kind { get; }

    public string Text { get; }

    public FormulaSourceSpan Span { get; }
}

public enum FormulaCoordinateKind
{
    Current,
    Relative,
    Absolute
}

public readonly struct FormulaReferenceCoordinate
{
    public FormulaReferenceCoordinate(FormulaCoordinateKind kind, int value)
    {
        Kind = kind;
        Value = value;
    }

    public FormulaCoordinateKind Kind { get; }

    public int Value { get; }
}

public sealed class FormulaReferenceEndpoint
{
    public FormulaReferenceEndpoint(
        string sourceText,
        FormulaReferenceCoordinate row,
        FormulaReferenceCoordinate column)
    {
        SourceText = sourceText ?? throw new ArgumentNullException(nameof(sourceText));
        Row = row;
        Column = column;
    }

    public string SourceText { get; }

    public FormulaReferenceCoordinate Row { get; }

    public FormulaReferenceCoordinate Column { get; }
}

public sealed class FormulaReference
{
    public FormulaReference(
        string sourceText,
        FormulaSourceSpan span,
        FormulaNotation notation,
        string? qualifier,
        FormulaReferenceEndpoint first,
        FormulaReferenceEndpoint? second,
        bool hasImplicitIntersection,
        bool hasSpillOperator)
    {
        SourceText = sourceText ?? throw new ArgumentNullException(nameof(sourceText));
        Span = span;
        Notation = notation;
        Qualifier = qualifier;
        First = first ?? throw new ArgumentNullException(nameof(first));
        Second = second;
        HasImplicitIntersection = hasImplicitIntersection;
        HasSpillOperator = hasSpillOperator;
    }

    public string SourceText { get; }

    public FormulaSourceSpan Span { get; }

    public FormulaNotation Notation { get; }

    public string? Qualifier { get; }

    public FormulaReferenceEndpoint First { get; }

    public FormulaReferenceEndpoint? Second { get; }

    public bool IsRange => Second is not null;

    public bool HasImplicitIntersection { get; }

    public bool HasSpillOperator { get; }
}

public sealed class FormulaSyntaxDocument
{
    public FormulaSyntaxDocument(
        string sourceText,
        FormulaDialect dialect,
        IReadOnlyList<FormulaToken> tokens,
        IReadOnlyList<FormulaReference> references,
        FormulaCoverageDisposition disposition,
        string? limitationCode,
        IReadOnlyList<string>? limitationCodes = null)
    {
        SourceText = sourceText ?? throw new ArgumentNullException(nameof(sourceText));
        Dialect = dialect ?? throw new ArgumentNullException(nameof(dialect));
        if (tokens is null)
        {
            throw new ArgumentNullException(nameof(tokens));
        }

        if (references is null)
        {
            throw new ArgumentNullException(nameof(references));
        }

        if (disposition == FormulaCoverageDisposition.Refuse)
        {
            throw new ArgumentException("A refused parse cannot expose a syntax document.", nameof(disposition));
        }

        if (disposition == FormulaCoverageDisposition.InspectOnly && string.IsNullOrWhiteSpace(limitationCode))
        {
            throw new ArgumentException("Inspect-only syntax requires a stable limitation code.", nameof(limitationCode));
        }

        if (disposition != FormulaCoverageDisposition.InspectOnly && limitationCode is not null)
        {
            throw new ArgumentException("Only inspect-only syntax can carry a limitation code.", nameof(limitationCode));
        }

        Tokens = Array.AsReadOnly(tokens.ToArray());
        References = Array.AsReadOnly(references.ToArray());
        Disposition = disposition;
        LimitationCode = limitationCode;
        LimitationCodes = limitationCodes is null
            ? Array.AsReadOnly(limitationCode is null ? new string[0] : new[] { limitationCode })
            : Array.AsReadOnly(limitationCodes.ToArray());
        if (limitationCodes is not null && limitationCode is not null &&
            (LimitationCodes.Count == 0 || LimitationCodes[0] != limitationCode))
        {
            throw new ArgumentException("The first coverage cause must be the reported limitation code.", nameof(limitationCodes));
        }
    }

    public string SourceText { get; }

    public FormulaDialect Dialect { get; }

    public IReadOnlyList<FormulaToken> Tokens { get; }

    public IReadOnlyList<FormulaReference> References { get; }

    public FormulaCoverageDisposition Disposition { get; }

    /// <summary>
    /// The first coverage limitation, in the parser's fixed precedence order.
    /// Kept for callers that only need to name one reason.
    /// </summary>
    public string? LimitationCode { get; }

    /// <summary>
    /// Every coverage limitation this formula carries, in the same precedence
    /// order, with the reported <see cref="LimitationCode"/> first. A caller that
    /// must decide whether a specific construct could hide something needs all of
    /// them, because the first cause alone does not rule the others out.
    /// </summary>
    public IReadOnlyList<string> LimitationCodes { get; }

    public string Serialize() => SourceText;
}

public sealed class FormulaParseResult
{
    private FormulaParseResult(FormulaSyntaxDocument? document, string? refusalCode, string? message, FormulaSourceSpan? failureSpan)
    {
        Document = document;
        RefusalCode = refusalCode;
        Message = message;
        FailureSpan = failureSpan;
    }

    public bool IsSuccess => Document is not null;

    public FormulaSyntaxDocument? Document { get; }

    public string? RefusalCode { get; }

    public string? Message { get; }

    public FormulaSourceSpan? FailureSpan { get; }

    public static FormulaParseResult Success(FormulaSyntaxDocument document) =>
        new FormulaParseResult(document ?? throw new ArgumentNullException(nameof(document)), null, null, null);

    public static FormulaParseResult Refused(string code, string message, FormulaSourceSpan? span = null) =>
        new FormulaParseResult(
            null,
            !string.IsNullOrWhiteSpace(code) ? code : throw new ArgumentException("A stable refusal code is required.", nameof(code)),
            !string.IsNullOrWhiteSpace(message) ? message : throw new ArgumentException("A refusal message is required.", nameof(message)),
            span);
}
