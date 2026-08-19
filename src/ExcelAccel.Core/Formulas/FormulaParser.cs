using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelAccel.Core.Formulas;

public sealed class FormulaParser
{
    private const int ExcelMaximumColumn = 16384;
    private const int ExcelMaximumRow = 1048576;

    private static readonly Regex A1ReferencePattern = new Regex(
        @"\G(?<implicit>@)?(?:(?<qualifier>'(?:[^']|'')+'|(?:\[[^\]\r\n]+\])?[A-Za-z_\\][A-Za-z0-9_.]*)!)?(?<first>\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})(?::(?<second>\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}))?(?<spill>#)?",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static readonly Regex R1C1ReferencePattern = new Regex(
        @"\G(?<implicit>@)?(?:(?<qualifier>'(?:[^']|'')+'|(?:\[[^\]\r\n]+\])?[A-Za-z_\\][A-Za-z0-9_.]*)!)?(?<first>R(?:\[-?[0-9]+\]|[0-9]+)?C(?:\[-?[0-9]+\]|[0-9]+)?)(?::(?<second>R(?:\[-?[0-9]+\]|[0-9]+)?C(?:\[-?[0-9]+\]|[0-9]+)?))?(?<spill>#)?",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static readonly Regex ErrorLiteralPattern = new Regex(
        @"\G#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|GETTING_DATA|SPILL!|CALC!|FIELD!|BLOCKED!|UNKNOWN!|CONNECT!|BUSY!)",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static readonly Regex R1C1EndpointPattern = new Regex(
        @"^R(?<row>\[-?[0-9]+\]|[0-9]+)?C(?<column>\[-?[0-9]+\]|[0-9]+)?$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    public FormulaParseResult Parse(string? formula, FormulaParseOptions? options = null)
    {
        options ??= FormulaParseOptions.DefaultA1;
        if (formula is null || formula.Length == 0)
        {
            return FormulaParseResult.Refused(FormulaRefusalCodes.Empty, "A formula is required.");
        }

        if (formula[0] != '=')
        {
            return FormulaParseResult.Refused(
                FormulaRefusalCodes.PrefixRequired,
                "The formula must begin with '='.",
                new FormulaSourceSpan(0, 1));
        }

        if (formula.Length > options.MaximumLength)
        {
            return FormulaParseResult.Refused(
                FormulaRefusalCodes.TooLong,
                $"The formula exceeds the {options.MaximumLength.ToString(CultureInfo.InvariantCulture)} character limit.");
        }

        var tokens = new List<FormulaToken>();
        var references = new List<FormulaReference>();
        tokens.Add(new FormulaToken(FormulaTokenKind.Prefix, "=", new FormulaSourceSpan(0, 1)));

        var depth = 0;
        var index = 1;
        while (index < formula.Length)
        {
            var start = index;
            var character = formula[index];
            FormulaTokenKind kind;

            if (char.IsWhiteSpace(character))
            {
                index++;
                while (index < formula.Length && char.IsWhiteSpace(formula[index]))
                {
                    index++;
                }

                kind = FormulaTokenKind.Whitespace;
            }
            else if (character == '"')
            {
                if (!TryReadDelimited(formula, ref index, '"'))
                {
                    return FormulaParseResult.Refused(
                        FormulaRefusalCodes.UnterminatedString,
                        "The formula contains an unterminated string literal.",
                        new FormulaSourceSpan(start, formula.Length - start));
                }

                kind = FormulaTokenKind.StringLiteral;
            }
            else
            {
                var referenceOutcome = TryReadReference(formula, options.Dialect.Notation, ref index, out var reference);
                if (referenceOutcome == ReferenceReadOutcome.Invalid)
                {
                    return FormulaParseResult.Refused(
                        FormulaRefusalCodes.InvalidReference,
                        "The formula contains a reference outside Excel's qualified bounds.",
                        new FormulaSourceSpan(start, Math.Max(1, index - start)));
                }

                if (referenceOutcome == ReferenceReadOutcome.Success)
                {
                    references.Add(reference!);
                    kind = FormulaTokenKind.Reference;
                }
                else if (character == '#' && TryReadErrorLiteral(formula, ref index))
                {
                    kind = FormulaTokenKind.ErrorLiteral;
                }
                else if (character == '\'')
                {
                    if (!TryReadDelimited(formula, ref index, '\''))
                    {
                        return FormulaParseResult.Refused(
                            FormulaRefusalCodes.UnterminatedString,
                            "The formula contains an unterminated quoted identifier.",
                            new FormulaSourceSpan(start, formula.Length - start));
                    }

                    kind = FormulaTokenKind.QuotedIdentifier;
                }
                else if (character == '[')
                {
                    if (!TryReadBracketed(formula, ref index))
                    {
                        return FormulaParseResult.Refused(
                            FormulaRefusalCodes.UnbalancedDelimiter,
                            "The formula contains an unclosed bracket.",
                            new FormulaSourceSpan(start, formula.Length - start));
                    }

                    kind = FormulaTokenKind.BracketedIdentifier;
                }
                else if (character == '{' || character == '}')
                {
                    return FormulaParseResult.Refused(
                        FormulaRefusalCodes.UnsupportedArraySyntax,
                        "Array constants and legacy array wrappers are outside the v1 parser transform set.",
                        new FormulaSourceSpan(start, 1));
                }
                else if (character == '(')
                {
                    depth++;
                    if (depth > options.MaximumNesting)
                    {
                        return FormulaParseResult.Refused(
                            FormulaRefusalCodes.NestingLimit,
                            $"Formula nesting exceeds the {options.MaximumNesting.ToString(CultureInfo.InvariantCulture)} level limit.",
                            new FormulaSourceSpan(start, 1));
                    }

                    index++;
                    kind = FormulaTokenKind.OpenParenthesis;
                }
                else if (character == ')')
                {
                    if (depth == 0)
                    {
                        return FormulaParseResult.Refused(
                            FormulaRefusalCodes.UnbalancedDelimiter,
                            "The formula contains an unmatched closing parenthesis.",
                            new FormulaSourceSpan(start, 1));
                    }

                    depth--;
                    index++;
                    kind = FormulaTokenKind.CloseParenthesis;
                }
                else if (character == options.Dialect.ListSeparator)
                {
                    index++;
                    kind = FormulaTokenKind.Separator;
                }
                else if ((character == ',' || character == ';') && character != options.Dialect.DecimalSeparator)
                {
                    return FormulaParseResult.Refused(
                        FormulaRefusalCodes.DialectMismatch,
                        $"The formula uses '{character}' where the selected dialect requires '{options.Dialect.ListSeparator}'.",
                        new FormulaSourceSpan(start, 1));
                }
                else if (char.IsDigit(character) ||
                         (character == options.Dialect.DecimalSeparator && index + 1 < formula.Length && char.IsDigit(formula[index + 1])))
                {
                    ReadNumber(formula, options.Dialect.DecimalSeparator, ref index);
                    kind = FormulaTokenKind.Number;
                }
                else if (IsIdentifierStart(character))
                {
                    index++;
                    while (index < formula.Length && IsIdentifierPart(formula[index]))
                    {
                        index++;
                    }

                    kind = FormulaTokenKind.Identifier;
                }
                else if (IsOperator(character))
                {
                    index++;
                    if (index < formula.Length && IsTwoCharacterOperator(character, formula[index]))
                    {
                        index++;
                    }

                    kind = FormulaTokenKind.Operator;
                }
                else
                {
                    return FormulaParseResult.Refused(
                        FormulaRefusalCodes.UnsupportedCharacter,
                        $"The character U+{((int)character).ToString("X4", CultureInfo.InvariantCulture)} is not in the v1 syntax set.",
                        new FormulaSourceSpan(start, 1));
                }
            }

            tokens.Add(new FormulaToken(kind, formula.Substring(start, index - start), new FormulaSourceSpan(start, index - start)));
            if (tokens.Count > options.MaximumTokens)
            {
                return FormulaParseResult.Refused(
                    FormulaRefusalCodes.TooManyTokens,
                    $"The formula exceeds the {options.MaximumTokens.ToString(CultureInfo.InvariantCulture)} token limit.");
            }
        }

        if (depth != 0)
        {
            return FormulaParseResult.Refused(
                FormulaRefusalCodes.UnbalancedDelimiter,
                "The formula contains an unclosed parenthesis.");
        }

        if (!TryValidateTokenSequence(tokens, out var invalidSpan))
        {
            return FormulaParseResult.Refused(
                FormulaRefusalCodes.InvalidSyntax,
                "The token sequence is outside the conservatively validated v1 expression grammar.",
                invalidSpan);
        }

        var coverage = ClassifyCoverage(formula, options.Dialect.Notation, tokens, references);
        var document = new FormulaSyntaxDocument(
            formula,
            options.Dialect,
            Array.AsReadOnly(tokens.ToArray()),
            Array.AsReadOnly(references.ToArray()),
            coverage.Disposition,
            coverage.LimitationCode);
        return FormulaParseResult.Success(document);
    }

    private static CoverageClassification ClassifyCoverage(
        string formula,
        FormulaNotation notation,
        IReadOnlyList<FormulaToken> tokens,
        IReadOnlyList<FormulaReference> references)
    {
        if (notation == FormulaNotation.A1 && ContainsStructuredReference(tokens))
        {
            return CoverageClassification.InspectOnly(FormulaRefusalCodes.StructuredReferenceInspectOnly);
        }

        if (references.Any(reference =>
                reference.HasImplicitIntersection || reference.HasSpillOperator))
        {
            return CoverageClassification.InspectOnly(FormulaRefusalCodes.DynamicArrayInspectOnly);
        }

        if (references.Any(reference =>
                reference.Qualifier is not null && reference.Qualifier.IndexOf('[') >= 0))
        {
            return CoverageClassification.InspectOnly(FormulaRefusalCodes.ExternalReferenceInspectOnly);
        }

        for (var index = 0; index < tokens.Count; index++)
        {
            if (tokens[index].Kind == FormulaTokenKind.BracketedIdentifier &&
                (index == 0 || tokens[index - 1].Kind != FormulaTokenKind.Identifier))
            {
                return CoverageClassification.InspectOnly(FormulaRefusalCodes.ExternalReferenceInspectOnly);
            }
        }

        for (var index = 0; index + 2 < tokens.Count; index++)
        {
            if (tokens[index].Kind == FormulaTokenKind.Reference &&
                tokens[index + 1].Kind == FormulaTokenKind.Whitespace &&
                tokens[index + 2].Kind == FormulaTokenKind.Reference)
            {
                return CoverageClassification.InspectOnly(FormulaRefusalCodes.IntersectionInspectOnly);
            }
        }


        var parenthesisDepth = 0;
        for (var index = 0; index < tokens.Count; index++)
        {
            if (tokens[index].Kind == FormulaTokenKind.OpenParenthesis)
            {
                parenthesisDepth++;
            }
            else if (tokens[index].Kind == FormulaTokenKind.CloseParenthesis)
            {
                parenthesisDepth--;
            }
            else if (tokens[index].Kind == FormulaTokenKind.Separator && parenthesisDepth == 0)
            {
                return CoverageClassification.InspectOnly(FormulaRefusalCodes.UnionInspectOnly);
            }
        }

        for (var index = 0; index < tokens.Count; index++)
        {
            if (tokens[index].Kind != FormulaTokenKind.Identifier)
            {
                continue;
            }

            var identifier = tokens[index].Text;
            if (identifier.Equals("TRUE", StringComparison.OrdinalIgnoreCase) ||
                identifier.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var next = NextSignificantToken(tokens, index + 1);
            if (next is not null && next.Kind == FormulaTokenKind.OpenParenthesis)
            {
                continue;
            }

            return CoverageClassification.InspectOnly(FormulaRefusalCodes.NameInspectOnly);
        }

        return references.Count == 0
            ? CoverageClassification.RoundTrip()
            : CoverageClassification.Transform();
    }

    private static bool ContainsStructuredReference(IReadOnlyList<FormulaToken> tokens)
    {
        for (var index = 0; index + 1 < tokens.Count; index++)
        {
            if (tokens[index].Kind == FormulaTokenKind.Identifier &&
                tokens[index + 1].Kind == FormulaTokenKind.BracketedIdentifier)
            {
                return true;
            }
        }

        return false;
    }

    private static bool TryValidateTokenSequence(
        IReadOnlyList<FormulaToken> tokens,
        out FormulaSourceSpan? invalidSpan)
    {
        var expectOperand = true;
        var parenthesisDepth = 0;
        FormulaToken? previous = null;
        var whitespaceSincePrevious = false;

        for (var index = 1; index < tokens.Count; index++)
        {
            var token = tokens[index];
            if (token.Kind == FormulaTokenKind.Whitespace)
            {
                whitespaceSincePrevious = true;
                continue;
            }

            switch (token.Kind)
            {
                case FormulaTokenKind.Reference:
                case FormulaTokenKind.Number:
                case FormulaTokenKind.ErrorLiteral:
                case FormulaTokenKind.StringLiteral:
                case FormulaTokenKind.QuotedIdentifier:
                    if (!expectOperand)
                    {
                        if (token.Kind == FormulaTokenKind.Reference &&
                            previous?.Kind == FormulaTokenKind.Reference &&
                            whitespaceSincePrevious)
                        {
                            break;
                        }

                        invalidSpan = token.Span;
                        return false;
                    }

                    expectOperand = false;
                    break;

                case FormulaTokenKind.Identifier:
                    if (!expectOperand)
                    {
                        if (previous?.Kind == FormulaTokenKind.BracketedIdentifier && !whitespaceSincePrevious)
                        {
                            break;
                        }

                        invalidSpan = token.Span;
                        return false;
                    }

                    if (NextSignificantToken(tokens, index + 1)?.Kind != FormulaTokenKind.OpenParenthesis)
                    {
                        expectOperand = false;
                    }

                    break;

                case FormulaTokenKind.BracketedIdentifier:
                    if (!expectOperand && (previous?.Kind != FormulaTokenKind.Identifier || whitespaceSincePrevious))
                    {
                        invalidSpan = token.Span;
                        return false;
                    }

                    break;

                case FormulaTokenKind.OpenParenthesis:
                    if (!expectOperand)
                    {
                        invalidSpan = token.Span;
                        return false;
                    }

                    parenthesisDepth++;
                    expectOperand = true;
                    break;

                case FormulaTokenKind.CloseParenthesis:
                    if (parenthesisDepth == 0 ||
                        (expectOperand && previous?.Kind != FormulaTokenKind.OpenParenthesis))
                    {
                        invalidSpan = token.Span;
                        return false;
                    }

                    parenthesisDepth--;
                    expectOperand = false;
                    break;

                case FormulaTokenKind.Separator:
                    if (expectOperand)
                    {
                        invalidSpan = token.Span;
                        return false;
                    }

                    if (parenthesisDepth == 0 &&
                        (previous?.Kind != FormulaTokenKind.Reference ||
                         NextSignificantToken(tokens, index + 1)?.Kind != FormulaTokenKind.Reference))
                    {
                        invalidSpan = token.Span;
                        return false;
                    }

                    expectOperand = true;
                    break;

                case FormulaTokenKind.Operator:
                    if (token.Text == "%" && !expectOperand)
                    {
                        break;
                    }

                    if (expectOperand && (token.Text == "+" || token.Text == "-" || token.Text == "@"))
                    {
                        break;
                    }

                    if (token.Text == "!" &&
                        !expectOperand &&
                        (previous?.Kind == FormulaTokenKind.Identifier ||
                         previous?.Kind == FormulaTokenKind.QuotedIdentifier) &&
                        NextSignificantToken(tokens, index + 1)?.Kind == FormulaTokenKind.Identifier)
                    {
                        expectOperand = true;
                        break;
                    }

                    if (expectOperand || token.Text == "!" || token.Text == ":" || token.Text == "#" || token.Text == "@")
                    {
                        invalidSpan = token.Span;
                        return false;
                    }

                    expectOperand = true;
                    break;

                default:
                    invalidSpan = token.Span;
                    return false;
            }

            previous = token;
            whitespaceSincePrevious = false;
        }

        if (expectOperand || parenthesisDepth != 0)
        {
            invalidSpan = previous?.Span;
            return false;
        }

        invalidSpan = null;
        return true;
    }

    private static FormulaToken? NextSignificantToken(IReadOnlyList<FormulaToken> tokens, int start)
    {
        for (var index = start; index < tokens.Count; index++)
        {
            if (tokens[index].Kind != FormulaTokenKind.Whitespace)
            {
                return tokens[index];
            }
        }

        return null;
    }

    private static ReferenceReadOutcome TryReadReference(
        string formula,
        FormulaNotation notation,
        ref int index,
        out FormulaReference? reference)
    {
        reference = null;
        var pattern = notation == FormulaNotation.A1 ? A1ReferencePattern : R1C1ReferencePattern;
        var match = pattern.Match(formula, index);
        if (!match.Success || match.Index != index)
        {
            return ReferenceReadOutcome.NoMatch;
        }

        var end = match.Index + match.Length;
        if (end < formula.Length && IsIdentifierPart(formula[end]))
        {
            return ReferenceReadOutcome.NoMatch;
        }

        if (end < formula.Length && formula[end] == '(' &&
            !match.Groups["qualifier"].Success &&
            !match.Groups["second"].Success &&
            !match.Groups["spill"].Success &&
            !match.Groups["implicit"].Success)
        {
            return ReferenceReadOutcome.NoMatch;
        }

        if (!TryParseEndpoint(match.Groups["first"].Value, notation, out var first) ||
            (match.Groups["second"].Success && !TryParseEndpoint(match.Groups["second"].Value, notation, out _)))
        {
            index = end;
            return ReferenceReadOutcome.Invalid;
        }

        FormulaReferenceEndpoint? second = null;
        if (match.Groups["second"].Success)
        {
            TryParseEndpoint(match.Groups["second"].Value, notation, out second);
        }

        reference = new FormulaReference(
            match.Value,
            new FormulaSourceSpan(index, match.Length),
            notation,
            match.Groups["qualifier"].Success ? match.Groups["qualifier"].Value : null,
            first!,
            second,
            match.Groups["implicit"].Success,
            match.Groups["spill"].Success);
        index = end;
        return ReferenceReadOutcome.Success;
    }

    private static bool TryParseEndpoint(
        string source,
        FormulaNotation notation,
        out FormulaReferenceEndpoint? endpoint)
    {
        return notation == FormulaNotation.A1
            ? TryParseA1Endpoint(source, out endpoint)
            : TryParseR1C1Endpoint(source, out endpoint);
    }

    private static bool TryParseA1Endpoint(string source, out FormulaReferenceEndpoint? endpoint)
    {
        endpoint = null;
        var index = 0;
        var columnAbsolute = source[index] == '$';
        if (columnAbsolute)
        {
            index++;
        }

        var columnStart = index;
        while (index < source.Length && char.IsLetter(source[index]))
        {
            index++;
        }

        var columnText = source.Substring(columnStart, index - columnStart);
        var rowAbsolute = index < source.Length && source[index] == '$';
        if (rowAbsolute)
        {
            index++;
        }

        if (!int.TryParse(source.Substring(index), NumberStyles.None, CultureInfo.InvariantCulture, out var row) ||
            !TryConvertColumn(columnText, out var column) ||
            row < 1 || row > ExcelMaximumRow || column < 1 || column > ExcelMaximumColumn)
        {
            return false;
        }

        endpoint = new FormulaReferenceEndpoint(
            source,
            new FormulaReferenceCoordinate(rowAbsolute ? FormulaCoordinateKind.Absolute : FormulaCoordinateKind.Relative, row),
            new FormulaReferenceCoordinate(columnAbsolute ? FormulaCoordinateKind.Absolute : FormulaCoordinateKind.Relative, column));
        return true;
    }

    private static bool TryParseR1C1Endpoint(string source, out FormulaReferenceEndpoint? endpoint)
    {
        endpoint = null;
        var match = R1C1EndpointPattern.Match(source);
        if (!match.Success ||
            !TryParseR1C1Coordinate(match.Groups["row"].Value, out var row) ||
            !TryParseR1C1Coordinate(match.Groups["column"].Value, out var column))
        {
            return false;
        }

        endpoint = new FormulaReferenceEndpoint(source, row, column);
        return true;
    }

    private static bool TryParseR1C1Coordinate(string source, out FormulaReferenceCoordinate coordinate)
    {
        if (string.IsNullOrEmpty(source))
        {
            coordinate = new FormulaReferenceCoordinate(FormulaCoordinateKind.Current, 0);
            return true;
        }

        var relative = source[0] == '[';
        var numericText = relative ? source.Substring(1, source.Length - 2) : source;
        if (!int.TryParse(numericText, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var value) ||
            (!relative && value < 1))
        {
            coordinate = default;
            return false;
        }

        coordinate = new FormulaReferenceCoordinate(
            relative ? FormulaCoordinateKind.Relative : FormulaCoordinateKind.Absolute,
            value);
        return true;
    }

    private static bool TryConvertColumn(string source, out int column)
    {
        column = 0;
        foreach (var character in source)
        {
            var value = char.ToUpperInvariant(character) - 'A' + 1;
            if (value < 1 || value > 26)
            {
                return false;
            }

            column = checked((column * 26) + value);
        }

        return source.Length != 0;
    }

    private static bool TryReadDelimited(string source, ref int index, char delimiter)
    {
        index++;
        while (index < source.Length)
        {
            if (source[index] != delimiter)
            {
                index++;
                continue;
            }

            if (index + 1 < source.Length && source[index + 1] == delimiter)
            {
                index += 2;
                continue;
            }

            index++;
            return true;
        }

        return false;
    }

    private static bool TryReadBracketed(string source, ref int index)
    {
        var depth = 0;
        while (index < source.Length)
        {
            if (source[index] == '[')
            {
                depth++;
            }
            else if (source[index] == ']')
            {
                depth--;
                if (depth == 0)
                {
                    index++;
                    return true;
                }
            }

            index++;
        }

        return false;
    }

    private static bool TryReadErrorLiteral(string source, ref int index)
    {
        var match = ErrorLiteralPattern.Match(source, index);
        if (!match.Success || match.Index != index)
        {
            return false;
        }

        var end = index + match.Length;
        if (end < source.Length && IsIdentifierPart(source[end]))
        {
            return false;
        }

        index = end;
        return true;
    }

    private static void ReadNumber(string source, char decimalSeparator, ref int index)
    {
        var sawDecimal = false;
        if (source[index] == decimalSeparator)
        {
            sawDecimal = true;
            index++;
        }

        while (index < source.Length && char.IsDigit(source[index]))
        {
            index++;
        }

        if (!sawDecimal && index < source.Length && source[index] == decimalSeparator)
        {
            sawDecimal = true;
            index++;
            while (index < source.Length && char.IsDigit(source[index]))
            {
                index++;
            }
        }

        if (index < source.Length && (source[index] == 'E' || source[index] == 'e'))
        {
            var exponentStart = index;
            index++;
            if (index < source.Length && (source[index] == '+' || source[index] == '-'))
            {
                index++;
            }

            var digitStart = index;
            while (index < source.Length && char.IsDigit(source[index]))
            {
                index++;
            }

            if (digitStart == index)
            {
                index = exponentStart;
            }
        }
    }

    private static bool IsIdentifierStart(char value) =>
        char.IsLetter(value) || value == '_' || value == '\\';

    private static bool IsIdentifierPart(char value) =>
        char.IsLetterOrDigit(value) || value == '_' || value == '.' || value == '\\';

    private static bool IsOperator(char value) =>
        value == '+' || value == '-' || value == '*' || value == '/' || value == '^' ||
        value == '&' || value == '%' || value == '=' || value == '<' || value == '>' ||
        value == ':' || value == '!' || value == '@' || value == '#';

    private static bool IsTwoCharacterOperator(char first, char second) =>
        (first == '<' && (second == '=' || second == '>')) ||
        (first == '>' && second == '=');

    private enum ReferenceReadOutcome
    {
        NoMatch,
        Success,
        Invalid
    }

    private sealed class CoverageClassification
    {
        private CoverageClassification(FormulaCoverageDisposition disposition, string? limitationCode)
        {
            Disposition = disposition;
            LimitationCode = limitationCode;
        }

        public FormulaCoverageDisposition Disposition { get; }

        public string? LimitationCode { get; }

        public static CoverageClassification Transform() => new CoverageClassification(FormulaCoverageDisposition.Transform, null);

        public static CoverageClassification RoundTrip() => new CoverageClassification(FormulaCoverageDisposition.RoundTrip, null);

        public static CoverageClassification InspectOnly(string limitationCode) =>
            new CoverageClassification(FormulaCoverageDisposition.InspectOnly, limitationCode);
    }
}
