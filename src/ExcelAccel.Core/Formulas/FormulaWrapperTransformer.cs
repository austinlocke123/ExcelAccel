using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAccel.Core.Formulas;

/// <summary>
/// Whole-expression wrappers preserve operator precedence by construction.
/// The parser still owns admission; inspect-only syntax never reaches output.
/// </summary>
public sealed class FormulaWrapperTransformer
{
    private readonly FormulaParser _parser;

    public FormulaWrapperTransformer(FormulaParser? parser = null)
    {
        _parser = parser ?? new FormulaParser();
    }

    public FormulaTransformResult ToggleIfError(
        string formula,
        string fallbackExpression,
        FormulaParseOptions? options = null)
    {
        var admission = Admit(formula, options, out var document);
        if (admission is not null)
        {
            return admission;
        }

        if (string.IsNullOrWhiteSpace(fallbackExpression) || fallbackExpression[0] == '=')
        {
            return InvalidArgument("The IFERROR fallback must be a nonempty formula expression without '='.");
        }

        var fallbackAdmission = Admit("=" + fallbackExpression, options, out _);
        if (fallbackAdmission is not null)
        {
            return InvalidArgument("The IFERROR fallback is outside the qualified formula subset.");
        }

        if (TryReadTopLevelIfError(document!, out var interior, out var existingFallback) &&
            string.Equals(existingFallback, fallbackExpression, StringComparison.Ordinal))
        {
            return FormulaTransformResult.Success("=" + interior);
        }

        return FormulaTransformResult.Success(
            "=IFERROR(" + formula.Substring(1) + document!.Dialect.ListSeparator + fallbackExpression + ")");
    }

    public FormulaTransformResult ReverseSign(string formula, FormulaParseOptions? options = null)
    {
        var admission = Admit(formula, options, out var document);
        if (admission is not null)
        {
            return admission;
        }

        const string prefix = "=-(";
        if (formula.StartsWith(prefix, StringComparison.Ordinal) &&
            formula.EndsWith(")", StringComparison.Ordinal) &&
            IsCanonicalNegation(document!))
        {
            return FormulaTransformResult.Success("=" + formula.Substring(prefix.Length, formula.Length - prefix.Length - 1));
        }

        return FormulaTransformResult.Success("=-(" + formula.Substring(1) + ")");
    }

    public FormulaTransformResult Scale(string formula, long scale, bool divide, FormulaParseOptions? options = null)
    {
        var admission = Admit(formula, options, out _);
        if (admission is not null)
        {
            return admission;
        }

        if (scale != 1000 && scale != 1000000)
        {
            return InvalidArgument("The qualified unit scale must be 1,000 or 1,000,000.");
        }

        return FormulaTransformResult.Success(
            "=(" + formula.Substring(1) + ")" + (divide ? "/" : "*") +
            scale.ToString(System.Globalization.CultureInfo.InvariantCulture));
    }

    private FormulaTransformResult? Admit(
        string formula,
        FormulaParseOptions? options,
        out FormulaSyntaxDocument? document)
    {
        var effectiveOptions = options ?? FormulaParseOptions.DefaultA1;
        if (effectiveOptions.Dialect.Notation != FormulaNotation.A1)
        {
            document = null;
            return FormulaTransformResult.Refused(
                FormulaTransformRefusalCodes.UnsupportedNotation,
                "Production formula transforms are qualified only for A1 notation.");
        }

        var parsed = _parser.Parse(formula, effectiveOptions);
        document = parsed.Document;
        if (!parsed.IsSuccess)
        {
            return FormulaTransformResult.Refused(
                parsed.RefusalCode ?? FormulaRefusalCodes.InvalidSyntax,
                parsed.Message ?? "The formula could not be parsed safely.");
        }

        if (document!.Disposition == FormulaCoverageDisposition.InspectOnly)
        {
            return FormulaTransformResult.Refused(
                document.LimitationCode ?? FormulaTransformRefusalCodes.UnsupportedCoverage,
                "The formula is inspectable but is outside the qualified mutation subset.");
        }

        return null;
    }

    private static bool TryReadTopLevelIfError(
        FormulaSyntaxDocument document,
        out string interior,
        out string fallback)
    {
        interior = string.Empty;
        fallback = string.Empty;
        var significant = document.Tokens
            .Where(token => token.Kind != FormulaTokenKind.Whitespace && token.Kind != FormulaTokenKind.Prefix)
            .ToArray();
        if (significant.Length < 5 ||
            significant[0].Kind != FormulaTokenKind.Identifier ||
            !significant[0].Text.Equals("IFERROR", StringComparison.OrdinalIgnoreCase) ||
            significant[1].Kind != FormulaTokenKind.OpenParenthesis ||
            significant[significant.Length - 1].Kind != FormulaTokenKind.CloseParenthesis)
        {
            return false;
        }

        var depth = 0;
        FormulaToken? separator = null;
        for (var index = 1; index < significant.Length; index++)
        {
            var token = significant[index];
            if (token.Kind == FormulaTokenKind.OpenParenthesis)
            {
                depth++;
            }
            else if (token.Kind == FormulaTokenKind.CloseParenthesis)
            {
                depth--;
                if (depth == 0 && index != significant.Length - 1)
                {
                    return false;
                }
            }
            else if (token.Kind == FormulaTokenKind.Separator && depth == 1)
            {
                if (separator is not null)
                {
                    return false;
                }

                separator = token;
            }
        }

        if (separator is null || depth != 0)
        {
            return false;
        }

        var opening = significant[1];
        var closing = significant[significant.Length - 1];
        interior = document.SourceText.Substring(opening.Span.End, separator.Span.Start - opening.Span.End);
        fallback = document.SourceText.Substring(separator.Span.End, closing.Span.Start - separator.Span.End);
        return interior.Length != 0 && fallback.Length != 0;
    }

    private static bool IsCanonicalNegation(FormulaSyntaxDocument document)
    {
        var significant = document.Tokens
            .Where(token => token.Kind != FormulaTokenKind.Whitespace && token.Kind != FormulaTokenKind.Prefix)
            .ToArray();
        if (significant.Length < 4 ||
            significant[0].Kind != FormulaTokenKind.Operator ||
            significant[0].Text != "-" ||
            significant[1].Kind != FormulaTokenKind.OpenParenthesis ||
            significant[significant.Length - 1].Kind != FormulaTokenKind.CloseParenthesis)
        {
            return false;
        }

        var depth = 0;
        for (var index = 1; index < significant.Length; index++)
        {
            var token = significant[index];
            if (token.Kind == FormulaTokenKind.OpenParenthesis)
            {
                depth++;
            }
            else if (token.Kind == FormulaTokenKind.CloseParenthesis)
            {
                depth--;
                if (depth == 0)
                {
                    return index == significant.Length - 1;
                }
            }
        }

        return false;
    }

    private static FormulaTransformResult InvalidArgument(string message) =>
        FormulaTransformResult.Refused(FormulaTransformRefusalCodes.InvalidTransformArgument, message);
}
