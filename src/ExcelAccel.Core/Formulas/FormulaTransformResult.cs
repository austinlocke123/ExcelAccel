using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAccel.Core.Formulas;

public static class FormulaTransformRefusalCodes
{
    public const string UnsupportedCoverage = "FORMULA_TRANSFORM_COVERAGE_UNSUPPORTED";
    public const string UnsupportedNotation = "FORMULA_TRANSFORM_NOTATION_UNSUPPORTED";
    public const string ReferenceNotFound = "FORMULA_REFERENCE_NOT_FOUND";
    public const string AmbiguousCaret = "FORMULA_REFERENCE_CARET_AMBIGUOUS";
    public const string ReferenceOutOfBounds = "FORMULA_REFERENCE_OUT_OF_BOUNDS";
    public const string InvalidTransformArgument = "FORMULA_TRANSFORM_ARGUMENT_INVALID";
}

public sealed class FormulaReferenceChange
{
    public FormulaReferenceChange(FormulaSourceSpan sourceSpan, string beforeText, string afterText)
    {
        SourceSpan = sourceSpan;
        BeforeText = beforeText ?? throw new ArgumentNullException(nameof(beforeText));
        AfterText = afterText ?? throw new ArgumentNullException(nameof(afterText));
    }

    public FormulaSourceSpan SourceSpan { get; }

    public string BeforeText { get; }

    public string AfterText { get; }
}

public sealed class FormulaTransformResult
{
    private FormulaTransformResult(
        string? formula,
        IEnumerable<FormulaReferenceChange>? changes,
        string? refusalCode,
        string? message)
    {
        Formula = formula;
        Changes = Array.AsReadOnly((changes ?? Array.Empty<FormulaReferenceChange>()).ToArray());
        RefusalCode = refusalCode;
        Message = message;
    }

    public bool IsSuccess => Formula is not null;

    public string? Formula { get; }

    public IReadOnlyList<FormulaReferenceChange> Changes { get; }

    public string? RefusalCode { get; }

    public string? Message { get; }

    public static FormulaTransformResult Success(string formula, IEnumerable<FormulaReferenceChange>? changes = null) =>
        new FormulaTransformResult(
            formula ?? throw new ArgumentNullException(nameof(formula)),
            changes,
            null,
            null);

    public static FormulaTransformResult Refused(string code, string message) =>
        new FormulaTransformResult(
            null,
            null,
            Require(code, nameof(code)),
            Require(message, nameof(message)));

    private static string Require(string value, string parameterName) =>
        string.IsNullOrWhiteSpace(value)
            ? throw new ArgumentException("A stable refusal value is required.", parameterName)
            : value;
}
