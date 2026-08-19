using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace ExcelAccel.Core.Formulas;

/// <summary>
/// Performs source-span-owned transformations over the deliberately narrow A1
/// formula subset qualified by <see cref="FormulaParser"/>. It never returns a
/// partially transformed formula.
/// </summary>
public sealed class A1FormulaTransformer
{
    private const int ExcelMaximumColumn = 16384;
    private const int ExcelMaximumRow = 1048576;
    private readonly FormulaParser _parser;

    public A1FormulaTransformer(FormulaParser? parser = null)
    {
        _parser = parser ?? new FormulaParser();
    }

    public FormulaTransformResult Translate(
        string formula,
        int rowDisplacement,
        int columnDisplacement,
        FormulaParseOptions? options = null)
    {
        var parsed = RequireTransformable(formula, options, out var document);
        if (parsed is not null)
        {
            return parsed;
        }

        var replacements = new List<FormulaReferenceChange>();
        foreach (var reference in document!.References)
        {
            if (!TryTranslateReference(reference, rowDisplacement, columnDisplacement, out var translated))
            {
                return FormulaTransformResult.Refused(
                    FormulaTransformRefusalCodes.ReferenceOutOfBounds,
                    $"Translating reference '{reference.SourceText}' would leave Excel's worksheet bounds.");
            }

            if (!string.Equals(reference.SourceText, translated, StringComparison.Ordinal))
            {
                replacements.Add(new FormulaReferenceChange(reference.Span, reference.SourceText, translated));
            }
        }

        return FormulaTransformResult.Success(ApplyReplacements(formula, replacements), replacements);
    }

    public FormulaTransformResult Transpose(
        string formula,
        int sourceFormulaRow,
        int sourceFormulaColumn,
        int destinationFormulaRow,
        int destinationFormulaColumn,
        FormulaParseOptions? options = null)
    {
        if (!IsCell(sourceFormulaRow, sourceFormulaColumn) ||
            !IsCell(destinationFormulaRow, destinationFormulaColumn))
        {
            return FormulaTransformResult.Refused(
                FormulaTransformRefusalCodes.InvalidTransformArgument,
                "Source and destination formula coordinates must be valid Excel cells.");
        }

        var parsed = RequireTransformable(formula, options, out var document);
        if (parsed is not null)
        {
            return parsed;
        }

        var replacements = new List<FormulaReferenceChange>();
        foreach (var reference in document!.References)
        {
            if (!TryTransposeReference(
                    reference,
                    sourceFormulaRow,
                    sourceFormulaColumn,
                    destinationFormulaRow,
                    destinationFormulaColumn,
                    out var translated))
            {
                return FormulaTransformResult.Refused(
                    FormulaTransformRefusalCodes.ReferenceOutOfBounds,
                    $"Transposing reference '{reference.SourceText}' would leave Excel's worksheet bounds.");
            }

            if (!string.Equals(reference.SourceText, translated, StringComparison.Ordinal))
            {
                replacements.Add(new FormulaReferenceChange(reference.Span, reference.SourceText, translated));
            }
        }

        return FormulaTransformResult.Success(ApplyReplacements(formula, replacements), replacements);
    }

    public FormulaTransformResult ToggleReferenceAtCaret(
        string formula,
        int caretOffset,
        FormulaParseOptions? options = null)
    {
        var parsed = RequireTransformable(formula, options, out var document);
        if (parsed is not null)
        {
            return parsed;
        }

        if (caretOffset < 1 || caretOffset > formula.Length)
        {
            return FormulaTransformResult.Refused(
                FormulaTransformRefusalCodes.ReferenceNotFound,
                "The caret is outside the formula expression.");
        }

        var candidates = document!.References
            .Where(reference => ContainsCaret(reference.Span, caretOffset))
            .ToArray();
        if (candidates.Length == 0)
        {
            return FormulaTransformResult.Refused(
                FormulaTransformRefusalCodes.ReferenceNotFound,
                "The caret is not within a qualified A1 reference.");
        }

        if (candidates.Length != 1)
        {
            return FormulaTransformResult.Refused(
                FormulaTransformRefusalCodes.AmbiguousCaret,
                "The caret maps to more than one reference.");
        }

        var reference = candidates[0];
        var relativeOffset = Math.Min(caretOffset, reference.Span.End - 1) - reference.Span.Start;
        if (!TryToggleEndpoint(reference, relativeOffset, out var toggled))
        {
            return FormulaTransformResult.Refused(
                FormulaTransformRefusalCodes.ReferenceNotFound,
                "The caret is on a qualifier or range separator rather than an A1 endpoint.");
        }

        var change = new FormulaReferenceChange(reference.Span, reference.SourceText, toggled);
        return FormulaTransformResult.Success(ApplyReplacements(formula, new[] { change }), new[] { change });
    }

    private FormulaTransformResult? RequireTransformable(
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

    private static bool TryTranslateReference(
        FormulaReference reference,
        int rowDisplacement,
        int columnDisplacement,
        out string translated)
    {
        if (!TryTranslateEndpoint(reference.First, rowDisplacement, columnDisplacement, out var first) ||
            (reference.Second is not null &&
             !TryTranslateEndpoint(reference.Second, rowDisplacement, columnDisplacement, out _)))
        {
            translated = string.Empty;
            return false;
        }

        string? second = null;
        if (reference.Second is not null)
        {
            TryTranslateEndpoint(reference.Second, rowDisplacement, columnDisplacement, out second);
        }

        translated = RenderReference(reference, first!, second);
        return true;
    }

    private static bool TryTranslateEndpoint(
        FormulaReferenceEndpoint endpoint,
        int rowDisplacement,
        int columnDisplacement,
        out string? translated)
    {
        long row = endpoint.Row.Value;
        long column = endpoint.Column.Value;
        if (endpoint.Row.Kind == FormulaCoordinateKind.Relative)
        {
            row = (long)row + rowDisplacement;
        }

        if (endpoint.Column.Kind == FormulaCoordinateKind.Relative)
        {
            column = (long)column + columnDisplacement;
        }

        if (!IsCell(row, column))
        {
            translated = null;
            return false;
        }

        translated = RenderEndpoint(
            (int)row,
            (int)column,
            endpoint.Row.Kind == FormulaCoordinateKind.Absolute,
            endpoint.Column.Kind == FormulaCoordinateKind.Absolute,
            endpoint.SourceText);
        return true;
    }

    private static bool TryTransposeReference(
        FormulaReference reference,
        int sourceRow,
        int sourceColumn,
        int destinationRow,
        int destinationColumn,
        out string translated)
    {
        if (!TryTransposeEndpoint(reference.First, sourceRow, sourceColumn, destinationRow, destinationColumn, out var first) ||
            (reference.Second is not null &&
             !TryTransposeEndpoint(reference.Second, sourceRow, sourceColumn, destinationRow, destinationColumn, out _)))
        {
            translated = string.Empty;
            return false;
        }

        string? second = null;
        if (reference.Second is not null)
        {
            TryTransposeEndpoint(reference.Second, sourceRow, sourceColumn, destinationRow, destinationColumn, out second);
        }

        translated = RenderReference(reference, first!, second);
        return true;
    }

    private static bool TryTransposeEndpoint(
        FormulaReferenceEndpoint endpoint,
        int sourceRow,
        int sourceColumn,
        int destinationRow,
        int destinationColumn,
        out string? translated)
    {
        long row;
        long column;
        var rowAbsolute = endpoint.Column.Kind == FormulaCoordinateKind.Absolute;
        var columnAbsolute = endpoint.Row.Kind == FormulaCoordinateKind.Absolute;

        if (endpoint.Column.Kind == FormulaCoordinateKind.Absolute)
        {
            row = endpoint.Column.Value;
        }
        else
        {
            row = (long)destinationRow + endpoint.Column.Value - sourceColumn;
        }

        if (endpoint.Row.Kind == FormulaCoordinateKind.Absolute)
        {
            column = endpoint.Row.Value;
        }
        else
        {
            column = (long)destinationColumn + endpoint.Row.Value - sourceRow;
        }

        if (!IsCell(row, column))
        {
            translated = null;
            return false;
        }

        // Both coordinate values and their absolute/relative anchors exchange
        // axes. Relative offsets are then based at the transposed formula cell.
        translated = RenderEndpoint((int)row, (int)column, rowAbsolute, columnAbsolute, endpoint.SourceText);
        return true;
    }

    private static bool TryToggleEndpoint(FormulaReference reference, int relativeOffset, out string toggled)
    {
        var qualifierLength = reference.Qualifier?.Length ?? 0;
        if (qualifierLength != 0)
        {
            qualifierLength++;
        }

        var firstStart = qualifierLength;
        var firstEnd = firstStart + reference.First.SourceText.Length;
        FormulaReferenceEndpoint? endpoint = null;
        var endpointStart = 0;
        if (relativeOffset >= firstStart && relativeOffset < firstEnd)
        {
            endpoint = reference.First;
            endpointStart = firstStart;
        }
        else if (reference.Second is not null)
        {
            var secondStart = firstEnd + 1;
            var secondEnd = secondStart + reference.Second.SourceText.Length;
            if (relativeOffset >= secondStart && relativeOffset < secondEnd)
            {
                endpoint = reference.Second;
                endpointStart = secondStart;
            }
        }

        if (endpoint is null)
        {
            toggled = string.Empty;
            return false;
        }

        var rowAbsolute = endpoint.Row.Kind == FormulaCoordinateKind.Absolute;
        var columnAbsolute = endpoint.Column.Kind == FormulaCoordinateKind.Absolute;
        bool nextRowAbsolute;
        bool nextColumnAbsolute;
        if (!rowAbsolute && !columnAbsolute)
        {
            nextRowAbsolute = true;
            nextColumnAbsolute = true;
        }
        else if (rowAbsolute && columnAbsolute)
        {
            nextRowAbsolute = true;
            nextColumnAbsolute = false;
        }
        else if (rowAbsolute)
        {
            nextRowAbsolute = false;
            nextColumnAbsolute = true;
        }
        else
        {
            nextRowAbsolute = false;
            nextColumnAbsolute = false;
        }

        var endpointText = RenderEndpoint(
            endpoint.Row.Value,
            endpoint.Column.Value,
            nextRowAbsolute,
            nextColumnAbsolute,
            endpoint.SourceText);
        toggled = reference.SourceText.Substring(0, endpointStart) + endpointText +
                  reference.SourceText.Substring(endpointStart + endpoint.SourceText.Length);
        return true;
    }

    private static string RenderReference(FormulaReference reference, string first, string? second)
    {
        var builder = new StringBuilder();
        if (reference.HasImplicitIntersection)
        {
            builder.Append('@');
        }

        if (reference.Qualifier is not null)
        {
            builder.Append(reference.Qualifier).Append('!');
        }

        builder.Append(first);
        if (second is not null)
        {
            builder.Append(':').Append(second);
        }

        if (reference.HasSpillOperator)
        {
            builder.Append('#');
        }

        return builder.ToString();
    }

    private static string RenderEndpoint(
        int row,
        int column,
        bool rowAbsolute,
        bool columnAbsolute,
        string original)
    {
        var lowerCaseColumn = original.Any(char.IsLower) && !original.Any(char.IsUpper);
        var columnText = ColumnName(column);
        if (lowerCaseColumn)
        {
            columnText = columnText.ToLowerInvariant();
        }

        return (columnAbsolute ? "$" : string.Empty) + columnText +
               (rowAbsolute ? "$" : string.Empty) + row.ToString(CultureInfo.InvariantCulture);
    }

    private static string ColumnName(int column)
    {
        var builder = new StringBuilder();
        var remaining = column;
        while (remaining > 0)
        {
            remaining--;
            builder.Insert(0, (char)('A' + (remaining % 26)));
            remaining /= 26;
        }

        return builder.ToString();
    }

    private static bool ContainsCaret(FormulaSourceSpan span, int caretOffset) =>
        caretOffset >= span.Start && caretOffset <= span.End;

    private static bool IsCell(long row, long column) =>
        row >= 1 && row <= ExcelMaximumRow && column >= 1 && column <= ExcelMaximumColumn;

    private static string ApplyReplacements(string source, IReadOnlyList<FormulaReferenceChange> replacements)
    {
        var builder = new StringBuilder(source);
        foreach (var replacement in replacements.OrderByDescending(value => value.SourceSpan.Start))
        {
            builder.Remove(replacement.SourceSpan.Start, replacement.SourceSpan.Length);
            builder.Insert(replacement.SourceSpan.Start, replacement.AfterText);
        }

        return builder.ToString();
    }
}
