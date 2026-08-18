using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using ExcelAccel.Core.Formulas;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class FormulaParserTests
{
    private readonly FormulaParser _parser = new FormulaParser();

    [Fact]
    public void VersionedCorpusMatchesCoverageAndRefusalContract()
    {
        var corpusPath = Path.Combine(AppContext.BaseDirectory, "Fixtures", "formula-v1-corpus.json");
        var corpus = JsonSerializer.Deserialize<List<CorpusCase>>(
            File.ReadAllText(corpusPath),
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

        Assert.NotNull(corpus);
        Assert.NotEmpty(corpus!);
        foreach (var testCase in corpus!)
        {
            var notation = Enum.Parse<FormulaNotation>(testCase.Notation, ignoreCase: true);
            var expectedDisposition = Enum.Parse<FormulaCoverageDisposition>(testCase.ExpectedDisposition, ignoreCase: true);
            var dialect = new FormulaDialect(
                notation,
                RequireSingleCharacter(testCase.ListSeparator),
                RequireSingleCharacter(testCase.DecimalSeparator));

            var result = _parser.Parse(testCase.Formula, new FormulaParseOptions(dialect));
            if (expectedDisposition == FormulaCoverageDisposition.Refuse)
            {
                Assert.False(result.IsSuccess);
                Assert.Equal(testCase.ExpectedLimitationCode, result.RefusalCode);
                Assert.Null(result.Document);
                continue;
            }

            Assert.True(result.IsSuccess, $"Corpus case '{testCase.Id}' refused: {result.RefusalCode} {result.Message}");
            var document = Assert.IsType<FormulaSyntaxDocument>(result.Document);
            Assert.True(
                expectedDisposition == document.Disposition,
                $"Corpus case '{testCase.Id}' expected {expectedDisposition} but received {document.Disposition} ({document.LimitationCode}).");
            Assert.Equal(testCase.ExpectedLimitationCode, document.LimitationCode);
            Assert.Equal(testCase.Formula, document.Serialize());
            Assert.Equal(testCase.Formula, string.Concat(document.Tokens.Select(token => token.Text)));
            Assert.Equal(testCase.ExpectedReferenceTexts, document.References.Select(reference => reference.SourceText).ToArray());
            AssertSpansAreContiguous(document);
        }
    }

    [Fact]
    public void A1ReferenceModelRetainsAbsoluteAndRelativeCoordinates()
    {
        var result = _parser.Parse("=$B3:C$4");

        var reference = Assert.Single(Assert.IsType<FormulaSyntaxDocument>(result.Document).References);
        Assert.Equal(FormulaCoordinateKind.Absolute, reference.First.Column.Kind);
        Assert.Equal(2, reference.First.Column.Value);
        Assert.Equal(FormulaCoordinateKind.Relative, reference.First.Row.Kind);
        Assert.Equal(3, reference.First.Row.Value);
        Assert.NotNull(reference.Second);
        Assert.Equal(FormulaCoordinateKind.Relative, reference.Second!.Column.Kind);
        Assert.Equal(3, reference.Second.Column.Value);
        Assert.Equal(FormulaCoordinateKind.Absolute, reference.Second.Row.Kind);
        Assert.Equal(4, reference.Second.Row.Value);
    }

    [Fact]
    public void R1C1ReferenceModelRetainsCurrentRelativeAndAbsoluteCoordinates()
    {
        var options = new FormulaParseOptions(FormulaDialect.InvariantR1C1);
        var result = _parser.Parse("=R[-2]C:R10C[3]", options);

        var reference = Assert.Single(Assert.IsType<FormulaSyntaxDocument>(result.Document).References);
        Assert.Equal(FormulaCoordinateKind.Relative, reference.First.Row.Kind);
        Assert.Equal(-2, reference.First.Row.Value);
        Assert.Equal(FormulaCoordinateKind.Current, reference.First.Column.Kind);
        Assert.NotNull(reference.Second);
        Assert.Equal(FormulaCoordinateKind.Absolute, reference.Second!.Row.Kind);
        Assert.Equal(10, reference.Second.Row.Value);
        Assert.Equal(FormulaCoordinateKind.Relative, reference.Second.Column.Kind);
        Assert.Equal(3, reference.Second.Column.Value);
    }

    [Theory]
    [InlineData("=12345", 5, 100, 10, FormulaRefusalCodes.TooLong)]
    [InlineData("=1+2", 100, 3, 10, FormulaRefusalCodes.TooManyTokens)]
    [InlineData("=(((1)))", 100, 100, 2, FormulaRefusalCodes.NestingLimit)]
    public void ResourceLimitsRefuseWithoutPartialDocument(
        string formula,
        int maximumLength,
        int maximumTokens,
        int maximumNesting,
        string expectedCode)
    {
        var options = new FormulaParseOptions(
            FormulaDialect.InvariantA1,
            maximumLength,
            maximumTokens,
            maximumNesting);

        var result = _parser.Parse(formula, options);

        Assert.False(result.IsSuccess);
        Assert.Null(result.Document);
        Assert.Equal(expectedCode, result.RefusalCode);
    }

    [Fact]
    public void RepeatedParseIsDeterministic()
    {
        const string formula = "=IF(A1>0, 'Model Sheet'!$B$2, \"A1\")";

        var first = Assert.IsType<FormulaSyntaxDocument>(_parser.Parse(formula).Document);
        var second = Assert.IsType<FormulaSyntaxDocument>(_parser.Parse(formula).Document);

        Assert.Equal(first.Serialize(), second.Serialize());
        Assert.Equal(
            first.Tokens.Select(token => (token.Kind, token.Text, token.Span.Start, token.Span.Length)),
            second.Tokens.Select(token => (token.Kind, token.Text, token.Span.Start, token.Span.Length)));
        Assert.Equal(
            first.References.Select(reference => (reference.SourceText, reference.Span.Start, reference.Span.Length)),
            second.References.Select(reference => (reference.SourceText, reference.Span.Start, reference.Span.Length)));
    }

    [Fact]
    public void ReferenceLookalikesInsideStringsAreNeverReferences()
    {
        var document = Assert.IsType<FormulaSyntaxDocument>(_parser.Parse("=\"$A$1 and R1C1\"").Document);

        Assert.Empty(document.References);
        Assert.Equal(FormulaCoverageDisposition.RoundTrip, document.Disposition);
    }

    [Fact]
    public void InvalidDialectConfigurationIsRejected()
    {
        Assert.Throws<ArgumentException>(() => new FormulaDialect(FormulaNotation.A1, ',', ','));
        Assert.Throws<ArgumentOutOfRangeException>(() => new FormulaDialect(FormulaNotation.A1, '|', '.'));
        Assert.Throws<ArgumentOutOfRangeException>(() => new FormulaDialect(FormulaNotation.A1, ';', ':'));
    }

    [Fact]
    public void SyntaxDocumentDefensivelyCopiesCollections()
    {
        var tokens = new List<FormulaToken>
        {
            new FormulaToken(FormulaTokenKind.Prefix, "=", new FormulaSourceSpan(0, 1)),
            new FormulaToken(FormulaTokenKind.Number, "1", new FormulaSourceSpan(1, 1))
        };
        var references = new List<FormulaReference>();
        var document = new FormulaSyntaxDocument(
            "=1",
            FormulaDialect.InvariantA1,
            tokens,
            references,
            FormulaCoverageDisposition.RoundTrip,
            null);

        tokens.Clear();
        references.Clear();

        Assert.Equal(2, document.Tokens.Count);
        Assert.Empty(document.References);
    }

    [Fact]
    public void ParseOptionsCannotDisableHardResourceCeilings()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new FormulaParseOptions(FormulaDialect.InvariantA1, FormulaParseOptions.HardMaximumLength + 1));
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new FormulaParseOptions(
                FormulaDialect.InvariantA1,
                maximumTokens: FormulaParseOptions.HardMaximumTokens + 1));
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new FormulaParseOptions(
                FormulaDialect.InvariantA1,
                maximumNesting: FormulaParseOptions.HardMaximumNesting + 1));
    }

    [Fact]
    public void GeneratedSupportedCorpusRoundTripsWithoutReferenceDrift()
    {
        for (var index = 1; index <= 500; index++)
        {
            var firstColumn = (char)('A' + (index % 26));
            var secondColumn = (char)('A' + ((index + 7) % 26));
            var row = (index % 1000) + 1;
            var formula = $"=${firstColumn}{row}+SUM({secondColumn}${row},$C$3)*{index}";

            var result = _parser.Parse(formula);

            Assert.True(result.IsSuccess, $"Generated formula refused: {formula} ({result.RefusalCode})");
            var document = Assert.IsType<FormulaSyntaxDocument>(result.Document);
            Assert.Equal(FormulaCoverageDisposition.Transform, document.Disposition);
            Assert.Equal(formula, document.Serialize());
            Assert.Equal(3, document.References.Count);
            AssertSpansAreContiguous(document);
        }
    }

    [Fact]
    public void ParserIsSafeForConcurrentPureCoreUse()
    {
        var failures = new ConcurrentQueue<string>();

        Parallel.For(0, 250, index =>
        {
            var formula = $"=A{index + 1}+SUM($B$2,C3)";
            var result = _parser.Parse(formula);
            if (!result.IsSuccess || result.Document?.Serialize() != formula)
            {
                failures.Enqueue($"{formula}: {result.RefusalCode}");
            }
        });

        Assert.Empty(failures);
    }

    [Theory]
    [InlineData("=\"unterminated", FormulaRefusalCodes.UnterminatedString)]
    [InlineData("=A1 2", FormulaRefusalCodes.InvalidSyntax)]
    [InlineData("=SUM(,A1)", FormulaRefusalCodes.InvalidSyntax)]
    [InlineData("=A1🙂", FormulaRefusalCodes.UnsupportedCharacter)]
    public void HostileOrAmbiguousSyntaxRefusesWithoutDocument(string formula, string expectedCode)
    {
        var result = _parser.Parse(formula);

        Assert.False(result.IsSuccess);
        Assert.Equal(expectedCode, result.RefusalCode);
        Assert.Null(result.Document);
    }

    private static void AssertSpansAreContiguous(FormulaSyntaxDocument document)
    {
        var cursor = 0;
        foreach (var token in document.Tokens)
        {
            Assert.Equal(cursor, token.Span.Start);
            Assert.Equal(token.Text.Length, token.Span.Length);
            cursor = token.Span.End;
        }

        Assert.Equal(document.SourceText.Length, cursor);
    }

    private static char RequireSingleCharacter(string value)
    {
        Assert.Single(value);
        return value[0];
    }

    private sealed class CorpusCase
    {
        public string Id { get; set; } = string.Empty;

        public string Formula { get; set; } = string.Empty;

        public string Notation { get; set; } = string.Empty;

        public string ListSeparator { get; set; } = string.Empty;

        public string DecimalSeparator { get; set; } = string.Empty;

        public string ExpectedDisposition { get; set; } = string.Empty;

        public string? ExpectedLimitationCode { get; set; }

        public string[] ExpectedReferenceTexts { get; set; } = Array.Empty<string>();
    }
}
