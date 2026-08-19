using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using ExcelAccel.Core.Formulas;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class FormulaTransformerTests
{
    private readonly A1FormulaTransformer _references = new A1FormulaTransformer();
    private readonly FormulaWrapperTransformer _wrappers = new FormulaWrapperTransformer();

    [Theory]
    [InlineData("=A1+$B2+C$3+$D$4", 2, 3, "=D3+$B4+F$3+$D$4")]
    [InlineData("='Model Sheet'!a1:bc2", 1, 1, "='Model Sheet'!b2:bd3")]
    [InlineData("=SUM(A1,B2)", -1, 2, "=SUM(C0,D1)", FormulaTransformRefusalCodes.ReferenceOutOfBounds)]
    public void TranslateHonorsRelativeAndAbsoluteAxes(
        string formula,
        int rowDelta,
        int columnDelta,
        string expected,
        string? expectedRefusal = null)
    {
        var result = _references.Translate(formula, rowDelta, columnDelta);

        if (expectedRefusal is not null)
        {
            Assert.False(result.IsSuccess);
            Assert.Equal(expectedRefusal, result.RefusalCode);
            Assert.Null(result.Formula);
            return;
        }

        Assert.True(result.IsSuccess, result.Message);
        Assert.Equal(expected, result.Formula);
    }

    [Fact]
    public void TranslationReturnsExactChangedReferenceSpans()
    {
        const string formula = "=IF(A1>0, 'Model'!$B2, \"A1\")";

        var result = _references.Translate(formula, 3, 4);

        Assert.True(result.IsSuccess, result.Message);
        Assert.Equal("=IF(E4>0, 'Model'!$B5, \"A1\")", result.Formula);
        Assert.Equal(new[] { "A1", "'Model'!$B2" }, result.Changes.Select(change => change.BeforeText));
        Assert.All(result.Changes, change =>
            Assert.Equal(change.BeforeText, formula.Substring(change.SourceSpan.Start, change.SourceSpan.Length)));
    }

    [Theory]
    [InlineData("=A1", 1, "=$A$1")]
    [InlineData("=$A$1", 2, "=A$1")]
    [InlineData("=A$1", 2, "=$A1")]
    [InlineData("=$A1", 2, "=A1")]
    [InlineData("='Model Sheet'!A1:B2", 15, "='Model Sheet'!$A$1:B2")]
    [InlineData("='Model Sheet'!A1:B2", 18, "='Model Sheet'!A1:$B$2")]
    public void CaretToggleCyclesOnlyTheContainingEndpoint(string formula, int caret, string expected)
    {
        var result = _references.ToggleReferenceAtCaret(formula, caret);

        Assert.True(result.IsSuccess, result.Message);
        Assert.Equal(expected, result.Formula);
        Assert.Single(result.Changes);
    }

    [Theory]
    [InlineData("=SUM(A1)", 2)]
    [InlineData("='Model Sheet'!A1", 3)]
    [InlineData("=A1:B2", 3)]
    public void CaretToggleRefusesNonEndpointPositions(string formula, int caret)
    {
        var result = _references.ToggleReferenceAtCaret(formula, caret);

        Assert.False(result.IsSuccess);
        Assert.Equal(FormulaTransformRefusalCodes.ReferenceNotFound, result.RefusalCode);
        Assert.Null(result.Formula);
    }

    [Theory]
    [InlineData("=A1", 2, 2, 10, 20, "=S9")]
    [InlineData("=C4", 2, 2, 10, 20, "=V11")]
    [InlineData("=$C4", 2, 2, 10, 20, "=V$3")]
    [InlineData("=C$4", 2, 2, 10, 20, "=$D11")]
    [InlineData("=$C$4", 2, 2, 10, 20, "=$D$3")]
    public void TransposeExchangesReferenceAxesAndAnchorKinds(
        string formula,
        int sourceRow,
        int sourceColumn,
        int destinationRow,
        int destinationColumn,
        string expected)
    {
        var result = _references.Transpose(
            formula,
            sourceRow,
            sourceColumn,
            destinationRow,
            destinationColumn);

        Assert.True(result.IsSuccess, result.Message);
        Assert.Equal(expected, result.Formula);
    }

    [Theory]
    [InlineData("=Table1[Amount]", FormulaRefusalCodes.StructuredReferenceInspectOnly)]
    [InlineData("=NamedRate*A1", FormulaRefusalCodes.NameInspectOnly)]
    [InlineData("='[Other.xlsx]Sheet1'!A1", FormulaRefusalCodes.ExternalReferenceInspectOnly)]
    [InlineData("=A1#", FormulaRefusalCodes.DynamicArrayInspectOnly)]
    public void EveryTransformRefusesInspectOnlyCoverage(string formula, string expectedCode)
    {
        var translated = _references.Translate(formula, 1, 0);
        var wrapped = _wrappers.ReverseSign(formula);

        Assert.False(translated.IsSuccess);
        Assert.Equal(expectedCode, translated.RefusalCode);
        Assert.Null(translated.Formula);
        Assert.False(wrapped.IsSuccess);
        Assert.Equal(expectedCode, wrapped.RefusalCode);
        Assert.Null(wrapped.Formula);
    }

    [Fact]
    public void R1C1MutationRemainsFailClosed()
    {
        var options = new FormulaParseOptions(FormulaDialect.InvariantR1C1);

        var result = _references.Translate("=RC[-1]", 1, 0, options);

        Assert.False(result.IsSuccess);
        Assert.Equal(FormulaTransformRefusalCodes.UnsupportedNotation, result.RefusalCode);
    }

    [Theory]
    [InlineData("=A1+B1", "0", "=IFERROR(A1+B1,0)")]
    [InlineData("=IFERROR(A1+B1,0)", "0", "=A1+B1")]
    [InlineData("=IFERROR(A1+B1,1)", "0", "=IFERROR(IFERROR(A1+B1,1),0)")]
    public void IfErrorToggleChangesOnlyTheExactConfiguredTopLevelWrapper(
        string formula,
        string fallback,
        string expected)
    {
        var result = _wrappers.ToggleIfError(formula, fallback);

        Assert.True(result.IsSuccess, result.Message);
        Assert.Equal(expected, result.Formula);
    }

    [Fact]
    public void IfErrorUsesTheSelectedListSeparator()
    {
        var options = new FormulaParseOptions(new FormulaDialect(FormulaNotation.A1, ';', ','));

        var result = _wrappers.ToggleIfError("=A1+1,5", "0", options);

        Assert.True(result.IsSuccess, result.Message);
        Assert.Equal("=IFERROR(A1+1,5;0)", result.Formula);
        Assert.Equal("=A1+1,5", _wrappers.ToggleIfError(result.Formula!, "0", options).Formula);
    }

    [Theory]
    [InlineData("=A1+B1")]
    [InlineData("=A1^2")]
    [InlineData("=-A1")]
    [InlineData("=IF(A1>0,B1,C1)")]
    public void ReverseSignIsCanonicalAndInvolutive(string formula)
    {
        var reversed = _wrappers.ReverseSign(formula);
        var restored = _wrappers.ReverseSign(reversed.Formula!);

        Assert.True(reversed.IsSuccess, reversed.Message);
        Assert.Equal("=-(" + formula.Substring(1) + ")", reversed.Formula);
        Assert.True(restored.IsSuccess, restored.Message);
        Assert.Equal(formula, restored.Formula);
    }

    [Theory]
    [InlineData(1000, true, "=(A1+B1)/1000")]
    [InlineData(1000, false, "=(A1+B1)*1000")]
    [InlineData(1000000, true, "=(A1+B1)/1000000")]
    [InlineData(1000000, false, "=(A1+B1)*1000000")]
    public void UnitScaleAlwaysNamesItsExactPrecedenceSafeOperation(long scale, bool divide, string expected)
    {
        var result = _wrappers.Scale("=A1+B1", scale, divide);

        Assert.True(result.IsSuccess, result.Message);
        Assert.Equal(expected, result.Formula);
    }

    [Fact]
    public void UnitScaleDoesNotInferOrTogglePriorUnits()
    {
        var once = _wrappers.Scale("=A1", 1000, divide: true);
        var twice = _wrappers.Scale(once.Formula!, 1000, divide: true);

        Assert.Equal("=((A1)/1000)/1000", twice.Formula);
    }

    [Fact]
    public void GeneratedTranslationCorpusIsDeterministicAndRoundTripsByInverseDisplacement()
    {
        for (var index = 1; index <= 1000; index++)
        {
            var row = 100 + index;
            var column = 10 + (index % 100);
            var formula = "=" + ColumnName(column) + row + "+$B" + row + "+H$7+$D$8";
            var rowDelta = (index % 17) - 8;
            var columnDelta = (index % 13) - 6;

            var forward = _references.Translate(formula, rowDelta, columnDelta);
            var inverse = _references.Translate(forward.Formula!, -rowDelta, -columnDelta);

            Assert.True(forward.IsSuccess, forward.Message);
            Assert.True(inverse.IsSuccess, inverse.Message);
            Assert.Equal(formula, inverse.Formula);
        }
    }

    [Fact]
    public void TransformEngineIsSafeForConcurrentPureCoreUse()
    {
        var failures = new ConcurrentQueue<string>();
        Parallel.For(0, 500, index =>
        {
            var formula = "=A" + (index + 10) + "+SUM($B$2,C3)";
            var result = _references.Translate(formula, 2, 3);
            if (!result.IsSuccess || result.Formula is null)
            {
                failures.Enqueue(formula + ": " + result.RefusalCode);
            }
        });

        Assert.Empty(failures);
    }

    [Fact]
    public void RepresentativeTransformsStayWithinThePureCoreBudget()
    {
        var formula = "=" + string.Join("+", Enumerable.Range(1, 200).Select(index => "A" + index));
        var stopwatch = Stopwatch.StartNew();

        for (var index = 0; index < 500; index++)
        {
            var result = _references.Translate(formula, 1, 1);
            Assert.True(result.IsSuccess, result.Message);
        }

        stopwatch.Stop();
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(5), $"Transforms took {stopwatch.Elapsed}.");
    }

    private static string ColumnName(int column)
    {
        var value = string.Empty;
        while (column > 0)
        {
            column--;
            value = (char)('A' + (column % 26)) + value;
            column /= 26;
        }

        return value;
    }
}
