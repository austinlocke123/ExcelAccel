using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Core.Compare;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class CellComparisonTests
{
    private static ComparisonCell Formula(string address, string formula, string numberFormat = "General") =>
        new ComparisonCell(address, formula, string.Empty, numberFormat);

    private static ComparisonCell Constant(string address, string value, string numberFormat = "General", string displayed = "") =>
        new ComparisonCell(address, string.Empty, value, numberFormat, displayed);

    private static ComparisonBlock Block(string workbook, params ComparisonCell[] cells) =>
        new ComparisonBlock(workbook, "Sheet1", "A1:A" + cells.Length, cells.Length, 1, cells);

    [Fact]
    public void IdenticalBlocksProduceNoDifferences()
    {
        var result = SameShapeComparer.Compare(
            Block("Left.xlsx", Formula("A1", "=B1+1"), Constant("A2", "42")),
            Block("Right.xlsx", Formula("A1", "=B1+1"), Constant("A2", "42")));

        Assert.Empty(result.Differences);
        Assert.Equal(2, result.ComparedCells);
    }

    /// <summary>
    /// The whole point of "same shape": no alignment, no shift detection, no
    /// truncation to a common size.
    /// </summary>
    [Fact]
    public void DifferentShapesRefuseRatherThanAlign()
    {
        var error = Assert.Throws<ShapeMismatchException>(() => SameShapeComparer.Compare(
            Block("Left.xlsx", Constant("A1", "1"), Constant("A2", "2")),
            Block("Right.xlsx", Constant("A1", "1"))));

        Assert.Contains("does not align or truncate", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AConstantDifferenceIsReported()
    {
        var result = SameShapeComparer.Compare(
            Block("Left.xlsx", Constant("A1", "42")),
            Block("Right.xlsx", Constant("A1", "43")));

        var difference = Assert.Single(result.Differences);
        Assert.Equal(DifferenceCategory.Constant, difference.Category);
        Assert.Equal(0, difference.Row);
    }

    /// <summary>
    /// A formula on one side and a constant on the other is its own category, not
    /// forced into either, because "the formulas differ" would be false.
    /// </summary>
    [Fact]
    public void AFormulaAgainstAConstantIsATypeMismatch()
    {
        var result = SameShapeComparer.Compare(
            Block("Left.xlsx", Formula("A1", "=B1")),
            Block("Right.xlsx", Constant("A1", "5")));

        var difference = Assert.Single(result.Differences);
        Assert.Equal(DifferenceCategory.TypeMismatch, difference.Category);
        Assert.Contains("Formula on the source", difference.Statement, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("=B1+1", "=B1+1", FormulaDifferenceKind.Identical)]
    [InlineData("=B1+1", "=C1+1", FormulaDifferenceKind.EquivalentShape)]
    [InlineData("=SUM(B1:B9)", "=SUM(C1:C9)", FormulaDifferenceKind.EquivalentShape)]
    [InlineData("=B1+1", "=B1+2", FormulaDifferenceKind.Structure)]
    [InlineData("=B1+1", "=SUM(B1:B9)", FormulaDifferenceKind.Structure)]
    [InlineData("=B1+1", "=((((", FormulaDifferenceKind.OutsideCoverage)]
    public void FormulaDifferencesDistinguishTextStructureAndCoverage(string left, string right, FormulaDifferenceKind expected) =>
        Assert.Equal(expected, SameShapeComparer.CompareFormulas(left, right));

    /// <summary>
    /// A copied formula reads differently but is the same shape from another
    /// anchor. Calling that a structural difference would bury the real ones.
    /// </summary>
    [Fact]
    public void ACopiedFormulaIsReportedAsSameShapeRatherThanAStructureChange()
    {
        var result = SameShapeComparer.Compare(
            Block("Left.xlsx", Formula("A1", "=B1*2")),
            Block("Right.xlsx", Formula("A1", "=C5*2")));

        var difference = Assert.Single(result.Differences);
        Assert.Equal(FormulaDifferenceKind.EquivalentShape, difference.FormulaKind);
        Assert.Contains("Same shape", difference.Statement, StringComparison.Ordinal);
    }

    [Fact]
    public void AFormulaOutsideCoverageIsNamedAsAGapRatherThanGuessedAt()
    {
        var result = SameShapeComparer.Compare(
            Block("Left.xlsx", Formula("A1", "=((((")),
            Block("Right.xlsx", Formula("A1", "=)))))")));

        Assert.Contains("formula_outside_parser_coverage", result.CoverageGaps);
        Assert.False(result.IsComplete);
        Assert.Equal(FormulaDifferenceKind.OutsideCoverage, result.Differences.Single().FormulaKind);
    }

    /// <summary>
    /// A category that was not examined must never look like a category that was
    /// examined and found equal.
    /// </summary>
    [Fact]
    public void UnexaminedCategoriesAreNamedInTheCoverageGaps()
    {
        var result = SameShapeComparer.Compare(
            Block("Left.xlsx", Constant("A1", "1", "0.00", "1.00")),
            Block("Right.xlsx", Constant("A1", "1", "0.0000", "1.0000")),
            new ComparisonCategories(formulas: true, constants: true, numberFormats: false, displayedValues: false));

        Assert.Empty(result.Differences);
        Assert.False(result.IsComplete);
        Assert.Contains("category_not_compared:number_format", result.CoverageGaps);
        Assert.Contains("category_not_compared:displayed_value", result.CoverageGaps);
    }

    [Fact]
    public void RequestedNumberFormatAndDisplayedValueDifferencesAreFound()
    {
        var categories = new ComparisonCategories(true, true, numberFormats: true, displayedValues: true);

        var formats = SameShapeComparer.Compare(
            Block("Left.xlsx", Constant("A1", "1", "0.00")),
            Block("Right.xlsx", Constant("A1", "1", "0.0000")),
            categories);
        Assert.Equal(DifferenceCategory.NumberFormat, formats.Differences.Single().Category);

        var displayed = SameShapeComparer.Compare(
            Block("Left.xlsx", Constant("A1", "1", "0.00", "1.00")),
            Block("Right.xlsx", Constant("A1", "1", "0.00", "1.0000")),
            categories);
        Assert.Equal(DifferenceCategory.DisplayedValue, displayed.Differences.Single().Category);
        Assert.True(displayed.IsComplete);
    }

    [Fact]
    public void TheFingerprintCoversBothSidesSoEitherEditInvalidatesIt()
    {
        var left = Block("Left.xlsx", Constant("A1", "1"));
        var right = Block("Right.xlsx", Constant("A1", "1"));
        var baseline = SameShapeComparer.Compare(left, right).Fingerprint;

        var sourceEdited = SameShapeComparer.Compare(Block("Left.xlsx", Constant("A1", "2")), right).Fingerprint;
        var targetEdited = SameShapeComparer.Compare(left, Block("Right.xlsx", Constant("A1", "2"))).Fingerprint;

        Assert.NotEqual(baseline, sourceEdited);
        Assert.NotEqual(baseline, targetEdited);
        Assert.NotEqual(sourceEdited, targetEdited);
    }

    [Fact]
    public void TheSameInputsProduceTheSameResultEveryTime()
    {
        var left = Block("Left.xlsx", Formula("A1", "=B1+1"), Constant("A2", "42"));
        var right = Block("Right.xlsx", Formula("A1", "=B1+2"), Constant("A2", "43"));

        var first = SameShapeComparer.Compare(left, right);
        var second = SameShapeComparer.Compare(left, right);

        Assert.Equal(first.Fingerprint, second.Fingerprint);
        Assert.Equal(
            first.Differences.Select(d => $"{d.Row}:{d.Column}:{d.Category}:{d.FormulaKind}"),
            second.Differences.Select(d => $"{d.Row}:{d.Column}:{d.Category}:{d.FormulaKind}"));
    }

    [Fact]
    public void ExceedingTheDifferenceBoundIsReportedRatherThanSilentlyTruncated()
    {
        var count = SameShapeComparer.MaximumDifferences + 10;
        var left = new ComparisonBlock("Left.xlsx", "Sheet1", "A1", count, 1,
            Enumerable.Range(1, count).Select(row => Constant("A" + row, row.ToString())));
        var right = new ComparisonBlock("Right.xlsx", "Sheet1", "A1", count, 1,
            Enumerable.Range(1, count).Select(row => Constant("A" + row, (row + 1000).ToString())));

        var result = SameShapeComparer.Compare(left, right);

        Assert.Equal(SameShapeComparer.MaximumDifferences, result.Differences.Count);
        Assert.Equal(10, result.ExcludedByBound);
        Assert.False(result.IsComplete);
        Assert.Contains("difference_count_exceeded_qualified_bound", result.CoverageGaps);
        // The count is of every difference found, not only those retained.
        Assert.Equal(count, result.Counts[DifferenceCategory.Constant]);
    }

    [Fact]
    public void ExceedingTheCellBoundIsRefused()
    {
        var count = SameShapeComparer.MaximumCells + 1;
        var cells = Enumerable.Range(1, count).Select(row => Constant("A" + row, "1")).ToArray();
        var left = new ComparisonBlock("Left.xlsx", "Sheet1", "A1", count, 1, cells);
        var right = new ComparisonBlock("Right.xlsx", "Sheet1", "A1", count, 1, cells);

        Assert.Throws<ComparisonBoundException>(() => SameShapeComparer.Compare(left, right));
    }

    [Fact]
    public void ABlockRequiresExactlyOneCellPerPosition() =>
        Assert.Throws<ArgumentException>(() =>
            new ComparisonBlock("Left.xlsx", "Sheet1", "A1:B2", 2, 2, new[] { Constant("A1", "1") }));

    [Fact]
    public void PositionsAreReportedRowMajor()
    {
        var left = new ComparisonBlock("Left.xlsx", "Sheet1", "A1:B2", 2, 2, new[]
        {
            Constant("A1", "1"), Constant("B1", "2"),
            Constant("A2", "3"), Constant("B2", "4"),
        });
        var right = new ComparisonBlock("Right.xlsx", "Sheet1", "A1:B2", 2, 2, new[]
        {
            Constant("A1", "1"), Constant("B1", "9"),
            Constant("A2", "3"), Constant("B2", "9"),
        });

        var result = SameShapeComparer.Compare(left, right);

        Assert.Equal(new[] { (0, 1), (1, 1) }, result.Differences.Select(d => (d.Row, d.Column)));
    }
}
