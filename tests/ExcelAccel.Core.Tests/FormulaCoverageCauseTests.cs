using ExcelAccel.Core.Formulas;
using Xunit;

namespace ExcelAccel.Core.Tests;

/// <summary>
/// The parser reports one limitation code but carries every cause. A caller
/// deciding whether a construct could hide a reference needs all of them,
/// because the first cause never rules the others out.
/// </summary>
public sealed class FormulaCoverageCauseTests
{
    [Fact]
    public void AFullyCoveredFormulaCarriesNoCause()
    {
        var document = Parse("=A1+B2");

        Assert.Equal(FormulaCoverageDisposition.Transform, document.Disposition);
        Assert.Null(document.LimitationCode);
        Assert.Empty(document.LimitationCodes);
    }

    [Theory]
    [InlineData("=SUM(Table1[Amount])", FormulaRefusalCodes.StructuredReferenceInspectOnly)]
    [InlineData("=A1#", FormulaRefusalCodes.DynamicArrayInspectOnly)]
    [InlineData("='[Other.xlsx]Model'!A1", FormulaRefusalCodes.ExternalReferenceInspectOnly)]
    [InlineData("=SUM(A1:C3 B1:B5)", FormulaRefusalCodes.IntersectionInspectOnly)]
    [InlineData("=A1,C1", FormulaRefusalCodes.UnionInspectOnly)]
    [InlineData("=Rate", FormulaRefusalCodes.NameInspectOnly)]
    public void EachCauseIsReportedOnItsOwn(string formula, string expected)
    {
        var document = Parse(formula);

        Assert.Equal(FormulaCoverageDisposition.InspectOnly, document.Disposition);
        Assert.Equal(expected, document.LimitationCode);
        Assert.Contains(expected, document.LimitationCodes);
    }

    [Fact]
    public void EveryCauseIsRetainedWhenAFormulaCarriesSeveral()
    {
        var document = Parse("='[Other.xlsx]Model'!A1+SUM(A1:C3 B1:B5)+Rate");

        Assert.Equal(FormulaRefusalCodes.ExternalReferenceInspectOnly, document.LimitationCode);
        Assert.Contains(FormulaRefusalCodes.IntersectionInspectOnly, document.LimitationCodes);
        Assert.Contains(FormulaRefusalCodes.NameInspectOnly, document.LimitationCodes);
    }

    [Fact]
    public void TheReportedCodeIsAlwaysTheFirstCauseInPrecedenceOrder()
    {
        var document = Parse("=SUM(Table1[Amount])+'[Other.xlsx]Model'!A1");

        Assert.Equal(FormulaRefusalCodes.StructuredReferenceInspectOnly, document.LimitationCode);
        Assert.Equal(FormulaRefusalCodes.StructuredReferenceInspectOnly, document.LimitationCodes[0]);
        Assert.Contains(FormulaRefusalCodes.ExternalReferenceInspectOnly, document.LimitationCodes);
    }

    [Fact]
    public void CausesAreDistinctAndOrderedConsistently()
    {
        var document = Parse("='[Other.xlsx]Model'!A1+'[Second.xlsx]Model'!B2");

        Assert.Single(document.LimitationCodes, code => code == FormulaRefusalCodes.ExternalReferenceInspectOnly);
    }

    private static FormulaSyntaxDocument Parse(string formula)
    {
        var result = new FormulaParser().Parse(formula, new FormulaParseOptions(FormulaDialect.InvariantA1));
        Assert.True(result.IsSuccess, formula + " => " + result.Message);
        return result.Document!;
    }
}
