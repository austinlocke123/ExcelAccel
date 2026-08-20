using System;
using System.Linq;
using ExcelAccel.Application.Auditing;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.Formulas;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class FormulaTreeBuilderTests
{
    [Fact]
    public void AFunctionCallBecomesAFunctionNodeWithOneChildPerArgument()
    {
        var root = Root("=SUM(A1,B2,3)");

        Assert.Equal(FormulaNodeKind.Function, root.Kind);
        Assert.Equal("SUM", root.Text);
        Assert.Equal(3, root.Children.Count);
        Assert.Equal(
            new[] { FormulaNodeKind.Reference, FormulaNodeKind.Reference, FormulaNodeKind.Number },
            root.Children.Select(child => child.Kind));
    }

    [Fact]
    public void OperatorPrecedenceNestsMultiplicationBeneathAddition()
    {
        var root = Root("=1+2*3");

        Assert.Equal(FormulaNodeKind.BinaryOperator, root.Kind);
        Assert.Equal("+", root.Text);
        Assert.Equal("1", root.Children[0].Text);
        Assert.Equal("*", root.Children[1].Text);
        Assert.Equal(new[] { "2", "3" }, root.Children[1].Children.Select(child => child.Text));
    }

    [Fact]
    public void ExponentiationIsRightAssociative()
    {
        var root = Root("=2^3^4");

        Assert.Equal("^", root.Text);
        Assert.Equal("2", root.Children[0].Text);
        Assert.Equal("^", root.Children[1].Text);
    }

    [Fact]
    public void ComparisonBindsMoreLoselyThanArithmetic()
    {
        var root = Root("=A1+1>B1*2");

        Assert.Equal(">", root.Text);
        Assert.Equal("+", root.Children[0].Text);
        Assert.Equal("*", root.Children[1].Text);
    }

    [Fact]
    public void ParenthesesProduceAnExplicitGroupNode()
    {
        var root = Root("=(1+2)*3");

        Assert.Equal("*", root.Text);
        Assert.Equal(FormulaNodeKind.Group, root.Children[0].Kind);
        Assert.Equal("+", root.Children[0].Children[0].Text);
    }

    [Fact]
    public void UnaryAndPostfixOperatorsAreDistinctNodeKinds()
    {
        var negation = Root("=-A1");
        var percent = Root("=50%");

        Assert.Equal(FormulaNodeKind.UnaryOperator, negation.Kind);
        Assert.Equal(FormulaNodeKind.PostfixOperator, percent.Kind);
        Assert.Equal("50", percent.Children[0].Text);
    }

    [Fact]
    public void ConstantsReferencesAndNamesAreClassifiedDistinctly()
    {
        var root = Root("=IF(TRUE,\"x\",Rate)");

        Assert.Equal(
            new[] { FormulaNodeKind.Boolean, FormulaNodeKind.Text, FormulaNodeKind.Name },
            root.Children.Select(child => child.Kind));
    }

    [Fact]
    public void EveryNodeCarriesASpanThatQuotesItsOwnSource()
    {
        const string formula = "=SUM(A1,2)*3";
        var tree = Build(formula);

        Assert.All(tree.Root!.Flatten(), entry =>
        {
            var span = entry.Node.Span;
            Assert.InRange(span.Start, 0, formula.Length);
            Assert.InRange(span.Start + span.Length, 0, formula.Length);
        });
        var sum = tree.Root.Flatten().Single(entry => entry.Node.Kind == FormulaNodeKind.Function).Node;
        Assert.Equal("SUM(A1,2)", formula.Substring(sum.Span.Start, sum.Span.Length));
    }

    [Fact]
    public void FlatteningIsPreOrderSoFocusOrderWalksTheTree()
    {
        var tree = Build("=1+2*3");

        Assert.Equal(
            new[] { "+", "1", "*", "2", "3" },
            tree.Root!.Flatten().Select(entry => entry.Node.Text));
        Assert.Equal(new[] { 0, 1, 1, 2, 2 }, tree.Root.Flatten().Select(entry => entry.Depth));
    }

    /// <summary>
    /// An unsupported construct must name its category and span, and must not be
    /// presented as a fully parsed formula.
    /// </summary>
    [Theory]
    [InlineData("=SUM(Table1[Amount])", FormulaRefusalCodes.StructuredReferenceInspectOnly)]
    [InlineData("=A1#", FormulaRefusalCodes.DynamicArrayInspectOnly)]
    [InlineData("=SUM(A1:C3 B1:B5)", FormulaRefusalCodes.IntersectionInspectOnly)]
    [InlineData("=A1,B1", FormulaRefusalCodes.UnionInspectOnly)]
    public void AnUnsupportedConstructIsNamedWithItsSpanAndNotPresentedAsComplete(string formula, string expected)
    {
        var tree = Build(formula);

        Assert.False(tree.IsComplete);
        Assert.Equal(expected, tree.LimitationCode);
        Assert.NotNull(tree.LimitationSpan);
        Assert.Null(tree.Root);
    }

    [Fact]
    public void AParserRefusalIsReportedWithTheParsersOwnReason()
    {
        var result = new FormulaParser().Parse("=SUM(", new FormulaParseOptions(FormulaDialect.InvariantA1));
        Assert.False(result.IsSuccess);

        var tree = FormulaTreeResult.Refused(result.RefusalCode, result.Message);

        Assert.False(tree.IsComplete);
        Assert.Null(tree.Root);
        Assert.Equal(result.RefusalCode, tree.LimitationCode);
    }

    [Fact]
    public void AnExternalReferenceStillProducesATree()
    {
        var tree = Build("='[Other.xlsx]Data'!A1+1");

        Assert.True(tree.IsComplete);
        Assert.Equal("+", tree.Root!.Text);
    }

    [Fact]
    public void RepeatBuildsOfTheSameFormulaAreIdentical()
    {
        var first = Build("=IF(A1>0,SUM(B1:B9),-C1)");
        var second = Build("=IF(A1>0,SUM(B1:B9),-C1)");

        Assert.Equal(
            first.Root!.Flatten().Select(entry => entry.Node.Kind + ":" + entry.Node.Text + ":" + entry.Depth),
            second.Root!.Flatten().Select(entry => entry.Node.Kind + ":" + entry.Node.Text + ":" + entry.Depth));
    }

    [Fact]
    public void BuildRejectsANullDocument() =>
        Assert.Throws<ArgumentNullException>(() => FormulaTreeBuilder.Build(null!));

    internal static FormulaTreeResult Build(string formula)
    {
        var parse = new FormulaParser().Parse(formula, new FormulaParseOptions(FormulaDialect.InvariantA1));
        Assert.True(parse.IsSuccess, formula + " => " + parse.Message);
        return FormulaTreeBuilder.Build(parse.Document!);
    }

    private static FormulaSyntaxNode Root(string formula)
    {
        var tree = Build(formula);
        Assert.True(tree.IsComplete, formula + " => " + tree.Message);
        return tree.Root!;
    }
}

public sealed class FormulaInspectorReportTests
{
    private static readonly AuditCellIdentity Cell = new AuditCellIdentity("Book.xlsx", "Model", "D10");

    [Fact]
    public void ACompleteTreeProjectsIndentedRowsInFocusOrder()
    {
        const string formula = "=1+2*3";
        var report = FormulaInspectorReport.Create(Cell, formula, FormulaTreeBuilderTests.Build(formula));
        var presentation = report.ToPresentation();

        Assert.Equal("ExcelAccel Formula Inspector", presentation.Title);
        Assert.Equal(AuditTraceStatus.Complete, presentation.Status);
        Assert.Equal(5, presentation.Rows.Count);
        Assert.StartsWith("+", presentation.Rows[0].Values[0]);
        Assert.StartsWith("   ", presentation.Rows[1].Values[0]);
        Assert.All(presentation.Rows, row => Assert.Equal(presentation.Columns.Count, row.Values.Count));
    }

    [Fact]
    public void EachRowQuotesItsOwnSourceSpan()
    {
        const string formula = "=SUM(A1,2)";
        var presentation = FormulaInspectorReport
            .Create(Cell, formula, FormulaTreeBuilderTests.Build(formula))
            .ToPresentation();

        Assert.Contains(presentation.Rows, row => row.Values[3] == "SUM(A1,2)");
        Assert.Contains(presentation.Rows, row => row.Values[3] == "A1");
    }

    [Fact]
    public void AnUnsupportedFormulaIsRefusedWithItsCategoryAndSpan()
    {
        const string formula = "=SUM(Table1[Amount])";
        var report = FormulaInspectorReport.Create(Cell, formula, FormulaTreeBuilderTests.Build(formula));

        Assert.Equal(AuditTraceStatus.Refused, report.Status);
        Assert.Empty(report.ToPresentation().Rows);
        Assert.Contains(report.SummaryLines, line => line.StartsWith("Limitation code: ", StringComparison.Ordinal));
        Assert.Contains(report.SummaryLines, line => line.StartsWith("Limitation span: ", StringComparison.Ordinal));
    }

    /// <summary>
    /// The inspector renders structure. It never evaluates, scores, or explains.
    /// </summary>
    [Fact]
    public void TheReportStatesThatNothingIsEvaluatedOrScored()
    {
        var report = FormulaInspectorReport.Create(Cell, "=1+1", FormulaTreeBuilderTests.Build("=1+1"));

        Assert.Contains("No subexpression is evaluated, scored, or explained.", report.SummaryLines);
        Assert.DoesNotContain(report.SummaryLines, line =>
            line.IndexOf("complexity", StringComparison.OrdinalIgnoreCase) >= 0 ||
            line.IndexOf("score", StringComparison.OrdinalIgnoreCase) >= 0 && !line.StartsWith("No subexpression", StringComparison.Ordinal));
    }

    [Fact]
    public void TheRegisteredCommandIsReadOnlyWithInspectorAcceptance()
    {
        var descriptor = BuiltInCommandRegistry.GetRequired(AuditingCommandCatalog.InspectFormulaId);

        Assert.Equal(Core.Commands.CommandImpact.ReadOnly, descriptor.Impact);
        Assert.Empty(descriptor.ChangedProperties);
        Assert.Equal(UndoPolicy.None, descriptor.UndoPolicy);
        Assert.Equal("CAP-AUD-002", descriptor.CapabilityId);
        Assert.Equal(
            new[] { "AC-AUD-016", "AC-AUD-017", "AC-AUD-018", "AC-AUD-019" },
            descriptor.AcceptanceIds);
        Assert.Equal("Alt, X, A, AF", descriptor.ShortcutLabel);
    }

    [Fact]
    public void CreateRejectsNullArguments()
    {
        var tree = FormulaTreeBuilderTests.Build("=1");
        Assert.Throws<ArgumentNullException>(() => FormulaInspectorReport.Create(null!, "=1", tree));
        Assert.Throws<ArgumentNullException>(() => FormulaInspectorReport.Create(Cell, null!, tree));
        Assert.Throws<ArgumentNullException>(() => FormulaInspectorReport.Create(Cell, "=1", null!));
    }
}
