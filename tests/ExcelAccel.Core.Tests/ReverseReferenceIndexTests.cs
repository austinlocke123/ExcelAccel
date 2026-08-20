using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.Formulas;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class ReverseReferenceIndexTests
{
    private const string Workbook = "Book.xlsx";
    private const string Sheet = "Model";
    private static readonly DependentScanScope WorksheetScope = DependentScanScope.Worksheet(Workbook, Sheet);

    [Fact]
    public void CellAndRangeReferencesBothProduceDependentsWithRetainedEvidence()
    {
        var index = Build(("B1", "=A1*2"), ("C1", "=SUM(A1:A10)"), ("D1", "=B1+1"));

        var result = index.FindDirectDependents(Cell("A1"));

        Assert.Equal(AuditTraceStatus.Complete, result.Status);
        Assert.True(result.CanClaimCompleteness);
        Assert.Equal("worksheet", result.ScanScope);
        Assert.Equal(new[] { "Book.xlsx|Model|B1", "Book.xlsx|Model|C1" },
            result.Dependents.Select(value => value.Dependent.ToString()));
        Assert.Equal("A1", result.Dependents[0].Evidence[0].SourceText);
        Assert.Equal(AuditReferenceKind.Cell, result.Dependents[0].Evidence[0].Kind);
        Assert.Equal("A1:A10", result.Dependents[1].Evidence[0].SourceText);
        Assert.Equal(AuditReferenceKind.Range, result.Dependents[1].Evidence[0].Kind);
        Assert.All(result.Dependents, dependent => Assert.Equal(1, dependent.Depth));
    }

    [Fact]
    public void ARangeTargetMatchesAnyReferenceThatOverlapsPartOfIt()
    {
        var index = Build(("B1", "=SUM(A5:A20)"), ("B2", "=A30"), ("B3", "=SUM(A1:A3)"));

        var result = index.FindDirectDependents(Cell("A10:A25"));

        Assert.Equal(new[] { "Book.xlsx|Model|B1" }, result.Dependents.Select(value => value.Dependent.ToString()));
    }

    [Fact]
    public void ADefinedNameBindingProducesADependent()
    {
        var names = new[] { new AuditNameBinding("Rate", AuditNameScope.Workbook, Cell("A1")) };
        var index = ReverseReferenceIndex.Build(WorksheetScope, Formulas(("B1", "=Rate*2")), names);

        var result = index.FindDirectDependents(Cell("A1"));

        var dependent = Assert.Single(result.Dependents);
        Assert.Equal("Book.xlsx|Model|B1", dependent.Dependent.ToString());
        Assert.Equal(AuditReferenceKind.Name, dependent.Evidence[0].Kind);
        Assert.Equal("Rate", dependent.Evidence[0].SourceText);
    }

    [Fact]
    public void AnUnboundNameIsACoverageGapAndBlocksCompleteness()
    {
        var index = Build(("B1", "=Rate*2"));

        var result = index.FindDirectDependents(Cell("A1"));

        Assert.Equal(AuditTraceStatus.Partial, result.Status);
        Assert.False(result.CanClaimCompleteness);
        Assert.Equal(1, result.CoverageGapCount);
        Assert.Empty(result.Dependents);
    }

    [Fact]
    public void AnInspectOnlyFormulaKeepsItsResolvableEdgeAndStillBlocksCompleteness()
    {
        var index = Build(("B1", "=A1+SUM(Table1[Amount])"));

        var result = index.FindDirectDependents(Cell("A1"));

        Assert.Equal(AuditTraceStatus.Partial, result.Status);
        Assert.Equal(1, result.CoverageGapCount);
        Assert.Single(result.Dependents);
    }

    [Fact]
    public void ASpillReferenceIsACoverageGapBecauseItsExtentIsUnknowable()
    {
        var index = Build(("B1", "=A1#"));

        var result = index.FindDirectDependents(Cell("A1"));

        Assert.Equal(1, result.CoverageGapCount);
        Assert.False(result.CanClaimCompleteness);
    }

    [Fact]
    public void AnExternalReferenceNeverCountsAsAnInScopeDependent()
    {
        var index = Build(("B1", "='[Other.xlsx]Model'!A1"));

        var result = index.FindDirectDependents(Cell("A1"));

        Assert.Empty(result.Dependents);
    }

    /// <summary>
    /// An external reference addresses another workbook, so it can never conceal
    /// a reference to an in-scope cell and must not block the completeness claim.
    /// This matters in practice: external links are common in the models this
    /// add-in targets, and treating them as gaps made the completeness signal
    /// read "not claimed" in nearly every real workbook.
    /// </summary>
    [Fact]
    public void AnExternalReferenceDoesNotBlockTheCompletenessClaim()
    {
        var index = Build(("B1", "='[Other.xlsx]Model'!A1+A1"));

        var result = index.FindDirectDependents(Cell("A1"));

        Assert.Single(result.Dependents);
        Assert.Equal(0, result.CoverageGapCount);
        Assert.True(result.CanClaimCompleteness);
    }

    [Fact]
    public void AUnionReadsEveryOperandSoItDoesNotBlockCompleteness()
    {
        var index = Build(("B1", "=A1,C1"));

        var result = index.FindDirectDependents(Cell("A1"));

        Assert.Single(result.Dependents);
        Assert.Equal(0, result.CoverageGapCount);
        Assert.True(result.CanClaimCompleteness);
    }

    [Fact]
    public void AResolvedNameDoesNotBlockCompletenessEvenThoughTheParserNotesIt()
    {
        var names = new[] { new AuditNameBinding("Rate", AuditNameScope.Workbook, Cell("A1")) };
        var index = ReverseReferenceIndex.Build(WorksheetScope, Formulas(("B1", "=Rate*2")), names);

        var result = index.FindDirectDependents(Cell("A1"));

        Assert.Single(result.Dependents);
        Assert.Equal(0, result.CoverageGapCount);
        Assert.True(result.CanClaimCompleteness);
    }

    /// <summary>
    /// A structured reference resolves to real cells but produces no parsed
    /// reference, so it is a genuine blind spot and must stay a gap.
    /// </summary>
    [Fact]
    public void AStructuredReferenceRemainsAGapEvenBesideAnExternalReference()
    {
        var index = Build(("B1", "='[Other.xlsx]Model'!A1+SUM(Table1[Amount])"));

        var result = index.FindDirectDependents(Cell("A1"));

        Assert.Equal(1, result.CoverageGapCount);
        Assert.False(result.CanClaimCompleteness);
    }

    /// <summary>
    /// The parser reports only its first limitation and checks external before
    /// intersection, so a formula carrying both must still be a gap. This is the
    /// case that made the old first-cause-only reading unsafe to refine.
    /// </summary>
    [Fact]
    public void AnIntersectionBesideAnExternalReferenceIsStillAGap()
    {
        var index = Build(("B1", "='[Other.xlsx]Model'!A1+SUM(A1:C3 B1:B5)"));

        var result = index.FindDirectDependents(Cell("A1"));

        Assert.Equal(1, result.CoverageGapCount);
        Assert.False(result.CanClaimCompleteness);
    }

    [Fact]
    public void AFormulaOnAnotherWorksheetIsCountedAsAGapRatherThanRead()
    {
        var outside = new AuditFormulaCell(new AuditCellIdentity(Workbook, "Other", "B1"), "=Model!A1");
        var index = ReverseReferenceIndex.Build(WorksheetScope, Formulas(("B2", "=A1")).Concat(new[] { outside }));

        var result = index.FindDirectDependents(Cell("A1"));

        Assert.Equal(new[] { "Book.xlsx|Model|B2" }, result.Dependents.Select(value => value.Dependent.ToString()));
        Assert.Equal(1, result.CoverageGapCount);
        Assert.False(result.CanClaimCompleteness);
    }

    [Fact]
    public void WorkbookScopeIsRefusedWithAStableCodeRatherThanSilentlyWidened()
    {
        var index = ReverseReferenceIndex.Build(DependentScanScope.Workbook(Workbook), Formulas(("B1", "=A1")));

        var result = index.FindDirectDependents(Cell("A1"));

        Assert.Equal(AuditTraceStatus.Refused, result.Status);
        Assert.Equal(AuditRefusalCodes.ScopeUnsupported, result.RefusalCode);
        Assert.Empty(result.Dependents);
    }

    [Fact]
    public void ATargetOutsideTheDeclaredScopeIsRefused()
    {
        var index = Build(("B1", "=A1"));

        var result = index.FindDirectDependents(new AuditCellIdentity(Workbook, "Other", "A1"));

        Assert.Equal(AuditTraceStatus.Refused, result.Status);
        Assert.Equal(AuditRefusalCodes.TargetOutsideScope, result.RefusalCode);
    }

    [Fact]
    public void AnUnsupportedTargetNotationIsRefused()
    {
        var index = Build(("B1", "=A1"));

        var result = index.FindDirectDependents(new AuditCellIdentity(Workbook, Sheet, "A:A"));

        Assert.Equal(AuditTraceStatus.Refused, result.Status);
        Assert.Equal(AuditRefusalCodes.NotationUnsupported, result.RefusalCode);
    }

    [Fact]
    public void ExceedingTheScanCapTruncatesExplicitlyInsteadOfRunningOn()
    {
        var formulas = Enumerable.Range(1, ReverseReferenceIndex.MaximumScannedFormulas + 50)
            .Select(row => new AuditFormulaCell(
                new AuditCellIdentity(Workbook, Sheet, "B" + row.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                "=A1"))
            .ToArray();

        var result = ReverseReferenceIndex.Build(WorksheetScope, formulas).FindDirectDependents(Cell("A1"));

        Assert.True(result.Truncated);
        Assert.Equal(AuditTraceStatus.Partial, result.Status);
        Assert.False(result.CanClaimCompleteness);
        Assert.Equal(AuditRefusalCodes.ScanTruncated, result.LimitationCode);
        Assert.Equal(ReverseReferenceIndex.MaximumScannedFormulas, result.ScannedFormulaCount);
    }

    [Fact]
    public void DependentOrderAndEvidenceAreDeterministicAcrossRepeatBuilds()
    {
        var cells = new[] { ("D4", "=A1"), ("B2", "=A1:B2"), ("C3", "=SUM(A1)") };

        var first = Build(cells).FindDirectDependents(Cell("A1"));
        var second = Build(cells.Reverse().ToArray()).FindDirectDependents(Cell("A1"));

        Assert.Equal(
            first.Dependents.Select(value => value.Dependent.ToString()),
            second.Dependents.Select(value => value.Dependent.ToString()));
        Assert.Equal(new[] { "Book.xlsx|Model|B2", "Book.xlsx|Model|C3", "Book.xlsx|Model|D4" },
            first.Dependents.Select(value => value.Dependent.ToString()));
    }

    [Fact]
    public void ACircularSelfReferenceIsReportedRatherThanHidden()
    {
        var index = Build(("A1", "=A1+1"));

        var result = index.FindDirectDependents(Cell("A1"));

        Assert.Equal("Book.xlsx|Model|A1", Assert.Single(result.Dependents).Dependent.ToString());
    }

    [Fact]
    public void IndexedResultsMatchTheBruteForceOracleAcrossTheCorpus()
    {
        var corpus = new[]
        {
            ("B1", "=A1"), ("B2", "=$A$1+A2"), ("B3", "=SUM(A1:A10)"), ("B4", "=SUM(Z1:AB5)"),
            ("B5", "=A26*2"), ("B6", "=SUM(A20:C40)"), ("B7", "=Model!A1"), ("B8", "='Model'!B2:D4"),
            ("B9", "=A1+A1"), ("B10", "=IF(A5>0,A6,A7)"), ("B11", "=COUNT(AZ1:BA2)"), ("B12", "=ZZ700"),
        };
        var index = Build(corpus);
        var targets = new[] { "A1", "A2", "A5", "A26", "Z1", "AA3", "AZ1", "ZZ700", "B2:D4", "A20:C40", "XFD1048576" };

        foreach (var address in targets)
        {
            var target = Cell(address);
            var actual = index.FindDirectDependents(target).Dependents
                .Select(value => value.Dependent.Address).OrderBy(value => value, StringComparer.Ordinal).ToArray();
            var expected = BruteForceOracle.DirectDependents(corpus, address)
                .OrderBy(value => value, StringComparer.Ordinal).ToArray();

            Assert.Equal(expected, actual);
        }
    }

    [Fact]
    public void RepeatedQueriesNeverReparseAndStayConsistent()
    {
        var index = Build(("B1", "=A1"), ("C1", "=A1"));

        var first = index.FindDirectDependents(Cell("A1"));
        var second = index.FindDirectDependents(Cell("A1"));

        Assert.Equal(first.ScannedFormulaCount, second.ScannedFormulaCount);
        Assert.Equal(
            first.Dependents.Select(value => value.Dependent.ToString()),
            second.Dependents.Select(value => value.Dependent.ToString()));
    }

    [Fact]
    public void BuildRejectsNullArguments()
    {
        Assert.Throws<ArgumentNullException>(() => ReverseReferenceIndex.Build(null!, Array.Empty<AuditFormulaCell>()));
        Assert.Throws<ArgumentNullException>(() => ReverseReferenceIndex.Build(WorksheetScope, null!));
    }

    private static AuditCellIdentity Cell(string address) => new AuditCellIdentity(Workbook, Sheet, address);

    private static IEnumerable<AuditFormulaCell> Formulas(params (string Address, string Formula)[] cells) =>
        cells.Select(cell => new AuditFormulaCell(Cell(cell.Address), cell.Formula));

    private static ReverseReferenceIndex Build(params (string Address, string Formula)[] cells) =>
        ReverseReferenceIndex.Build(WorksheetScope, Formulas(cells));

    /// <summary>
    /// Independent reference implementation for AC-AUD-007. It deliberately
    /// avoids the production rectangle arithmetic: it expands every reference
    /// into an explicit set of cell addresses and tests set intersection.
    /// </summary>
    private static class BruteForceOracle
    {
        public static IEnumerable<string> DirectDependents(
            IReadOnlyList<(string Address, string Formula)> corpus,
            string targetAddress)
        {
            var targetCells = ExpandAddress(targetAddress);
            foreach (var (address, formula) in corpus)
            {
                var parse = new FormulaParser().Parse(formula, new FormulaParseOptions(FormulaDialect.InvariantA1));
                if (!parse.IsSuccess) continue;
                foreach (var reference in parse.Document!.References)
                {
                    if (reference.Qualifier is not null && reference.Qualifier.IndexOf('[') >= 0) continue;
                    if (reference.HasSpillOperator || reference.HasImplicitIntersection) continue;
                    var covered = Expand(reference);
                    if (covered.Overlaps(targetCells))
                    {
                        yield return address.ToUpperInvariant();
                        break;
                    }
                }
            }
        }

        private static HashSet<string> Expand(FormulaReference reference)
        {
            var firstRow = reference.First.Row.Value;
            var firstColumn = reference.First.Column.Value;
            var lastRow = reference.Second?.Row.Value ?? firstRow;
            var lastColumn = reference.Second?.Column.Value ?? firstColumn;
            return Enumerate(
                Math.Min(firstRow, lastRow), Math.Max(firstRow, lastRow),
                Math.Min(firstColumn, lastColumn), Math.Max(firstColumn, lastColumn));
        }

        private static HashSet<string> ExpandAddress(string address)
        {
            var parts = address.Split(':');
            var (firstRow, firstColumn) = Split(parts[0]);
            var (lastRow, lastColumn) = parts.Length == 1 ? (firstRow, firstColumn) : Split(parts[1]);
            return Enumerate(
                Math.Min(firstRow, lastRow), Math.Max(firstRow, lastRow),
                Math.Min(firstColumn, lastColumn), Math.Max(firstColumn, lastColumn));
        }

        private static HashSet<string> Enumerate(int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            var cells = new HashSet<string>(StringComparer.Ordinal);
            // The corpus is bounded, but guard against an accidental sweep.
            Assert.InRange((long)(lastRow - firstRow + 1) * (lastColumn - firstColumn + 1), 1, 10_000);
            for (var row = firstRow; row <= lastRow; row++)
            {
                for (var column = firstColumn; column <= lastColumn; column++)
                {
                    cells.Add(column.ToString(System.Globalization.CultureInfo.InvariantCulture) + "/" +
                        row.ToString(System.Globalization.CultureInfo.InvariantCulture));
                }
            }

            return cells;
        }

        private static (int Row, int Column) Split(string endpoint)
        {
            var text = endpoint.Replace("$", string.Empty);
            var index = 0;
            while (index < text.Length && char.IsLetter(text[index])) index++;
            var column = text.Substring(0, index).Aggregate(0, (total, letter) => (total * 26) + (char.ToUpperInvariant(letter) - 'A' + 1));
            return (int.Parse(text.Substring(index), System.Globalization.CultureInfo.InvariantCulture), column);
        }
    }
}
