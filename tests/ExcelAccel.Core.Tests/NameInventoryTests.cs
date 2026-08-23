using System.Linq;
using ExcelAccel.Core.Names;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class NameInventoryTests
{
    private static NameRecord Workbook(string name, string refersTo, bool visible = true, bool builtIn = false) =>
        new NameRecord(name, NameScopeKind.Workbook, null, refersTo, visible, builtIn);

    private static NameRecord Sheet(string name, string worksheet, string refersTo, bool visible = true) =>
        new NameRecord(name, NameScopeKind.Worksheet, worksheet, refersTo, visible, false);

    [Theory]
    [InlineData("=Sheet1!$A$1", NameTargetKind.Range)]
    [InlineData("=Sheet1!$A$1:$B$9", NameTargetKind.Range)]
    [InlineData("=42", NameTargetKind.Constant)]
    [InlineData("=-1", NameTargetKind.Constant)]
    [InlineData("=\"Region\"", NameTargetKind.Constant)]
    [InlineData("=Sheet1!$A$1*2", NameTargetKind.Formula)]
    [InlineData("=SUM(Sheet1!$A$1:$A$9)", NameTargetKind.Formula)]
    [InlineData("='[Other.xlsx]Sheet1'!$A$1", NameTargetKind.External)]
    [InlineData("=Sheet1!#REF!", NameTargetKind.Broken)]
    [InlineData("", NameTargetKind.Unresolved)]
    public void TargetsAreClassifiedFromLocalMetadataAlone(string refersTo, NameTargetKind expected) =>
        Assert.Equal(expected, NameInventoryBuilder.Classify(refersTo));

    /// <summary>
    /// A broken external reference is still broken; the #REF! check has to win.
    /// </summary>
    [Fact]
    public void ABrokenExternalTargetReportsAsBrokenRatherThanExternal() =>
        Assert.Equal(NameTargetKind.Broken, NameInventoryBuilder.Classify("='[Other.xlsx]Sheet1'!#REF!"));

    /// <summary>
    /// A real-Excel run found _xlfn.SINGLE in a fresh workbook: a function shim
    /// Excel adds for itself, which has no business in a list of the user's
    /// named ranges.
    /// </summary>
    [Theory]
    [InlineData("_xlfn.SINGLE", true)]
    [InlineData("_xlfn.ANCHORARRAY", true)]
    [InlineData("_xludf.MyUdf", true)]
    [InlineData("_xlchart.v1.0", true)]
    [InlineData("Print_Area", true)]
    [InlineData("_FilterDatabase", true)]
    [InlineData("print_titles", true)]
    [InlineData("GrowthRate", false)]
    [InlineData("_MyOwnName", false)]
    [InlineData("", false)]
    public void ExcelOwnedNamesAreRecognisedAsReserved(string localName, bool expected) =>
        Assert.Equal(expected, ReservedNames.IsReserved(localName));

    [Fact]
    public void HiddenAndBuiltInNamesAreExcludedUnlessAskedFor()
    {
        var records = new[]
        {
            Workbook("Visible", "=Sheet1!$A$1"),
            Workbook("Concealed", "=Sheet1!$A$2", visible: false),
            Workbook("Print_Area", "=Sheet1!$A$1:$B$9", builtIn: true),
        };

        Assert.Single(NameInventoryBuilder.Build(records, includeHidden: false, includeBuiltIn: false).Entries);
        Assert.Equal(2, NameInventoryBuilder.Build(records, includeHidden: true, includeBuiltIn: false).Entries.Count);
        Assert.Equal(3, NameInventoryBuilder.Build(records, includeHidden: true, includeBuiltIn: true).Entries.Count);
    }

    /// <summary>
    /// Excel permits the same display name in more than one scope, and it is a
    /// common reason a formula resolves to a different name than intended.
    /// </summary>
    [Fact]
    public void TheSameNameInTwoScopesIsFlaggedOnBothEntries()
    {
        var inventory = NameInventoryBuilder.Build(
            new[]
            {
                Workbook("Rate", "=Sheet1!$A$1"),
                Sheet("Rate", "Sheet2", "=Sheet2!$A$1"),
                Workbook("Unique", "=Sheet1!$A$2"),
            },
            includeHidden: true,
            includeBuiltIn: true);

        Assert.All(
            inventory.Entries.Where(entry => entry.Record.Name == "Rate"),
            entry => Assert.True(entry.HasDuplicateDisplayName));
        Assert.False(inventory.Entries.Single(entry => entry.Record.Name == "Unique").HasDuplicateDisplayName);
    }

    [Fact]
    public void IdentitySeparatesScopeWorksheetAndNameSoItCannotCollide()
    {
        var first = NameInventoryBuilder.IdFor(Sheet("C", "AB", "=Sheet1!$A$1"));
        var second = NameInventoryBuilder.IdFor(Sheet("BC", "A", "=Sheet1!$A$1"));

        Assert.NotEqual(first, second);
    }

    [Fact]
    public void OrderingIsDeterministicRegardlessOfCaptureOrder()
    {
        var records = new[]
        {
            Sheet("zeta", "Sheet2", "=Sheet2!$A$1"),
            Workbook("alpha", "=Sheet1!$A$1"),
            Sheet("Alpha", "Sheet1", "=Sheet1!$A$2"),
        };

        var forward = NameInventoryBuilder.Build(records, true, true).Entries.Select(entry => entry.Id);
        var reversed = NameInventoryBuilder.Build(records.Reverse().ToArray(), true, true).Entries.Select(entry => entry.Id);

        Assert.Equal(forward, reversed);
        Assert.Equal("alpha", NameInventoryBuilder.Build(records, true, true).Entries[0].Record.Name);
    }

    /// <summary>
    /// Non-navigable names stay in the inventory with a stated reason. Dropping
    /// them would hide exactly the names a reviewer most wants to see.
    /// </summary>
    [Theory]
    [InlineData("=42")]
    [InlineData("='[Other.xlsx]Sheet1'!$A$1")]
    [InlineData("=Sheet1!#REF!")]
    public void ANonNavigableNameRemainsVisibleWithAReason(string refersTo)
    {
        var entry = NameInventoryBuilder
            .Build(new[] { Workbook("Target", refersTo) }, true, true)
            .Entries.Single();

        Assert.False(entry.IsNavigable);
        Assert.False(string.IsNullOrWhiteSpace(entry.NonNavigableReason));
    }

    [Fact]
    public void ARangeNameIsNavigable()
    {
        var entry = NameInventoryBuilder
            .Build(new[] { Workbook("Target", "=Sheet1!$A$1") }, true, true)
            .Entries.Single();

        Assert.True(entry.IsNavigable);
        Assert.Null(entry.NonNavigableReason);
    }

    [Fact]
    public void ExceedingTheBoundIsReportedRatherThanSilentlyTruncated()
    {
        var records = Enumerable.Range(0, NameInventoryBuilder.MaximumNames + 5)
            .Select(index => Workbook("Name" + index.ToString("D5"), "=Sheet1!$A$1"))
            .ToArray();

        var inventory = NameInventoryBuilder.Build(records, true, true);

        Assert.Equal(NameInventoryBuilder.MaximumNames, inventory.Entries.Count);
        Assert.Equal(5, inventory.ExcludedByBound);
        Assert.False(inventory.IsComplete);
        Assert.Contains("name_count_exceeded_qualified_bound", inventory.CoverageGaps);
        Assert.Contains("beyond the qualified bound", NameInventorySearch.CoverageStatement(inventory));
    }

    [Fact]
    public void AnEmptyInventoryReportsComplete()
    {
        var inventory = NameInventoryBuilder.Build(new NameRecord[0], true, true);

        Assert.True(inventory.IsComplete);
        Assert.Equal("Complete: 0 name(s).", NameInventorySearch.CoverageStatement(inventory));
    }

    [Fact]
    public void SearchFiltersDeterministicallyWithoutTouchingTheWorkbook()
    {
        var inventory = NameInventoryBuilder.Build(
            new[]
            {
                Workbook("GrowthRate", "=Sheet1!$A$1"),
                Workbook("Discount", "=42"),
                Sheet("GrowthCap", "Model", "=Model!$B$2"),
            },
            true,
            true);

        Assert.Equal(2, NameInventorySearch.Filter(inventory, "growth").Count);
        Assert.Single(NameInventorySearch.Filter(inventory, null, scope: NameScopeKind.Worksheet));
        Assert.Single(NameInventorySearch.Filter(inventory, null, targetKind: NameTargetKind.Constant));
        Assert.Equal(3, NameInventorySearch.Filter(inventory, "   ").Count);
    }

    [Fact]
    public void SearchCanIsolateBrokenNames()
    {
        var inventory = NameInventoryBuilder.Build(
            new[] { Workbook("Good", "=Sheet1!$A$1"), Workbook("Bad", "=Sheet1!#REF!") }, true, true);

        var broken = NameInventorySearch.Filter(inventory, null, brokenOnly: true).Single();

        Assert.Equal("Bad", broken.Record.Name);
    }

    /// <summary>
    /// A name's target can carry a path or a business term, so expressions are
    /// excluded from the export unless the user opts in.
    /// </summary>
    [Fact]
    public void ExportExcludesExpressionsUnlessRequested()
    {
        var entry = NameInventoryBuilder
            .Build(new[] { Workbook("Secret", "='[Budget.xlsx]Q3'!$A$1") }, true, true)
            .Entries.Single();

        var redacted = NameInventorySearch.ExportRow(entry, includeExpressions: false);
        var full = NameInventorySearch.ExportRow(entry, includeExpressions: true);

        Assert.DoesNotContain(redacted, field => field.Contains("Budget.xlsx"));
        Assert.Contains(full, field => field.Contains("Budget.xlsx"));
        Assert.Equal(NameInventorySearch.ExportFields(false).Count, redacted.Count);
        Assert.Equal(NameInventorySearch.ExportFields(true).Count, full.Count);
    }
}
