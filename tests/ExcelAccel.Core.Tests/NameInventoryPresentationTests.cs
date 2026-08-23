using System.Linq;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.Names;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class NameInventoryPresentationTests
{
    private static NameInventory Build(params NameRecord[] records) =>
        NameInventoryBuilder.Build(records, includeHidden: true, includeBuiltIn: true);

    private static NameRecord Workbook(string name, string refersTo) =>
        new NameRecord(name, NameScopeKind.Workbook, null, refersTo, true, false);

    [Fact]
    public void EveryRowSuppliesOneValuePerColumn()
    {
        var presentation = NameInventoryPresentation.Create(
            Build(Workbook("Rate", "=Sheet1!$A$1"), Workbook("Cap", "=42")), "Book.xlsx");

        Assert.Equal(2, presentation.Rows.Count);
        Assert.All(presentation.Rows, row => Assert.Equal(presentation.Columns.Count, row.Values.Count));
    }

    [Fact]
    public void ARangeNameIsNavigableToItsSheetAndAddress()
    {
        var presentation = NameInventoryPresentation.Create(Build(Workbook("Rate", "=Sheet1!$A$1")), "Book.xlsx");

        var target = presentation.Rows.Single().NavigationTarget;

        Assert.NotNull(target);
        Assert.Equal("Sheet1", target!.WorksheetName);
        Assert.Equal("A1", target.Address);
        Assert.Equal("Book.xlsx", target.WorkbookId);
    }

    [Fact]
    public void AQuotedSheetNameIsUnquotedForNavigation()
    {
        var presentation = NameInventoryPresentation.Create(
            Build(Workbook("Rate", "='Model Sheet'!$B$2:$C$9")), "Book.xlsx");

        var target = presentation.Rows.Single().NavigationTarget;

        Assert.NotNull(target);
        Assert.Equal("Model Sheet", target!.WorksheetName);
        Assert.Equal("B2:C9", target.Address);
    }

    /// <summary>
    /// A name that cannot be selected still appears, with the reason in its note.
    /// Dropping it would hide exactly the names a reviewer wants to find.
    /// </summary>
    [Theory]
    [InlineData("=42")]
    [InlineData("='[Other.xlsx]Sheet1'!$A$1")]
    [InlineData("=Sheet1!#REF!")]
    [InlineData("=SUM(Sheet1!$A$1:$A$9)")]
    public void ANonNavigableNameIsShownWithoutATarget(string refersTo)
    {
        var presentation = NameInventoryPresentation.Create(Build(Workbook("Target", refersTo)), "Book.xlsx");

        var row = presentation.Rows.Single();

        Assert.Null(row.NavigationTarget);
        Assert.False(row.IsNavigable);
        Assert.False(string.IsNullOrWhiteSpace(row.Values.Last()));
    }

    [Fact]
    public void AnUnqualifiedReferenceIsNotOfferedAsAJump()
    {
        // No sheet qualifier means no unambiguous target to select.
        var presentation = NameInventoryPresentation.Create(Build(Workbook("Local", "=$A$1")), "Book.xlsx");

        Assert.Null(presentation.Rows.Single().NavigationTarget);
    }

    [Fact]
    public void CompletenessIsCarriedIntoThePresentation()
    {
        var complete = NameInventoryPresentation.Create(Build(Workbook("Rate", "=Sheet1!$A$1")), "Book.xlsx");
        Assert.Equal(AuditTraceStatus.Complete, complete.Status);

        var withGap = NameInventoryPresentation.Create(Build(Workbook("Odd", "=!!!")), "Book.xlsx");
        Assert.Equal(AuditTraceStatus.Partial, withGap.Status);
        Assert.Contains("Partial", withGap.CompletenessStatement);
    }

    [Fact]
    public void AFilteredViewSaysHowManyOfTheWholeItIsShowing()
    {
        var inventory = Build(Workbook("GrowthRate", "=Sheet1!$A$1"), Workbook("Discount", "=42"));
        var filtered = NameInventorySearch.Filter(inventory, "growth");

        var presentation = NameInventoryPresentation.Create(inventory, "Book.xlsx", filtered);

        Assert.Single(presentation.Rows);
        Assert.Contains("1 of 2", presentation.Headline);
    }

    [Fact]
    public void DuplicateAndHiddenFlagsReachTheNoteColumn()
    {
        var inventory = NameInventoryBuilder.Build(
            new[]
            {
                new NameRecord("Rate", NameScopeKind.Workbook, null, "=Sheet1!$A$1", true, false),
                new NameRecord("Rate", NameScopeKind.Worksheet, "Sheet2", "=Sheet2!$A$1", false, false),
            },
            includeHidden: true,
            includeBuiltIn: true);

        var presentation = NameInventoryPresentation.Create(inventory, "Book.xlsx");

        Assert.All(presentation.Rows, row => Assert.Contains("Duplicate name", row.Values.Last()));
        Assert.Contains(presentation.Rows, row => row.Values.Last().Contains("Hidden"));
    }

    [Fact]
    public void SummaryCountsEveryTargetKindPresent()
    {
        var presentation = NameInventoryPresentation.Create(
            Build(Workbook("A", "=Sheet1!$A$1"), Workbook("B", "=Sheet1!$A$2"), Workbook("C", "=42")),
            "Book.xlsx");

        Assert.Contains(presentation.SummaryLines, line => line.StartsWith("Range: 2"));
        Assert.Contains(presentation.SummaryLines, line => line.StartsWith("Constant: 1"));
    }
}
