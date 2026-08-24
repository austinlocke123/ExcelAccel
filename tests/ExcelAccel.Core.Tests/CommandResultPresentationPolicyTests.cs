using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Compare;
using ExcelAccel.Application.Links;
using ExcelAccel.Application.Names;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class CommandResultPresentationPolicyTests
{
    [Theory]
    [InlineData(CompareCommandCatalog.RangesId)]
    [InlineData(CompareCommandCatalog.NavigateTargetId)]
    [InlineData(CompareCommandCatalog.ExportId)]
    [InlineData(NamesCommandCatalog.NavigateTargetId)]
    [InlineData(NamesCommandCatalog.ExportId)]
    [InlineData(LinksCommandCatalog.NavigateUsageId)]
    [InlineData(LinksCommandCatalog.ExportId)]
    public void ACommandThatCanRefuseBeforeOpeningAViewKeepsFailureReporting(string commandId) =>
        Assert.False(CommandResultPresentationPolicy.PresentsOwnAuditResult(commandId));

    [Theory]
    [InlineData("audit.precedents.direct")]
    [InlineData("audit.dependents.workbook")]
    [InlineData("audit.formula.inspect")]
    [InlineData(NamesCommandCatalog.InventoryId)]
    [InlineData(LinksCommandCatalog.InventoryId)]
    public void ACommandThatAlwaysPresentsAResultSuppressesDuplicateFailureReporting(string commandId) =>
        Assert.True(CommandResultPresentationPolicy.PresentsOwnAuditResult(commandId));
}
