using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Core.Auditing;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class AuditAddressTests
{
    [Theory]
    [InlineData(1, 1, "A1")]
    [InlineData(1, 25, "Y1")]
    [InlineData(1, 26, "Z1")]
    [InlineData(1, 27, "AA1")]
    [InlineData(1, 51, "AY1")]
    [InlineData(1, 52, "AZ1")]
    [InlineData(1, 53, "BA1")]
    [InlineData(700, 702, "ZZ700")]
    [InlineData(1, 703, "AAA1")]
    [InlineData(1048576, 16384, "XFD1048576")]
    public void ColumnNamesAreCorrectAtEveryMultipleOfTwentySix(int row, int column, string expected) =>
        Assert.Equal(expected, AuditAddress.Cell(row, column));

    [Theory]
    [InlineData("A1")]
    [InlineData("Z1")]
    [InlineData("AZ52")]
    [InlineData("ZZ700")]
    [InlineData("XFD1048576")]
    public void CellFormattingRoundTripsThroughParsing(string address)
    {
        Assert.True(AuditAddress.TryParse(address, out var rectangle));

        Assert.Equal(address, AuditAddress.Cell(rectangle.FirstRow, rectangle.FirstColumn));
    }

    [Fact]
    public void AbsoluteMarkersAndReversedEndpointsNormalizeToTheSameRectangle()
    {
        Assert.True(AuditAddress.TryParse("$B$5:$A$1", out var reversed));
        Assert.True(AuditAddress.TryParse("A1:B5", out var forward));

        Assert.Equal(forward, reversed);
        Assert.Equal("A1:B5", forward.ToString());
    }

    [Theory]
    [InlineData("A:A")]
    [InlineData("1:1")]
    [InlineData("Model!A1")]
    [InlineData("A1:B2:C3")]
    [InlineData("AAAA1")]
    [InlineData("A0")]
    [InlineData("A1048577")]
    [InlineData("")]
    [InlineData(null)]
    public void UnqualifiedNotationFailsClosed(string? address) =>
        Assert.False(AuditAddress.TryParse(address, out _));

    [Fact]
    public void IntersectionIsSymmetricAndExcludesDisjointRectangles()
    {
        Assert.True(AuditAddress.TryParse("B2:D4", out var first));
        Assert.True(AuditAddress.TryParse("C3:E9", out var overlapping));
        Assert.True(AuditAddress.TryParse("F1:G2", out var disjoint));

        Assert.True(first.Intersects(overlapping));
        Assert.True(overlapping.Intersects(first));
        Assert.False(first.Intersects(disjoint));
        Assert.False(disjoint.Intersects(first));
    }

    [Fact]
    public void DirectPrecedentsResolveColumnZWithoutShiftingTheTarget()
    {
        var source = new AuditCellIdentity("Book.xlsx", "Model", "B1");
        var index = new ReferenceSnapshotIndex(new[]
        {
            new KeyValuePair<AuditCellIdentity, AuditCellClassification>(
                new AuditCellIdentity("Book.xlsx", "Model", "Z1"), AuditCellClassification.Value),
        });

        var result = new DirectPrecedentAnalyzer().Analyze(
            new FormulaReferenceSnapshot(source, "=Z1+AZ2+ZZ700", index));

        Assert.Equal(
            new[] { "Book.xlsx|Model|AZ2", "Book.xlsx|Model|Z1", "Book.xlsx|Model|ZZ700" },
            result.Precedents.Select(precedent => precedent.Target!.ToString()).OrderBy(value => value, System.StringComparer.Ordinal));
        Assert.Equal(
            AuditCellClassification.Value,
            result.Precedents.Single(precedent => precedent.Target!.Address == "Z1").Classification);
    }

    [Fact]
    public void ACapturePlanRequestsColumnZItselfRatherThanAShiftedCell()
    {
        var source = new AuditCellIdentity("Book.xlsx", "Model", "B1");

        var plan = new DirectPrecedentAnalyzer().CreateCapturePlan(source, "=Z1");

        Assert.Equal(new[] { "Book.xlsx|Model|Z1" }, plan.LocalTargets.Select(value => value.ToString()));
    }
}
