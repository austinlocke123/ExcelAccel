using System.Linq;
using ExcelAccel.Application.Formatting;
using ExcelAccel.Persistence.Profiles;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class NumberFormatDiagnosticsTests
{
    [Theory]
    [InlineData("#,##0_);(#,##0)")]
    [InlineData("0.0%_);(0.0%)")]
    [InlineData("m/d/yyyy")]
    [InlineData("\"TRUE\";\"TRUE\";\"FALSE\"")]
    [InlineData("0\" bps\"")]
    [InlineData("[Red]0.00")]
    public void AUsableFormatIsAccepted(string candidate) =>
        Assert.Equal(NumberFormatVerdict.Accepted, NumberFormatDiagnostics.Inspect(candidate).Verdict);

    [Theory]
    [InlineData("", "empty")]
    [InlineData("   ", "empty")]
    [InlineData("=0.00", "formula")]
    [InlineData("0.00\"unterminated", "unterminated")]
    [InlineData("[Red0.00", "unclosed")]
    [InlineData("0.00]", "matching")]
    [InlineData("0;0;0;0;0", "sections")]
    public void AnUnusableFormatIsRejectedWithAReason(string candidate, string expectedFragment)
    {
        var diagnostic = NumberFormatDiagnostics.Inspect(candidate);

        Assert.Equal(NumberFormatVerdict.Rejected, diagnostic.Verdict);
        Assert.Contains(expectedFragment, diagnostic.Message, System.StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Measured against Excel: writing these and reading back yields
    /// <c>[$£-809]</c> and <c>[$€-2]</c>. They are exactly what Excel's own
    /// currency dialog produces, which is what makes this a trap rather than an
    /// obvious mistake.
    /// </summary>
    [Theory]
    [InlineData("[$£-en-GB]#,##0_);([$£-en-GB]#,##0)")]
    [InlineData("[$€-x-euro2]#,##0_);([$€-x-euro2]#,##0)")]
    public void ALocaleQualifiedCurrencyTokenIsFlaggedAsRewritten(string candidate)
    {
        var diagnostic = NumberFormatDiagnostics.Inspect(candidate);

        Assert.Equal(NumberFormatVerdict.RewriteExpected, diagnostic.Verdict);
        Assert.Contains("stays on its first entry", diagnostic.Message, System.StringComparison.Ordinal);
        Assert.NotNull(diagnostic.Suggestion);
    }

    [Fact]
    public void TheSuggestedReplacementIsTheBareSymbolFormThatRoundTrips()
    {
        var diagnostic = NumberFormatDiagnostics.Inspect("[$£-en-GB]#,##0_);([$£-en-GB]#,##0)");

        Assert.StartsWith("£#,##0_)", diagnostic.Suggestion!, System.StringComparison.Ordinal);
        Assert.Equal(NumberFormatVerdict.Accepted, NumberFormatDiagnostics.Inspect(diagnostic.Suggestion).Verdict);
    }

    [Fact]
    public void ACurrencyTokenWithNoLocaleIsLeftAlone() =>
        Assert.Equal(NumberFormatVerdict.Accepted, NumberFormatDiagnostics.Inspect("[$€]#,##0").Verdict);

    [Fact]
    public void AFormatExcelStoresVerbatimRoundTrips() =>
        Assert.Equal(
            NumberFormatVerdict.Accepted,
            NumberFormatDiagnostics.EvaluateRoundTrip("£#,##0_);(£#,##0)", "£#,##0_);(£#,##0)").Verdict);

    [Fact]
    public void AFormatExcelRewritesIsReportedWithWhatItStored()
    {
        var diagnostic = NumberFormatDiagnostics.EvaluateRoundTrip(
            "[$£-en-GB]#,##0_);([$£-en-GB]#,##0)",
            "[$£-809]#,##0_);([$£-809]#,##0)");

        Assert.Equal(NumberFormatVerdict.RewriteExpected, diagnostic.Verdict);
        Assert.Contains("[$£-809]", diagnostic.Message, System.StringComparison.Ordinal);
        Assert.Equal("[$£-809]#,##0_);([$£-809]#,##0)", diagnostic.Suggestion);
    }

    /// <summary>
    /// Every shipped default must pass its own check, or the product ships an
    /// entry it would refuse if the user typed it.
    /// </summary>
    [Fact]
    public void EveryDefaultNumberFormatEntryPassesInspection()
    {
        var cycles = new ProfileStore().LoadDefault().Cycles["number_format"];

        var failures = cycles
            .SelectMany(cycle => cycle.Entries.Select(entry => new { cycle.CycleId, entry }))
            .Where(item => NumberFormatDiagnostics.Inspect(item.entry).Verdict != NumberFormatVerdict.Accepted)
            .Select(item => $"{item.CycleId}: {item.entry} -> {NumberFormatDiagnostics.Inspect(item.entry).Message}")
            .ToArray();

        Assert.True(failures.Length == 0, string.Join("\n", failures));
    }
}
