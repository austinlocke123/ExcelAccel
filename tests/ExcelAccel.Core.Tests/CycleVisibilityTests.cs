using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Formatting;
using ExcelAccel.Application.Profiles;
using ExcelAccel.Persistence.Profiles;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class CycleVisibilityTests
{
    private static ProfileDefinition Defaults() => new ProfileStore().LoadDefault();

    [Theory]
    [InlineData("format.number.currency")]
    [InlineData("format.number.date")]
    [InlineData("format.font_color.cycle")]
    [InlineData("format.underline.cycle")]
    [InlineData("format.column_width.cycle")]
    public void AConfiguredCycleKeepsItsButton(string commandId) =>
        Assert.True(CycleVisibility.IsVisible(Defaults(), commandId));

    /// <summary>
    /// The gap WP-F-02 left open: a deleted cycle's button used to remain on the
    /// ribbon, refusing by name on every press.
    /// </summary>
    [Fact]
    public void DeletingANumberFormatCycleHidesItsButton()
    {
        var profile = Defaults();
        var trimmed = profile.WithCycles(ProfileCycleEditor.Remove(profile.Cycles, "number_format", "date"));

        Assert.False(CycleVisibility.IsVisible(trimmed, "format.number.date"));
        Assert.True(CycleVisibility.IsVisible(trimmed, "format.number.currency"));
    }

    [Fact]
    public void EmptyingAPropertyFamilyHidesItsButton()
    {
        var profile = Defaults();
        var trimmed = profile.WithCycles(ProfileCycleEditor.Remove(profile.Cycles, "underline", "standard"));

        Assert.False(CycleVisibility.IsVisible(trimmed, "format.underline.cycle"));
        Assert.True(CycleVisibility.IsVisible(trimmed, "format.font_color.cycle"));
    }

    /// <summary>
    /// The decimals commands operate on any number format, including formats
    /// belonging to no cycle, so deleting every cycle must not hide them.
    /// </summary>
    [Theory]
    [InlineData("format.number.decimals.increase")]
    [InlineData("format.number.decimals.decrease")]
    public void TheDecimalsCommandsAreNeverHidden(string commandId)
    {
        var profile = Defaults();
        var cycles = profile.Cycles;
        foreach (var cycle in cycles["number_format"].Select(value => value.CycleId).ToArray())
        {
            cycles = ProfileCycleEditor.Remove(cycles, "number_format", cycle);
        }

        Assert.True(CycleVisibility.IsVisible(profile.WithCycles(cycles), commandId));
    }

    [Theory]
    [InlineData("inspect.selection.summary")]
    [InlineData("format.border.sum_bar.apply")]
    [InlineData("view.zoom.set")]
    [InlineData("")]
    public void ACommandBoundToNoCycleIsAlwaysVisible(string commandId) =>
        Assert.True(CycleVisibility.IsVisible(Defaults(), commandId));

    /// <summary>
    /// Generated cycle commands come from the profile, so they cannot name a
    /// cycle that is not configured.
    /// </summary>
    [Fact]
    public void AGeneratedCycleCommandIsAlwaysVisible() =>
        Assert.True(CycleVisibility.IsVisible(Defaults(), "format.cycle.number_format.basis_points"));

    /// <summary>
    /// Every ribbon control that carries getVisible must be one this rule can
    /// actually decide, or the attribute is decoration.
    /// </summary>
    [Fact]
    public void EveryHideableButtonRespondsToItsFamilyBeingEmptied()
    {
        var profile = Defaults();
        var families = new Dictionary<string, string>
        {
            ["format.font_color.cycle"] = "font_color",
            ["format.fill_color.cycle"] = "fill_color",
            ["format.font_size.cycle"] = "font_size",
            ["format.row_height.cycle"] = "row_height",
            ["format.column_width.cycle"] = "column_width",
            ["format.alignment.horizontal.cycle"] = "horizontal_alignment",
            ["format.alignment.vertical.cycle"] = "vertical_alignment",
            ["format.underline.cycle"] = "underline",
        };

        foreach (var pair in families)
        {
            var cycles = profile.Cycles;
            foreach (var cycle in cycles[pair.Value].Select(value => value.CycleId).ToArray())
            {
                cycles = ProfileCycleEditor.Remove(cycles, pair.Value, cycle);
            }

            Assert.False(
                CycleVisibility.IsVisible(profile.WithCycles(cycles), pair.Key),
                $"{pair.Key} stayed visible after '{pair.Value}' was emptied.");
        }
    }
}
