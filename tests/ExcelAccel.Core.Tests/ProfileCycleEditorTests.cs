using System;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Formatting;
using ExcelAccel.Application.Profiles;
using ExcelAccel.Persistence.Profiles;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class ProfileCycleEditorTests
{
    private static ProfileCycles Defaults() => new ProfileStore().LoadDefault().Cycles;

    [Fact]
    public void AUserCanAddACycleThatDidNotShipWithTheProduct()
    {
        var added = ProfileCycleEditor.Add(Defaults(), new ProfileCycle(
            "number_format", "spread", "Spread", new[] { "0\" bps\"", "0.0\" bps\"" }));

        Assert.True(added.TryGet("number_format", "spread", out var cycle));
        Assert.Equal(new[] { "0\" bps\"", "0.0\" bps\"" }, cycle.Entries);
        Assert.Equal(7, added["number_format"].Count);
    }

    [Fact]
    public void AddingANinthCycleToAFamilyIsRefusedNamingTheLimit()
    {
        var cycles = Defaults();
        for (var index = 0; index < 2; index++)
        {
            cycles = ProfileCycleEditor.Add(cycles, new ProfileCycle(
                "number_format", "extra" + index, "Extra " + index, new[] { "0.0" + index }));
        }

        var error = Assert.Throws<ArgumentException>(() => ProfileCycleEditor.Add(cycles, new ProfileCycle(
            "number_format", "overflow", "Overflow", new[] { "0.000" })));

        Assert.Contains("8", error.Message, StringComparison.Ordinal);
        Assert.Equal(8, cycles["number_format"].Count);
    }

    [Fact]
    public void ACycleNameCannotBeUsedTwiceInOneFamily() =>
        Assert.Throws<ArgumentException>(() => ProfileCycleEditor.Add(Defaults(), new ProfileCycle(
            "number_format", "currency", "Currency Again", new[] { "0.00" })));

    /// <summary>
    /// Removing the last cycle in a family removes the family, so nothing is left
    /// for a command to find and refuse on. An empty family is the phantom slot
    /// AC-FMT-039 forbids.
    /// </summary>
    [Fact]
    public void RemovingTheLastCycleInAFamilyRemovesTheFamily()
    {
        var trimmed = ProfileCycleEditor.Remove(Defaults(), "underline", "standard");

        Assert.Empty(trimmed["underline"]);
        Assert.DoesNotContain("underline", trimmed.Families);
    }

    [Fact]
    public void RemovingACycleThatIsNotThereIsRefused() =>
        Assert.Throws<ArgumentException>(() => ProfileCycleEditor.Remove(Defaults(), "number_format", "nonexistent"));

    /// <summary>
    /// Slot order is user data, and slot zero is load-bearing: commands whose
    /// ribbon label names no particular cycle follow whichever sits first.
    /// </summary>
    [Fact]
    public void MovingACycleChangesWhichOneAnUnnamedCommandFollows()
    {
        var profile = new ProfileStore().LoadDefault();
        var extended = ProfileCycleEditor.Add(profile.Cycles, new ProfileCycle(
            "font_color", "highlight", "Highlight", new[] { "#FF00FF", "#00FFFF" }));

        var before = profile.WithCycles(extended).ResolveFirstCycle("font_color");
        var moved = ProfileCycleEditor.Move(extended, "font_color", "highlight", -1);
        var after = profile.WithCycles(moved).ResolveFirstCycle("font_color");

        Assert.Equal("#000000", before[0]);
        Assert.Equal("#FF00FF", after[0]);
    }

    [Fact]
    public void MovingBeyondTheFamilyIsRefused() =>
        Assert.Throws<ArgumentOutOfRangeException>(() => ProfileCycleEditor.Move(Defaults(), "number_format", "general", -1));

    [Fact]
    public void AnInvalidEntryIsRefusedAndLeavesTheOriginalUntouched()
    {
        var original = Defaults();

        Assert.Throws<ArgumentException>(() =>
            ProfileCycleEditor.SetEntries(original, "font_color", "standard", new[] { "not-a-colour" }));

        Assert.Equal(new[] { "@same_sheet", "@text", "@hardcode", "@cross_sheet", "@error", "@external" },
            original["font_color"][0].Entries);
    }

    [Fact]
    public void RenameChangesOnlyTheDisplayName()
    {
        var renamed = ProfileCycleEditor.Rename(Defaults(), "number_format", "multiple", "Turns");

        Assert.True(renamed.TryGet("number_format", "multiple", out var cycle));
        Assert.Equal("Turns", cycle.DisplayName);
        Assert.Equal(Defaults()["number_format"].Single(value => value.CycleId == "multiple").Entries, cycle.Entries);
    }
}

public sealed class CycleCommandFactoryTests
{
    [Fact]
    public void CyclesAlreadyCoveredByARibbonCommandGetNoDuplicateEntry()
    {
        var descriptors = CycleCommandFactory.Descriptors(new ProfileStore().LoadDefault());

        Assert.Empty(descriptors);
    }

    [Fact]
    public void AUserAddedCycleBecomesSearchableUnderItsOwnName()
    {
        var profile = new ProfileStore().LoadDefault();
        var extended = profile.WithCycles(ProfileCycleEditor.Add(profile.Cycles, new ProfileCycle(
            "number_format", "spread", "Spread", new[] { "0\" bps\"" })));

        var descriptor = Assert.Single(CycleCommandFactory.Descriptors(extended));

        Assert.Equal("format.cycle.number_format.spread", descriptor.Id);
        Assert.Equal("Spread", descriptor.DisplayName);
        Assert.Equal(new[] { "number_format" }, descriptor.ChangedProperties);
        Assert.Equal(UndoPolicy.SessionPropertyReceipt, descriptor.UndoPolicy);
    }

    [Fact]
    public void AGeneratedIdResolvesBackToARunnableCycle()
    {
        var profile = new ProfileStore().LoadDefault();
        var extended = profile.WithCycles(ProfileCycleEditor.Add(profile.Cycles, new ProfileCycle(
            "number_format", "spread", "Spread", new[] { "0\" bps\"", "0.0\" bps\"" })));

        var command = CycleCommandFactory.Create(extended, "format.cycle.number_format.spread");
        var port = new StubPort("0\" bps\"");
        var plan = command.Plan(extended, port);
        var result = command.Execute(plan, port);

        Assert.True(result.Succeeded);
        Assert.Equal("0.0\" bps\"", port.Value);
    }

    /// <summary>
    /// A stored route or cheat-sheet entry can outlive the cycle it named. That
    /// is the one case where saying nothing would leave the user guessing.
    /// </summary>
    [Fact]
    public void ADeletedCycleRefusesByNameAndWritesNothing()
    {
        var profile = new ProfileStore().LoadDefault();

        var refusal = Assert.Throws<CommandRefusedException>(
            () => CycleCommandFactory.Create(profile, "format.cycle.number_format.spread"));

        Assert.Equal(RefusalCodes.CommandUnavailable, refusal.RefusalCode);
        Assert.Contains("spread", refusal.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AnUnknownFamilyIsRefusedRatherThanGuessed() =>
        Assert.Throws<CommandRefusedException>(
            () => CycleCommandFactory.Create(new ProfileStore().LoadDefault(), "format.cycle.not_a_family.x"));

    [Fact]
    public void GeneratedCommandsAdvertiseAnHonestRouteRatherThanAnAltSequence()
    {
        var profile = new ProfileStore().LoadDefault();
        var extended = profile.WithCycles(ProfileCycleEditor.Add(profile.Cycles, new ProfileCycle(
            "number_format", "spread", "Spread", new[] { "0\" bps\"" })));

        var descriptor = Assert.Single(CycleCommandFactory.Descriptors(extended));

        Assert.DoesNotContain("Alt,", descriptor.KeyboardRoute, StringComparison.Ordinal);
        Assert.Equal(descriptor.KeyboardRoute, descriptor.ShortcutLabel);
    }

    private sealed class StubPort : IFormattingPort
    {
        public StubPort(string value) => Value = value;

        public string Value { get; private set; }

        public Core.Commands.SelectionSnapshot CaptureSelection() =>
            new Core.Commands.SelectionSnapshot(
                new Core.Commands.SelectionContext("Book.xlsx", "Sheet1", "A1"), 1, false, "General");

        public void SetNumberFormat(string formatCode) => throw new NotSupportedException();

        public string ReadFormattingProperty(string propertyId) => Value;

        public void WriteFormattingProperty(string propertyId, string invariantValue) => Value = invariantValue;
    }
}
