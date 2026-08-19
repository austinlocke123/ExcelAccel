using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Commands;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class CommandFrameworkTests
{
    private static readonly SelectionContext Context = new SelectionContext("Book.xlsx", "Sheet1", "A1:B2");

    [Fact]
    public void CanonicalPlanIsStableAcrossInputEnumerationOrder()
    {
        var first = Plan(
            new[] { "font_color", "number_format" },
            new[]
            {
                new KeyValuePair<string, string>("style", "currency"),
                new KeyValuePair<string, string>("locale", "invariant"),
            });
        var second = Plan(
            new[] { "number_format", "font_color" },
            new[]
            {
                new KeyValuePair<string, string>("locale", "invariant"),
                new KeyValuePair<string, string>("style", "currency"),
            });

        Assert.Equal(first.CanonicalJson, second.CanonicalJson);
        Assert.Equal(first.PlanHash, second.PlanHash);
        Assert.DoesNotContain("summary", first.CanonicalJson, StringComparison.Ordinal);
    }

    [Fact]
    public void MandatoryPreviewRequiresExactPlanHash()
    {
        var descriptor = Descriptor(PreviewPolicy.Mandatory);
        var plan = Plan(new[] { "font_color" }, Array.Empty<KeyValuePair<string, string>>());

        var missing = CommandExecutionGate.Authorize(descriptor, plan);
        var wrong = CommandExecutionGate.Authorize(descriptor, plan, new string('0', 64));
        var exact = CommandExecutionGate.Authorize(descriptor, plan, plan.PlanHash);

        Assert.False(missing.Allowed);
        Assert.Equal(RefusalCodes.PreviewRequired, missing.RefusalCode);
        Assert.False(wrong.Allowed);
        Assert.True(exact.Allowed);
    }

    [Fact]
    public void ContractDriftIsRefusedBeforeExecution()
    {
        var descriptor = Descriptor(PreviewPolicy.None);
        var plan = new CommandPlan(
            "test.command",
            CommandImpact.Low,
            Context,
            new[] { "different_property" },
            4,
            "test");

        var result = CommandExecutionGate.Authorize(descriptor, plan);

        Assert.False(result.Allowed);
        Assert.Equal(RefusalCodes.ContractMismatch, result.RefusalCode);
    }

    [Fact]
    public void RegistryEntriesContainCompleteReleaseMetadata()
    {
        Assert.All(BuiltInCommandRegistry.All, descriptor =>
        {
            Assert.False(string.IsNullOrWhiteSpace(descriptor.Id));
            Assert.True(descriptor.ContractVersion > 0);
            Assert.False(string.IsNullOrWhiteSpace(descriptor.CapabilityId));
            Assert.NotEqual(CommandContextRequirement.None, descriptor.ContextRequirement);
            Assert.NotEmpty(descriptor.AcceptanceIds);
            Assert.False(string.IsNullOrWhiteSpace(descriptor.KeyboardRoute));
            Assert.Equal(
                descriptor.ChangedProperties.OrderBy(value => value, StringComparer.Ordinal),
                descriptor.ChangedProperties);
        });
    }

    private static CommandPlan Plan(
        IEnumerable<string> changedProperties,
        IEnumerable<KeyValuePair<string, string>> arguments) =>
        new CommandPlan(
            "test.command",
            CommandImpact.Low,
            Context,
            changedProperties,
            4,
            "Localized summaries are deliberately excluded from the canonical form.",
            "ABC123",
            arguments: arguments);

    private static CommandDescriptor Descriptor(PreviewPolicy previewPolicy) =>
        new CommandDescriptor(
            "test.command",
            1,
            "Test Command",
            CommandImpact.Low,
            new[] { "font_color" },
            true,
            "Test only",
            "CAP-TEST",
            CommandContextRequirement.Selection,
            previewPolicy,
            UndoPolicy.SessionPropertyReceipt,
            new[] { "AC-TEST-001" });
}
