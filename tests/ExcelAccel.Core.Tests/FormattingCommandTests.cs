using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Formatting;
using ExcelAccel.Application.Profiles;
using ExcelAccel.Core.Commands;
using ExcelAccel.Persistence.Profiles;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class FormattingCommandTests
{
    [Fact]
    public void EveryFormattingContractHasUniqueKeyboardRouteAndCanPlan()
    {
        var commands = Phase1AFormattingCatalog.All.ToArray();
        Assert.Equal(commands.Length, commands.Select(value => value.Id).Distinct(StringComparer.Ordinal).Count());
        Assert.Equal(commands.Length, commands.Select(value => value.KeyboardRoute).Distinct(StringComparer.OrdinalIgnoreCase).Count());
        var profile = new ProfileStore().LoadDefault();

        foreach (var descriptor in commands.Where(value => value.Id != "format.number.decimals.decrease"))
        {
            var port = new FakeFormattingPort(PropertyFor(descriptor));
            var plan = Phase1AFormattingCatalog.Create(descriptor.Id).Plan(profile, port);
            Assert.Equal(descriptor.ChangedProperties, plan.ChangedProperties);
            Assert.Equal(descriptor.ContractVersion, plan.ContractVersion);
        }
    }

    [Fact]
    public void CycleFromUnknownUsesFirstProfileValueAndWritesOnlyDeclaredProperty()
    {
        var profile = new ProfileStore().LoadDefault();
        var port = new FakeFormattingPort("#123456");
        var command = Phase1AFormattingCatalog.Create("format.font_color.cycle");
        var plan = command.Plan(profile, port);

        var result = command.Execute(plan, port);

        Assert.True(result.Succeeded);
        Assert.Equal("font_color", port.LastProperty);
        Assert.Equal(profile.FontColorCycle[0], port.Value);
        Assert.Equal(1, port.WriteCount);
    }

    [Fact]
    public void InterveningPropertyChangeRefusesWithoutWrite()
    {
        var port = new FakeFormattingPort("general");
        var command = Phase1AFormattingCatalog.Create("format.alignment.horizontal.cycle");
        var plan = command.Plan(new ProfileStore().LoadDefault(), port);
        port.Value = "right";

        var result = command.Execute(plan, port);

        Assert.Equal(CommandResultStatus.Refused, result.Status);
        Assert.Equal(RefusalCodes.StaleContext, result.RefusalCode);
        Assert.Equal(0, port.WriteCount);
    }

    [Fact]
    public void PostconditionMismatchCannotReportSuccess()
    {
        var port = new FakeFormattingPort("#123456") { IgnoreWrites = true };
        var command = Phase1AFormattingCatalog.Create("format.font_color.cycle");
        var plan = command.Plan(new ProfileStore().LoadDefault(), port);

        var result = command.Execute(plan, port);

        Assert.Equal(CommandResultStatus.Failed, result.Status);
        Assert.Equal("POSTCONDITION_MISMATCH", result.DiagnosticId);
    }

    [Theory]
    [InlineData("0.00", 1, "0.000")]
    [InlineData("0.00%", -1, "0.0%")]
    [InlineData("$#,##0.00;($#,##0.00);-", 1, "$#,##0.000;($#,##0.000);-")]
    [InlineData("0.00E+00", 1, "")]
    public void DecimalTransformIsDeterministicAndFailClosed(string format, int delta, string expected) =>
        Assert.Equal(expected, NumberFormatDecimals.Change(format, delta));

    [Fact]
    public void FreezeRequiresConfirmationOfExactPlanHash()
    {
        var command = Phase1AFormattingCatalog.Create("view.panes.freeze");
        var port = new FakeFormattingPort("false");
        var plan = command.Plan(new ProfileStore().LoadDefault(), port);

        Assert.Equal(CommandResultStatus.Refused, command.Execute(plan, port).Status);
        Assert.True(command.Execute(plan, port, plan.PlanHash).Succeeded);
    }

    private static string PropertyFor(CommandDescriptor descriptor)
    {
        switch (descriptor.ChangedProperties[0])
        {
            case "number_format": return "0.00";
            case "font_color": return "#000000";
            case "fill_color": return "#FFFFFF";
            case "font_size": return "10";
            case "horizontal_alignment": return "general";
            case "vertical_alignment": return "bottom";
            case "indent_level": return "1";
            case "underline": return "none";
            case "row_height": return "15";
            case "column_width": return "10";
            case "borders": return "none";
            case "gridlines": return "true";
            case "zoom": return "90";
            case "freeze_panes": return "false";
            default: return string.Empty;
        }
    }

    private sealed class FakeFormattingPort : IFormattingPort
    {
        public FakeFormattingPort(string value) => Value = value;
        public string Value { get; set; }
        public string? LastProperty { get; private set; }
        public int WriteCount { get; private set; }
        public bool IgnoreWrites { get; set; }
        public SelectionSnapshot CaptureSelection() => new SelectionSnapshot(new SelectionContext("Book.xlsx", "Sheet1", "A1:B2"), 4, false, "General");
        public void SetNumberFormat(string formatCode) => throw new NotSupportedException();
        public string ReadFormattingProperty(string propertyId) => Value;
        public void WriteFormattingProperty(string propertyId, string invariantValue)
        {
            WriteCount++;
            LastProperty = propertyId;
            if (!IgnoreWrites) Value = invariantValue == "autofit" ? "15" : invariantValue;
        }
    }
}
