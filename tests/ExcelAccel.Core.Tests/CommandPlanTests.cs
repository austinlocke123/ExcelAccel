using System;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Commands;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class CommandPlanTests
{
    private static readonly SelectionContext Context = new SelectionContext("Book.xlsx", "Sheet1", "A1:B2");

    [Fact]
    public void ReadOnlyPlanCannotDeclareChangedProperties()
    {
        Assert.Throws<ArgumentException>(() => new CommandPlan(
            "read.test",
            CommandImpact.ReadOnly,
            Context,
            new[] { "value" },
            4,
            "invalid"));
    }

    [Fact]
    public void MutationMustDeclareChangedProperties()
    {
        Assert.Throws<ArgumentException>(() => new CommandPlan(
            "write.test",
            CommandImpact.Low,
            Context,
            Array.Empty<string>(),
            4,
            "invalid"));
    }

    [Fact]
    public void SelectionContextUsesCaseInsensitiveWorkbookIdentity()
    {
        var left = new SelectionContext("C:\\MODEL.XLSX", "Sheet1", "A1");
        var right = new SelectionContext("c:\\model.xlsx", "Sheet1", "A1");

        Assert.Equal(left, right);
        Assert.Equal(left.GetHashCode(), right.GetHashCode());
    }

    [Fact]
    public void EveryRegisteredCommandHasUniqueIdAndKeyboardRoute()
    {
        var commands = BuiltInCommandRegistry.All;

        Assert.Equal(commands.Count, commands.Select(command => command.Id).Distinct(StringComparer.Ordinal).Count());
        Assert.All(commands, command => Assert.False(string.IsNullOrWhiteSpace(command.KeyboardRoute)));
    }
}
