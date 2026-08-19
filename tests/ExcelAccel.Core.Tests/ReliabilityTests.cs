using ExcelAccel.Core.Reliability;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class ReliabilityTests
{
    [Fact]
    public void ReentrancyGateRejectsOverlapAndReopensAfterDispose()
    {
        var gate = new ReentrancyGate();

        var first = gate.TryEnter();
        Assert.NotNull(first);
        Assert.Null(gate.TryEnter());

        first!.Dispose();
        using var second = gate.TryEnter();
        Assert.NotNull(second);
    }

    [Fact]
    public void CoreAssemblyHasNoExcelOrComReferences()
    {
        var references = typeof(ReentrancyGate).Assembly.GetReferencedAssemblies();

        Assert.DoesNotContain(references, reference => reference.Name?.StartsWith("ExcelDna", System.StringComparison.Ordinal) == true);
        Assert.DoesNotContain(references, reference => reference.Name?.StartsWith("Microsoft.Office.Interop", System.StringComparison.Ordinal) == true);
    }
}
