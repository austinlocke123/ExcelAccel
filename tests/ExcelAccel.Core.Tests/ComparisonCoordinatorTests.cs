using System;
using System.IO;
using System.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Compare;
using ExcelAccel.Core.Commands;
using ExcelAccel.Core.Compare;
using ExcelAccel.Persistence.Compare;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class ComparisonCoordinatorTests
{
    private static ComparisonBlock Block(string workbook, string sheet, string address, params string[] values) =>
        new ComparisonBlock(workbook, sheet, address, values.Length, 1,
            values.Select((value, index) => new ComparisonCell("A" + (index + 1), string.Empty, value, "General")));

    [Fact]
    public void ComparingWithoutACapturedSourceRefusesWithGuidance()
    {
        var refusal = Assert.Throws<CommandRefusedException>(
            () => new ComparisonCoordinator().Compare(new FakePort { Source = null }));

        Assert.Equal(RefusalCodes.CommandUnavailable, refusal.RefusalCode);
        Assert.Contains("Capture a comparison source first", refusal.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ComparingARangeWithItselfIsRefused()
    {
        var block = Block("Book.xlsx", "Sheet1", "A1:A2", "1", "2");

        var refusal = Assert.Throws<CommandRefusedException>(
            () => new ComparisonCoordinator().Compare(new FakePort { Source = block, Target = block }));

        Assert.Equal(RefusalCodes.SelectionUnsupported, refusal.RefusalCode);
    }

    /// <summary>
    /// A shape mismatch becomes a refusal with a reason, never an alignment.
    /// </summary>
    [Fact]
    public void AShapeMismatchBecomesARefusalRatherThanAnException()
    {
        var refusal = Assert.Throws<CommandRefusedException>(() => new ComparisonCoordinator().Compare(new FakePort
        {
            Source = Block("Left.xlsx", "Sheet1", "A1:A2", "1", "2"),
            Target = Block("Right.xlsx", "Sheet1", "A1", "1"),
        }));

        Assert.Equal(RefusalCodes.SelectionUnsupported, refusal.RefusalCode);
        Assert.Contains("does not align or truncate", refusal.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AComparisonProjectsIntoTheSharedView()
    {
        var session = new ComparisonCoordinator().Compare(new FakePort
        {
            Source = Block("Left.xlsx", "Sheet1", "A1:A2", "1", "2"),
            Target = Block("Right.xlsx", "Sheet1", "A1:A2", "1", "9"),
        });

        Assert.Single(session.Result.Differences);
        Assert.Single(session.Presentation.Rows);
        Assert.All(session.Presentation.Rows, row => Assert.Equal(session.Presentation.Columns.Count, row.Values.Count));
    }

    /// <summary>
    /// Rows navigate to the source side by default; the target side is a separate
    /// projection of the same result, and re-reads nothing.
    /// </summary>
    [Fact]
    public void TheTargetSideIsAReprojectionRatherThanASecondCapture()
    {
        var port = new FakePort
        {
            Source = Block("Left.xlsx", "SourceSheet", "A1:A2", "1", "2"),
            Target = Block("Right.xlsx", "TargetSheet", "A1:A2", "1", "9"),
        };
        var coordinator = new ComparisonCoordinator();
        var session = coordinator.Compare(port);

        var target = coordinator.ShowTargetSide(session);

        Assert.Equal("SourceSheet", session.Presentation.Rows.Single().NavigationTarget!.WorksheetName);
        Assert.Equal("TargetSheet", target.Presentation.Rows.Single().NavigationTarget!.WorksheetName);
        Assert.Equal(1, port.TargetCaptureCount);
    }

    [Fact]
    public void EveryCompareCommandIsRegisteredAsReadOnly()
    {
        foreach (var id in new[]
        {
            CompareCommandCatalog.CaptureSourceId,
            CompareCommandCatalog.RangesId,
            CompareCommandCatalog.NavigateTargetId,
            CompareCommandCatalog.ExportId,
        })
        {
            var descriptor = BuiltInCommandRegistry.GetRequired(id);
            Assert.Equal(CommandImpact.ReadOnly, descriptor.Impact);
            Assert.Empty(descriptor.ChangedProperties);
            Assert.Equal(UndoPolicy.None, descriptor.UndoPolicy);
            Assert.Equal("CAP-CMP-001", descriptor.CapabilityId);
        }
    }

    private sealed class FakePort : IComparisonPort
    {
        public ComparisonBlock? Source { get; set; }
        public ComparisonBlock? Target { get; set; }
        public int TargetCaptureCount { get; private set; }

        public ComparisonBlock? CaptureSource() => Source;

        public ComparisonBlock CaptureTarget()
        {
            TargetCaptureCount++;
            return Target ?? throw new InvalidOperationException("No target was configured.");
        }
    }
}

public sealed class ComparisonExporterTests
{
    private static ComparisonResult Sample()
    {
        var source = new ComparisonBlock("Left.xlsx", "Sheet1", "A1:A2", 2, 1, new[]
        {
            new ComparisonCell("A1", "=SECRETRATE*2", string.Empty, "General"),
            new ComparisonCell("A2", string.Empty, "1234", "General"),
        });
        var target = new ComparisonBlock("Right.xlsx", "Sheet1", "A1:A2", 2, 1, new[]
        {
            new ComparisonCell("A1", "=SECRETRATE*3", string.Empty, "General"),
            new ComparisonCell("A2", string.Empty, "5678", "General"),
        });
        return SameShapeComparer.Compare(source, target);
    }

    /// <summary>
    /// The differing cells are the ones most likely to carry the numbers and
    /// formulas a model is about, so they leave the machine only on request.
    /// </summary>
    [Fact]
    public void CellContentIsExcludedByDefault()
    {
        using var sandbox = new TemporaryDirectory();
        var result = Sample();
        var exporter = new ComparisonExporter();
        var redacted = Path.Combine(sandbox.Path, "redacted.csv");
        var full = Path.Combine(sandbox.Path, "full.csv");

        exporter.Export(exporter.Plan(redacted, result, false), result);
        exporter.Export(exporter.Plan(full, result, true), result);

        var redactedText = File.ReadAllText(redacted);
        Assert.DoesNotContain("SECRETRATE", redactedText, StringComparison.Ordinal);
        Assert.DoesNotContain("5678", redactedText, StringComparison.Ordinal);
        Assert.Contains("SECRETRATE", File.ReadAllText(full), StringComparison.Ordinal);
    }

    [Fact]
    public void TheManifestStatesSourcesCountsFieldsAndCoverage()
    {
        var manifest = new ComparisonExporter()
            .Plan(Path.Combine(Path.GetTempPath(), "compare.csv"), Sample(), false);

        Assert.Equal("Left.xlsx!Sheet1!A1:A2", manifest.SourceIdentity);
        Assert.Equal("Right.xlsx!Sheet1!A1:A2", manifest.TargetIdentity);
        Assert.Equal(2, manifest.DifferenceCount);
        Assert.Equal(2, manifest.ComparedCells);
        Assert.Contains("excluded", manifest.ToString(), StringComparison.Ordinal);
    }

    /// <summary>
    /// A failed or repeated export must never leave a half-written file behind.
    /// </summary>
    [Fact]
    public void RewritingAnExportReplacesItWholeAndLeavesNoTemporaryFile()
    {
        using var sandbox = new TemporaryDirectory();
        var exporter = new ComparisonExporter();
        var path = Path.Combine(sandbox.Path, "compare.csv");
        var result = Sample();

        exporter.Export(exporter.Plan(path, result, true), result);
        exporter.Export(exporter.Plan(path, result, false), result);

        Assert.DoesNotContain("SECRETRATE", File.ReadAllText(path), StringComparison.Ordinal);
        Assert.Empty(Directory.GetFiles(sandbox.Path, "*.tmp"));
    }

    [Fact]
    public void TheExportNamesBothSidesAndItsCoverage()
    {
        using var sandbox = new TemporaryDirectory();
        var exporter = new ComparisonExporter();
        var path = Path.Combine(sandbox.Path, "compare.csv");
        var result = Sample();

        exporter.Export(exporter.Plan(path, result, false), result);

        var text = File.ReadAllText(path);
        Assert.Contains("Left.xlsx!Sheet1!A1:A2 vs Right.xlsx!Sheet1!A1:A2", text, StringComparison.Ordinal);
        Assert.Contains("Partial", text, StringComparison.Ordinal);
    }

    [Fact]
    public void AnEmptyDestinationIsRefused() =>
        Assert.Throws<ArgumentException>(() => new ComparisonExporter().Plan(" ", Sample(), false));

    private sealed class TemporaryDirectory : IDisposable
    {
        public TemporaryDirectory()
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(), "excelaccel-compare-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            try { Directory.Delete(Path, recursive: true); } catch (IOException) { }
        }
    }
}
