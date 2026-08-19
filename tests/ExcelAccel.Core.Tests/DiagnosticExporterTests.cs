using System;
using System.IO;
using System.Text;
using ExcelAccel.Persistence.Diagnostics;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class DiagnosticExporterTests
{
    [Fact]
    public void ExportRequiresExactManifestAndPreservesExistingDestinationOnRefusal()
    {
        var directory = Path.Combine(Path.GetTempPath(), "excelaccel-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var path = Path.Combine(directory, "support.txt");
            File.WriteAllText(path, "existing");
            var bytes = Encoding.UTF8.GetBytes("sanitized diagnostic\n");
            var exporter = new DiagnosticExporter();
            var plan = exporter.Plan(path, bytes);
            Assert.Contains("no workbook formulas", plan.Manifest, StringComparison.Ordinal);

            Assert.Throws<InvalidOperationException>(() => exporter.Export(plan, "wrong", bytes));
            Assert.Equal("existing", File.ReadAllText(path));
            exporter.Export(plan, plan.PlanHash, bytes);
            Assert.Equal(bytes, File.ReadAllBytes(path));
            Assert.Empty(Directory.GetFiles(directory, "*.tmp"));
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void ChangedContentAfterPreviewIsRefused()
    {
        var exporter = new DiagnosticExporter();
        var plan = exporter.Plan(Path.Combine(Path.GetTempPath(), "support.txt"), Encoding.UTF8.GetBytes("one"));
        Assert.Throws<InvalidOperationException>(() => exporter.Export(plan, plan.PlanHash, Encoding.UTF8.GetBytes("two")));
    }
}
