using System;
using System.Drawing;
using System.Windows.Forms;
using ExcelAccel.Application.Auditing;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Compare;
using ExcelAccel.Core.Compare;
using ExcelAccel.ExcelAddIn.Reliability;
using ExcelAccel.ExcelInterop;
using ExcelAccel.Persistence.Compare;

namespace ExcelAccel.ExcelAddIn;

/// <summary>
/// Owns the captured comparison source, the comparison view, and the last
/// result, through the same shared <see cref="TraceViewRuntime"/> every other
/// read-only result uses.
/// </summary>
internal static class ComparisonViewRuntime
{
    private static readonly object Sync = new object();

    private static readonly TraceViewRuntime Runtime = new TraceViewRuntime(
        CompareCommandCatalog.RangesId,
        "Read-only comparison of two open ranges. Neither side is opened, saved, recalculated, or changed.",
        CommandDispatcher.NavigateToTraceTarget);

    private static ComparisonAnchor? _source;
    private static ComparisonSession? _session;

    public static bool IsOpen => Runtime.IsOpen;

    public static ComparisonAnchor? Source
    {
        get { lock (Sync) return _source; }
    }

    public static ComparisonSession? Session
    {
        get { lock (Sync) return _session; }
    }

    public static void CaptureSource(ComparisonAnchor anchor)
    {
        lock (Sync) _source = anchor;
    }

    public static CommandResult Present(ComparisonSession session, IWorkbookPresencePort presence)
    {
        if (session is null) throw new ArgumentNullException(nameof(session));
        lock (Sync) _session = session;
        return Runtime.Present(session.Presentation, session.Result.Source.WorkbookId, presence);
    }

    public static bool RevalidateSource() => Runtime.RevalidateSource();

    public static void Reset()
    {
        Runtime.Reset();
        lock (Sync)
        {
            _source = null;
            _session = null;
        }
    }
}

/// <summary>Confirms and writes a comparison export. Nothing leaves the machine.</summary>
internal static class ComparisonExportRuntime
{
    public static CommandResult Export(ComparisonSession session)
    {
        if (session is null) throw new ArgumentNullException(nameof(session));

        using (var save = new SaveFileDialog
        {
            Title = "Export comparison results",
            Filter = "CSV file (*.csv)|*.csv",
            DefaultExt = "csv",
        })
        {
            if (save.ShowDialog() != DialogResult.OK)
            {
                return CommandResult.Refused(
                    CompareCommandCatalog.ExportId, "No destination was chosen.", RefusalCodes.PreviewRequired);
            }

            var exporter = new ComparisonExporter();
            using (var confirm = new ComparisonExportManifestDialog(exporter.Plan(save.FileName, session.Result, false)))
            {
                var owner = ExcelWindowOwner.TryCreate();
                if ((owner is null ? confirm.ShowDialog() : confirm.ShowDialog(owner)) != DialogResult.OK)
                {
                    return CommandResult.Refused(
                        CompareCommandCatalog.ExportId, "The export manifest was not confirmed.", RefusalCodes.PreviewRequired);
                }

                var manifest = exporter.Plan(save.FileName, session.Result, confirm.IncludeContent);
                exporter.Export(manifest, session.Result);
                return CommandResult.Success(
                    CompareCommandCatalog.ExportId,
                    $"Exported {manifest.DifferenceCount:N0} difference(s) to {manifest.Destination}.");
            }
        }
    }
}

internal sealed class ComparisonExportManifestDialog : Form
{
    private readonly CheckBox _includeContent = new CheckBox();

    public ComparisonExportManifestDialog(ComparisonExportManifest manifest)
    {
        if (manifest is null) throw new ArgumentNullException(nameof(manifest));
        Text = "Confirm comparison export";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MinimizeBox = false;
        MaximizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new Size(600, 300);
        AccessibleName = "Confirm comparison export";
        AccessibleDescription = "Review exactly what will be written before the file is created.";

        var summary = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            Dock = DockStyle.Top,
            Height = 185,
            Text = manifest.ToString(),
            AccessibleName = "Export manifest",
            AccessibleDescription = "The sources, categories, counts, fields, and destination.",
        };

        _includeContent.Text = "Include the differing cell contents (formulas and values)";
        _includeContent.Dock = DockStyle.Top;
        _includeContent.Height = 32;
        _includeContent.Checked = manifest.IncludeContent;
        _includeContent.AccessibleName = "Include cell contents";
        _includeContent.AccessibleDescription =
            "Off by default. Without it the export carries positions and categories but no workbook content.";

        var buttons = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            FlowDirection = FlowDirection.RightToLeft,
            Height = 48,
            Padding = new Padding(8),
        };
        var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AccessibleName = "Cancel export" };
        var export = new Button { Text = "Export", DialogResult = DialogResult.OK, AccessibleName = "Confirm export" };
        buttons.Controls.Add(cancel);
        buttons.Controls.Add(export);

        Controls.Add(_includeContent);
        Controls.Add(summary);
        Controls.Add(buttons);
        AcceptButton = export;
        CancelButton = cancel;
    }

    public bool IncludeContent => _includeContent.Checked;
}
