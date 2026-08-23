using System;
using System.Drawing;
using System.Windows.Forms;
using ExcelAccel.Application.Auditing;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Links;
using ExcelAccel.ExcelAddIn.Reliability;
using ExcelAccel.Persistence.Links;

namespace ExcelAccel.ExcelAddIn;

/// <summary>
/// Owns the external-link inventory view and the snapshot behind it, through the
/// same shared <see cref="TraceViewRuntime"/> every other read-only result uses.
/// </summary>
internal static class LinkInventoryViewRuntime
{
    private static readonly object Sync = new object();

    private static readonly TraceViewRuntime Runtime = new TraceViewRuntime(
        LinksCommandCatalog.InventoryId,
        "Read-only inventory of external links. This view never opens, updates, repoints, or breaks a link.",
        CommandDispatcher.NavigateToTraceTarget);

    private static LinkInventoryResult? _captured;

    public static bool IsOpen => Runtime.IsOpen;

    public static LinkInventoryResult? Captured
    {
        get { lock (Sync) return _captured; }
    }

    public static CommandResult Present(LinkInventoryResult result, IWorkbookPresencePort presence)
    {
        if (result is null) throw new ArgumentNullException(nameof(result));
        lock (Sync) _captured = result;
        return Runtime.Present(result.Presentation, result.WorkbookId, presence);
    }

    public static bool RevalidateSource() => Runtime.RevalidateSource();

    public static void Reset()
    {
        Runtime.Reset();
        lock (Sync) _captured = null;
    }
}

/// <summary>Confirms and writes a link-inventory export. Nothing leaves the machine.</summary>
internal static class LinkInventoryExportRuntime
{
    public static CommandResult Export(LinkInventoryResult captured)
    {
        if (captured is null) throw new ArgumentNullException(nameof(captured));

        using (var save = new SaveFileDialog
        {
            Title = "Export external-link inventory",
            Filter = "CSV file (*.csv)|*.csv",
            DefaultExt = "csv",
        })
        {
            if (save.ShowDialog() != DialogResult.OK)
            {
                return CommandResult.Refused(
                    LinksCommandCatalog.ExportId, "No destination was chosen.", RefusalCodes.PreviewRequired);
            }

            var exporter = new LinkInventoryExporter();
            using (var confirm = new LinkExportManifestDialog(exporter.Plan(save.FileName, captured.Inventory, false)))
            {
                var owner = ExcelWindowOwner.TryCreate();
                if ((owner is null ? confirm.ShowDialog() : confirm.ShowDialog(owner)) != DialogResult.OK)
                {
                    return CommandResult.Refused(
                        LinksCommandCatalog.ExportId, "The export manifest was not confirmed.", RefusalCodes.PreviewRequired);
                }

                var manifest = exporter.Plan(save.FileName, captured.Inventory, confirm.IncludePaths);
                exporter.Export(manifest, captured.Inventory);
                return CommandResult.Success(
                    LinksCommandCatalog.ExportId,
                    $"Exported {manifest.UsageCount:N0} usage(s) across {manifest.SourceCount:N0} source(s) to {manifest.Destination}.");
            }
        }
    }
}

internal sealed class LinkExportManifestDialog : Form
{
    private readonly CheckBox _includePaths = new CheckBox();

    public LinkExportManifestDialog(LinkExportManifest manifest)
    {
        if (manifest is null) throw new ArgumentNullException(nameof(manifest));
        Text = "Confirm external-link export";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MinimizeBox = false;
        MaximizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new Size(580, 280);
        AccessibleName = "Confirm external-link export";
        AccessibleDescription = "Review exactly what will be written before the file is created.";

        var summary = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            Dock = DockStyle.Top,
            Height = 165,
            Text = manifest.ToString(),
            AccessibleName = "Export manifest",
            AccessibleDescription = "The destination, the counts, the fields, and the coverage statement.",
        };

        _includePaths.Text = "Include full source paths (these name drives, clients, or deals)";
        _includePaths.Dock = DockStyle.Top;
        _includePaths.Height = 32;
        _includePaths.Checked = manifest.IncludePaths;
        _includePaths.AccessibleName = "Include source paths";
        _includePaths.AccessibleDescription =
            "Off by default. Without it the export carries only each source's file name.";

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

        Controls.Add(_includePaths);
        Controls.Add(summary);
        Controls.Add(buttons);
        AcceptButton = export;
        CancelButton = cancel;
    }

    public bool IncludePaths => _includePaths.Checked;
}
