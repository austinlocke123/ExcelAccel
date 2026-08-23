using System;
using System.Drawing;
using System.Windows.Forms;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Names;
using ExcelAccel.ExcelAddIn.Reliability;
using ExcelAccel.Persistence.Names;

namespace ExcelAccel.ExcelAddIn;

/// <summary>
/// Confirms and writes a name-inventory export. Nothing leaves the machine, and
/// the manifest is shown before any file is written.
/// </summary>
internal static class NameInventoryExportRuntime
{
    public static CommandResult Export(NameInventoryResult captured)
    {
        if (captured is null) throw new ArgumentNullException(nameof(captured));

        using (var save = new SaveFileDialog
        {
            Title = "Export named-range inventory",
            Filter = "CSV file (*.csv)|*.csv",
            DefaultExt = "csv",
        })
        {
            if (save.ShowDialog() != DialogResult.OK)
            {
                return CommandResult.Refused(
                    NamesCommandCatalog.ExportId, "No destination was chosen.", RefusalCodes.PreviewRequired);
            }

            var exporter = new NameInventoryExporter();
            using (var confirm = new NameExportManifestDialog(exporter.Plan(save.FileName, captured.Inventory, false)))
            {
                var owner = ExcelWindowOwner.TryCreate();
                if ((owner is null ? confirm.ShowDialog() : confirm.ShowDialog(owner)) != DialogResult.OK)
                {
                    return CommandResult.Refused(
                        NamesCommandCatalog.ExportId, "The export manifest was not confirmed.", RefusalCodes.PreviewRequired);
                }

                var manifest = exporter.Plan(save.FileName, captured.Inventory, confirm.IncludeExpressions);
                exporter.Export(manifest, captured.Inventory);
                return CommandResult.Success(
                    NamesCommandCatalog.ExportId,
                    $"Exported {manifest.NameCount:N0} name(s) to {manifest.Destination}.");
            }
        }
    }
}

internal sealed class NameExportManifestDialog : Form
{
    private readonly CheckBox _includeExpressions = new CheckBox();

    public NameExportManifestDialog(NameExportManifest manifest)
    {
        if (manifest is null) throw new ArgumentNullException(nameof(manifest));
        Text = "Confirm named-range export";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MinimizeBox = false;
        MaximizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new Size(560, 260);
        AccessibleName = "Confirm named-range export";
        AccessibleDescription = "Review exactly what will be written before the file is created.";

        var summary = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            Dock = DockStyle.Top,
            Height = 150,
            Text = manifest.ToString(),
            AccessibleName = "Export manifest",
            AccessibleDescription = "The destination, the number of names, and the fields that will be written.",
        };

        _includeExpressions.Text = "Include each name's expression (may contain file paths)";
        _includeExpressions.Dock = DockStyle.Top;
        _includeExpressions.Height = 32;
        _includeExpressions.Checked = manifest.IncludeExpressions;
        _includeExpressions.AccessibleName = "Include name expressions";
        _includeExpressions.AccessibleDescription =
            "Off by default. A name's expression can carry a file path or a business term.";

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

        Controls.Add(_includeExpressions);
        Controls.Add(summary);
        Controls.Add(buttons);
        AcceptButton = export;
        CancelButton = cancel;
    }

    public bool IncludeExpressions => _includeExpressions.Checked;
}
