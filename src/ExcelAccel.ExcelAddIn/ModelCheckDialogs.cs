using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ExcelAccel.Core.Auditing;
using ExcelAccel.Core.ModelCheck;
using ExcelAccel.Persistence.ModelCheck;

namespace ExcelAccel.ExcelAddIn;

/// <summary>
/// Confirms exactly which findings will be suppressed, and states what is
/// stored: the rule identity and a normalized fingerprint, never content.
/// </summary>
internal sealed class ModelCheckIgnoreDialog : Form
{
    private readonly CheckedListBox _findings = new CheckedListBox();

    public ModelCheckIgnoreDialog(IReadOnlyList<ModelCheckFinding> findings, IReadOnlyList<ModelCheckIgnoreEntry> existing)
    {
        Text = "Ignore Model Check Findings Locally";
        StartPosition = FormStartPosition.CenterParent;
        Size = new Size(760, 520);
        MinimumSize = new Size(600, 380);
        AccessibleName = "Ignore Model Check findings locally";

        var explanation = new Label
        {
            Dock = DockStyle.Top,
            Height = 66,
            Padding = new Padding(6),
            AccessibleName = "What an ignore stores",
            Text =
                "An ignore stores the rule ID, rule version, and a normalized fingerprint only." + Environment.NewLine +
                "No formula or value content is written. It suppresses only findings whose fingerprint matches exactly." + Environment.NewLine +
                "A rescan is required for an ignore to take effect.",
        };

        _findings.Dock = DockStyle.Fill;
        _findings.CheckOnClick = true;
        _findings.AccessibleName = "Findings to ignore";
        var alreadyIgnored = new HashSet<string>(existing.Select(entry => entry.Fingerprint), StringComparer.Ordinal);
        foreach (var finding in findings.Where(value => !alreadyIgnored.Contains(value.Fingerprint)))
        {
            _findings.Items.Add(new Row(finding));
        }

        var ok = new Button { Text = "Ignore Selected", AutoSize = true, DialogResult = DialogResult.OK };
        var cancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 42, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(4) };
        buttons.Controls.Add(cancel);
        buttons.Controls.Add(ok);

        Controls.Add(_findings);
        Controls.Add(explanation);
        Controls.Add(buttons);
        AcceptButton = ok;
        CancelButton = cancel;
    }

    public IReadOnlyList<ModelCheckFinding> SelectedFindings =>
        _findings.CheckedItems.Cast<Row>().Select(row => row.Finding).ToArray();

    private sealed class Row
    {
        public Row(ModelCheckFinding finding) => Finding = finding;

        public ModelCheckFinding Finding { get; }

        public override string ToString() =>
            Finding.RuleId + "  " + AuditPresentationLabels.Location(Finding.Target) +
            "  [" + Finding.Fingerprint.Substring(0, 12) + "]";
    }
}

/// <summary>Shows the active local ignores and removes selected entries.</summary>
internal sealed class ModelCheckUnignoreDialog : Form
{
    private readonly CheckedListBox _ignores = new CheckedListBox();

    public ModelCheckUnignoreDialog(IReadOnlyList<ModelCheckIgnoreEntry> ignores)
    {
        Text = "Manage Local Model Check Ignores";
        StartPosition = FormStartPosition.CenterParent;
        Size = new Size(760, 480);
        MinimumSize = new Size(600, 360);
        AccessibleName = "Manage local Model Check ignores";

        var explanation = new Label
        {
            Dock = DockStyle.Top,
            Height = 44,
            Padding = new Padding(6),
            AccessibleName = "How removal works",
            Text =
                "These fingerprints are suppressed locally. Check the entries to remove." + Environment.NewLine +
                "A rescan is required for a removal to take effect.",
        };

        _ignores.Dock = DockStyle.Fill;
        _ignores.CheckOnClick = true;
        _ignores.AccessibleName = "Active local ignores";
        foreach (var entry in ignores) _ignores.Items.Add(new Row(entry));

        var ok = new Button { Text = "Remove Selected", AutoSize = true, DialogResult = DialogResult.OK };
        var cancel = new Button { Text = "Close", AutoSize = true, DialogResult = DialogResult.Cancel };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 42, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(4) };
        buttons.Controls.Add(cancel);
        buttons.Controls.Add(ok);

        Controls.Add(_ignores);
        Controls.Add(explanation);
        Controls.Add(buttons);
        AcceptButton = ok;
        CancelButton = cancel;
    }

    public IReadOnlyCollection<string> RemovedFingerprints =>
        new HashSet<string>(_ignores.CheckedItems.Cast<Row>().Select(row => row.Entry.Fingerprint), StringComparer.Ordinal);

    private sealed class Row
    {
        public Row(ModelCheckIgnoreEntry entry) => Entry = entry;

        public ModelCheckIgnoreEntry Entry { get; }

        public override string ToString() =>
            Entry.RuleId + " v" + Entry.RuleVersion + "  " + Entry.ScopeNote + "  [" + Entry.Fingerprint.Substring(0, 12) + "]";
    }
}

/// <summary>
/// The export manifest. It states every field the file will contain before
/// anything is written, and evidence is opt-in.
/// </summary>
internal sealed class ModelCheckExportManifestDialog : Form
{
    private readonly CheckBox _includeEvidence = new CheckBox();

    public ModelCheckExportManifestDialog(ExportManifest manifest)
    {
        Text = "Confirm Model Check Export";
        StartPosition = FormStartPosition.CenterParent;
        Size = new Size(700, 420);
        MinimumSize = new Size(560, 320);
        AccessibleName = "Confirm the Model Check export manifest";

        var lines = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            AccessibleName = "Export manifest",
            Text = string.Join(Environment.NewLine, manifest.Lines),
        };

        _includeEvidence.Dock = DockStyle.Bottom;
        _includeEvidence.Height = 30;
        _includeEvidence.Checked = false;
        _includeEvidence.Text = "Include rule evidence (may quote formula structure)";
        _includeEvidence.AccessibleName = "Include rule evidence in the export";

        var ok = new Button { Text = "Export", AutoSize = true, DialogResult = DialogResult.OK };
        var cancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 42, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(4) };
        buttons.Controls.Add(cancel);
        buttons.Controls.Add(ok);

        Controls.Add(lines);
        Controls.Add(_includeEvidence);
        Controls.Add(buttons);
        AcceptButton = ok;
        CancelButton = cancel;
    }

    public bool IncludeEvidence => _includeEvidence.Checked;
}
