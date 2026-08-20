using System;
using System.Drawing;
using System.Windows.Forms;
using ExcelAccel.Application.Auditing;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Auditing;
using ExcelAccel.ExcelAddIn.Reliability;

namespace ExcelAccel.ExcelAddIn;

/// <summary>
/// Owns the single read-only direct-precedent view. The view presents one
/// captured analysis, never writes to the workbook, and is discarded when the
/// user closes it, when the source workbook is no longer open, or when the
/// add-in unloads.
/// </summary>
internal static class PrecedentViewRuntime
{
    private static DirectPrecedentView? _view;

    public static bool IsOpen => _view is not null && !_view.IsDisposed;

    public static CommandResult Present(DirectPrecedentReport report, string workbookId, IWorkbookPresencePort presence)
    {
        if (report is null) throw new ArgumentNullException(nameof(report));
        if (presence is null) throw new ArgumentNullException(nameof(presence));
        if (!IsOpen)
        {
            _view = new DirectPrecedentView();
            _view.FormClosed += (_, __) => _view = null;
            var owner = ExcelWindowOwner.TryCreate();
            if (owner is null) _view.Show(); else _view.Show(owner);
        }

        _view!.Present(report, workbookId, presence);
        _view.Activate();
        _view.FocusResults();
        return report.Status == AuditTraceStatus.Refused
            ? CommandResult.Refused(AuditingCommandCatalog.DirectPrecedentsId, report.Headline, report.RefusalCode ?? RefusalCodes.CommandUnavailable)
            : CommandResult.Success(AuditingCommandCatalog.DirectPrecedentsId, report.Headline);
    }

    /// <summary>
    /// Re-probes the source workbook and discards the view when the workbook is
    /// no longer open. Returns true when a view is still presenting a result.
    /// </summary>
    public static bool RevalidateSource() => IsOpen && _view!.RevalidateSource();

    public static void Reset()
    {
        if (IsOpen) _view!.Close();
        _view = null;
    }
}

internal sealed class DirectPrecedentView : Form
{
    private readonly Label _headline = new Label();
    private readonly Label _notice = new Label();
    private readonly ListView _precedents = new ListView();
    private readonly TextBox _summary = new TextBox();
    private string _workbookId = string.Empty;
    private IWorkbookPresencePort? _presence;
    private bool _revalidating;

    public DirectPrecedentView()
    {
        Text = "ExcelAccel Direct Precedents";
        StartPosition = FormStartPosition.CenterParent;
        Size = new Size(880, 520);
        MinimumSize = new Size(640, 380);
        ShowInTaskbar = false;
        KeyPreview = true;
        AccessibleName = "ExcelAccel Direct Precedents";
        AccessibleDescription = "Read-only direct precedents of one formula cell. This view never changes the workbook.";

        _headline.Dock = DockStyle.Top;
        _headline.Height = 38;
        _headline.Padding = new Padding(4);
        _headline.AccessibleName = "Direct precedent summary";

        _notice.Dock = DockStyle.Top;
        _notice.Height = 22;
        _notice.Padding = new Padding(4, 0, 4, 0);
        _notice.Visible = false;
        _notice.AccessibleName = "Direct precedent notice";

        _precedents.Dock = DockStyle.Fill;
        _precedents.View = View.Details;
        _precedents.FullRowSelect = true;
        _precedents.MultiSelect = false;
        _precedents.HideSelection = false;
        _precedents.AccessibleName = "Direct precedents";
        _precedents.Columns.Add("Target", 190);
        _precedents.Columns.Add("Kind", 90);
        _precedents.Columns.Add("Contents", 100);
        _precedents.Columns.Add("State", 170);
        _precedents.Columns.Add("Edges", 60);
        _precedents.Columns.Add("Source reference", 240);

        _summary.Dock = DockStyle.Right;
        _summary.Width = 280;
        _summary.Multiline = true;
        _summary.ReadOnly = true;
        _summary.ScrollBars = ScrollBars.Vertical;
        _summary.AccessibleName = "Coverage and completeness summary";

        var close = new Button { Text = "Close", AutoSize = true, DialogResult = DialogResult.Cancel, AccessibleName = "Close the direct precedent view" };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 42, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(4) };
        buttons.Controls.Add(close);

        Controls.Add(_precedents);
        Controls.Add(_summary);
        Controls.Add(_notice);
        Controls.Add(_headline);
        Controls.Add(buttons);
        CancelButton = close;
        Activated += (_, __) => RevalidateSource();
    }

    public void Present(DirectPrecedentReport report, string workbookId, IWorkbookPresencePort presence)
    {
        _workbookId = workbookId ?? string.Empty;
        _presence = presence;
        _headline.Text = report.Headline + Environment.NewLine + report.CompletenessStatement;
        _notice.Visible = false;
        _notice.Text = string.Empty;
        _summary.Text = string.Join(Environment.NewLine, report.SummaryLines);
        _precedents.BeginUpdate();
        try
        {
            _precedents.Items.Clear();
            foreach (var row in report.Rows)
            {
                var item = new ListViewItem(row.DisplayTarget) { Tag = row.NodeId };
                item.SubItems.Add(row.Kind);
                item.SubItems.Add(row.Classification);
                item.SubItems.Add(row.State);
                item.SubItems.Add(row.EdgeCount.ToString(System.Globalization.CultureInfo.InvariantCulture));
                item.SubItems.Add(row.SourceEvidence);
                _precedents.Items.Add(item);
            }
        }
        finally
        {
            _precedents.EndUpdate();
        }
    }

    public void FocusResults()
    {
        _precedents.Focus();
        if (_precedents.Items.Count > 0 && _precedents.SelectedItems.Count == 0) _precedents.Items[0].Selected = true;
    }

    public bool RevalidateSource()
    {
        if (_presence is null || string.IsNullOrEmpty(_workbookId) || _revalidating) return true;
        WorkbookPresence presence;
        _revalidating = true;
        try
        {
            presence = _presence.Probe(_workbookId);
        }
        catch (Exception exception)
        {
            DiagnosticLog.Error("audit.precedents.direct.revalidate", exception);
            presence = WorkbookPresence.Unknown;
        }
        finally
        {
            _revalidating = false;
        }

        if (presence == WorkbookPresence.Closed)
        {
            DiagnosticLog.Info(AuditingCommandCatalog.DirectPrecedentsId, "view_closed:source_workbook_closed");
            Close();
            return false;
        }

        if (presence == WorkbookPresence.Unknown)
        {
            _notice.Text = "The source workbook state could not be verified; this result is a point-in-time capture.";
            _notice.Visible = true;
        }

        return true;
    }
}
