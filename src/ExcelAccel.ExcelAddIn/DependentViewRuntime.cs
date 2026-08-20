using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using ExcelAccel.Application.Auditing;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Auditing;
using ExcelAccel.ExcelAddIn.Reliability;

namespace ExcelAccel.ExcelAddIn;

/// <summary>
/// Owns the single read-only direct-dependent view. It mirrors the precedent
/// view's lifecycle deliberately: a shared view will be extracted from the two
/// once both are proven against real Excel, rather than designed up front while
/// the lifecycle has no unit coverage.
/// </summary>
internal static class DependentViewRuntime
{
    private static DirectDependentView? _view;

    public static bool IsOpen => _view is not null && !_view.IsDisposed;

    public static CommandResult Present(DirectDependentReport report, string workbookId, IWorkbookPresencePort presence)
    {
        if (report is null) throw new ArgumentNullException(nameof(report));
        if (presence is null) throw new ArgumentNullException(nameof(presence));
        if (!IsOpen)
        {
            _view = new DirectDependentView();
            _view.FormClosed += (_, __) => _view = null;
            var owner = ExcelWindowOwner.TryCreate();
            if (owner is null) _view.Show(); else _view.Show(owner);
        }

        _view!.Present(report, workbookId, presence);
        _view.Activate();
        _view.FocusResults();
        return report.Status == AuditTraceStatus.Refused
            ? CommandResult.Refused(AuditingCommandCatalog.DirectDependentsId, report.Headline, report.RefusalCode ?? RefusalCodes.CommandUnavailable)
            : CommandResult.Success(AuditingCommandCatalog.DirectDependentsId, report.Headline);
    }

    public static bool RevalidateSource() => IsOpen && _view!.RevalidateSource();

    public static void Reset()
    {
        if (IsOpen) _view!.Close();
        _view = null;
    }
}

internal sealed class DirectDependentView : Form
{
    private readonly Label _headline = new Label();
    private readonly Label _notice = new Label();
    private readonly ListView _dependents = new ListView();
    private readonly TextBox _summary = new TextBox();
    private string _workbookId = string.Empty;
    private IWorkbookPresencePort? _presence;
    private bool _revalidating;

    public DirectDependentView()
    {
        Text = "ExcelAccel Direct Dependents";
        StartPosition = FormStartPosition.CenterParent;
        Size = new Size(880, 520);
        MinimumSize = new Size(640, 380);
        ShowInTaskbar = false;
        KeyPreview = true;
        AccessibleName = "ExcelAccel Direct Dependents";
        AccessibleDescription = "Read-only direct dependents of one target within one worksheet. This view never changes the workbook.";

        _headline.Dock = DockStyle.Top;
        _headline.Height = 38;
        _headline.Padding = new Padding(4);
        _headline.AccessibleName = "Direct dependent summary";

        _notice.Dock = DockStyle.Top;
        _notice.Height = 22;
        _notice.Padding = new Padding(4, 0, 4, 0);
        _notice.Visible = false;
        _notice.AccessibleName = "Direct dependent notice";

        _dependents.Dock = DockStyle.Fill;
        _dependents.View = View.Details;
        _dependents.FullRowSelect = true;
        _dependents.MultiSelect = false;
        _dependents.HideSelection = false;
        _dependents.AccessibleName = "Direct dependents";
        _dependents.Columns.Add("Dependent", 200);
        _dependents.Columns.Add("Reached by", 120);
        _dependents.Columns.Add("Edges", 60);
        _dependents.Columns.Add("Source reference", 300);

        _summary.Dock = DockStyle.Right;
        _summary.Width = 280;
        _summary.Multiline = true;
        _summary.ReadOnly = true;
        _summary.ScrollBars = ScrollBars.Vertical;
        _summary.AccessibleName = "Scan scope and coverage summary";

        var close = new Button { Text = "Close", AutoSize = true, DialogResult = DialogResult.Cancel, AccessibleName = "Close the direct dependent view" };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 42, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(4) };
        buttons.Controls.Add(close);

        Controls.Add(_dependents);
        Controls.Add(_summary);
        Controls.Add(_notice);
        Controls.Add(_headline);
        Controls.Add(buttons);
        CancelButton = close;
        Activated += (_, __) => RevalidateSource();
    }

    public void Present(DirectDependentReport report, string workbookId, IWorkbookPresencePort presence)
    {
        _workbookId = workbookId ?? string.Empty;
        _presence = presence;
        _headline.Text = report.Headline + Environment.NewLine + report.CompletenessStatement;
        _notice.Visible = false;
        _notice.Text = string.Empty;
        _summary.Text = string.Join(Environment.NewLine, report.SummaryLines);
        _dependents.BeginUpdate();
        try
        {
            _dependents.Items.Clear();
            foreach (var row in report.Rows)
            {
                var item = new ListViewItem(row.DisplayTarget);
                item.SubItems.Add(row.Kinds);
                item.SubItems.Add(row.EdgeCount.ToString(CultureInfo.InvariantCulture));
                item.SubItems.Add(row.SourceEvidence);
                _dependents.Items.Add(item);
            }
        }
        finally
        {
            _dependents.EndUpdate();
        }
    }

    public void FocusResults()
    {
        _dependents.Focus();
        if (_dependents.Items.Count > 0 && _dependents.SelectedItems.Count == 0) _dependents.Items[0].Selected = true;
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
            DiagnosticLog.Error("audit.dependents.direct.revalidate", exception);
            presence = WorkbookPresence.Unknown;
        }
        finally
        {
            _revalidating = false;
        }

        if (presence == WorkbookPresence.Closed)
        {
            DiagnosticLog.Info(AuditingCommandCatalog.DirectDependentsId, "view_closed:source_workbook_closed");
            Close();
            return false;
        }

        if (presence == WorkbookPresence.Unknown)
        {
            _notice.Text = "The source workbook state could not be verified; this result is a point-in-time scan.";
            _notice.Visible = true;
        }

        return true;
    }
}
