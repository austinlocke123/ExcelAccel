using System;
using System.Drawing;
using System.Windows.Forms;
using ExcelAccel.Application.Auditing;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Auditing;
using ExcelAccel.ExcelAddIn.Reliability;

namespace ExcelAccel.ExcelAddIn;

/// <summary>
/// Owns one read-only trace view. Every auditing result — precedents,
/// dependents, and later traversals — is presented through this one runtime, so
/// the lifecycle exists in exactly one place.
///
/// All lifecycle decisions live in <see cref="TraceViewSession"/>, which the test
/// project can reach. This class only creates, renders, and closes a window.
/// </summary>
internal sealed class TraceViewRuntime
{
    private readonly TraceViewSession _session;
    private readonly string _accessibleDescription;
    private TraceResultView? _view;

    public TraceViewRuntime(string commandId, string accessibleDescription)
    {
        _session = new TraceViewSession(commandId);
        _accessibleDescription = accessibleDescription;
    }

    public bool IsOpen => _view is not null && !_view.IsDisposed;

    public CommandResult Present(TraceResultPresentation presentation, string workbookId, IWorkbookPresencePort presence)
    {
        if (presentation is null) throw new ArgumentNullException(nameof(presentation));
        if (presence is null) throw new ArgumentNullException(nameof(presence));
        _session.Present(presentation, workbookId, presence);
        if (!IsOpen)
        {
            _view = new TraceResultView(_accessibleDescription, RevalidateSource);
            _view.FormClosed += (_, __) =>
            {
                _view = null;
                _session.Clear();
            };
            var owner = ExcelWindowOwner.TryCreate();
            if (owner is null) _view.Show(); else _view.Show(owner);
        }

        _view!.Render(presentation, _session.Notice);
        _view.Activate();
        _view.FocusResults();
        return presentation.Status == AuditTraceStatus.Refused
            ? CommandResult.Refused(_session.CommandId, presentation.Headline, presentation.RefusalCode ?? RefusalCodes.CommandUnavailable)
            : CommandResult.Success(_session.CommandId, presentation.Headline);
    }

    /// <summary>Returns true while a view is still presenting a result.</summary>
    public bool RevalidateSource()
    {
        if (!IsOpen) return false;
        var action = _session.Revalidate();
        if (_session.LastProbeError is not null)
        {
            DiagnosticLog.Error(_session.CommandId + ".revalidate", _session.LastProbeError);
        }

        switch (action)
        {
            case TraceViewAction.Discard:
                DiagnosticLog.Info(_session.CommandId, "view_closed:source_workbook_closed");
                _view!.Close();
                return false;
            case TraceViewAction.Warn:
                _view!.ShowNotice(_session.Notice);
                return true;
            default:
                _view!.ShowNotice(null);
                return true;
        }
    }

    public void Reset()
    {
        if (IsOpen) _view!.Close();
        _view = null;
        _session.Clear();
    }
}

internal sealed class TraceResultView : Form
{
    private readonly Label _headline = new Label();
    private readonly Label _notice = new Label();
    private readonly ListView _results = new ListView();
    private readonly TextBox _summary = new TextBox();

    public TraceResultView(string accessibleDescription, Func<bool> revalidate)
    {
        StartPosition = FormStartPosition.CenterParent;
        Size = new Size(880, 520);
        MinimumSize = new Size(640, 380);
        ShowInTaskbar = false;
        KeyPreview = true;
        AccessibleDescription = accessibleDescription;

        _headline.Dock = DockStyle.Top;
        _headline.Height = 38;
        _headline.Padding = new Padding(4);
        _headline.AccessibleName = "Trace result summary";

        _notice.Dock = DockStyle.Top;
        _notice.Height = 22;
        _notice.Padding = new Padding(4, 0, 4, 0);
        _notice.Visible = false;
        _notice.AccessibleName = "Trace result notice";

        _results.Dock = DockStyle.Fill;
        _results.View = View.Details;
        _results.FullRowSelect = true;
        _results.MultiSelect = false;
        _results.HideSelection = false;
        _results.AccessibleName = "Trace results";

        _summary.Dock = DockStyle.Right;
        _summary.Width = 280;
        _summary.Multiline = true;
        _summary.ReadOnly = true;
        _summary.ScrollBars = ScrollBars.Vertical;
        _summary.AccessibleName = "Scope, coverage, and completeness summary";

        var close = new Button { Text = "Close", AutoSize = true, DialogResult = DialogResult.Cancel, AccessibleName = "Close the trace result view" };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 42, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(4) };
        buttons.Controls.Add(close);

        Controls.Add(_results);
        Controls.Add(_summary);
        Controls.Add(_notice);
        Controls.Add(_headline);
        Controls.Add(buttons);
        CancelButton = close;
        Activated += (_, __) => revalidate();
    }

    public void Render(TraceResultPresentation presentation, string? notice)
    {
        Text = presentation.Title;
        AccessibleName = presentation.Title;
        _headline.Text = presentation.Headline + Environment.NewLine + presentation.CompletenessStatement;
        _summary.Text = string.Join(Environment.NewLine, presentation.SummaryLines);
        ShowNotice(notice);
        _results.BeginUpdate();
        try
        {
            _results.Items.Clear();
            _results.Columns.Clear();
            foreach (var column in presentation.Columns) _results.Columns.Add(column.Header, column.Width);
            foreach (var row in presentation.Rows)
            {
                var item = new ListViewItem(row[0]);
                for (var index = 1; index < row.Count; index++) item.SubItems.Add(row[index]);
                _results.Items.Add(item);
            }
        }
        finally
        {
            _results.EndUpdate();
        }
    }

    public void ShowNotice(string? notice)
    {
        _notice.Text = notice ?? string.Empty;
        _notice.Visible = !string.IsNullOrEmpty(notice);
    }

    public void FocusResults()
    {
        _results.Focus();
        if (_results.Items.Count > 0 && _results.SelectedItems.Count == 0) _results.Items[0].Selected = true;
    }
}
