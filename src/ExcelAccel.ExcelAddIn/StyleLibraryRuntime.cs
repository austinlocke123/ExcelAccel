using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Styles;
using ExcelAccel.ExcelAddIn.Reliability;

namespace ExcelAccel.ExcelAddIn;

internal static class StyleLibraryRuntime
{
    private static StyleLibraryDialog? _dialog;
    public static CommandResult Open()
    {
        if (_dialog is not null && !_dialog.IsDisposed) { _dialog.Activate(); _dialog.FocusList(); return CommandResult.Success("style.library.open", "Style Library is focused."); }
        _dialog = new StyleLibraryDialog();
        _dialog.FormClosed += (_, __) => _dialog = null;
        var owner = ExcelWindowOwner.TryCreate();
        if (owner is null) _dialog.Show(); else _dialog.Show(owner);
        _dialog.FocusList();
        return CommandResult.Success("style.library.open", "Style Library opened.");
    }
    public static void Reset() { if (_dialog is not null && !_dialog.IsDisposed) _dialog.Close(); _dialog = null; }
}

internal sealed class StyleLibraryDialog : Form
{
    private readonly ListView _styles = new ListView();
    private readonly TextBox _properties = new TextBox();
    private readonly Button _apply = new Button();
    private readonly Button _delete = new Button();

    public StyleLibraryDialog()
    {
        Text = "ExcelAccel Style Library"; StartPosition = FormStartPosition.CenterParent; Size = new Size(720, 500);
        MinimumSize = new Size(600, 400); ShowInTaskbar = false; KeyPreview = true;
        AccessibleName = "ExcelAccel Style Library";
        _styles.Dock = DockStyle.Fill; _styles.View = View.Details; _styles.FullRowSelect = true; _styles.MultiSelect = false; _styles.HideSelection = false;
        _styles.AccessibleName = "Available styles"; _styles.Columns.Add("Style", 260); _styles.Columns.Add("Origin", 120); _styles.Columns.Add("Properties", 100);
        _styles.SelectedIndexChanged += (_, __) => UpdateSelection(); _styles.DoubleClick += (_, __) => ApplySelected(); _styles.KeyDown += OnListKeyDown;
        _properties.Dock = DockStyle.Right; _properties.Width = 250; _properties.Multiline = true; _properties.ReadOnly = true; _properties.ScrollBars = ScrollBars.Vertical;
        _properties.AccessibleName = "Selected style properties";
        _apply.Text = "Apply"; _apply.AutoSize = true; _apply.Click += (_, __) => ApplySelected(); _apply.AccessibleName = "Apply selected style";
        var capture = new Button { Text = "Capture New...", AutoSize = true, AccessibleName = "Capture a new local style" };
        capture.Click += (_, __) => CaptureNew();
        _delete.Text = "Delete Local"; _delete.AutoSize = true; _delete.Click += (_, __) => DeleteSelected(); _delete.AccessibleName = "Delete selected local style";
        var close = new Button { Text = "Close", AutoSize = true, DialogResult = DialogResult.Cancel };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 42, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(4) };
        buttons.Controls.Add(close); buttons.Controls.Add(_apply); buttons.Controls.Add(_delete); buttons.Controls.Add(capture);
        Controls.Add(_styles); Controls.Add(_properties); Controls.Add(buttons); CancelButton = close;
        RefreshStyles();
    }

    public void FocusList() { _styles.Focus(); if (_styles.Items.Count > 0 && _styles.SelectedItems.Count == 0) _styles.Items[0].Selected = true; }

    private void RefreshStyles()
    {
        var selected = Selected?.StyleId; _styles.BeginUpdate();
        try
        {
            _styles.Items.Clear();
            foreach (var style in StyleLibrary.Effective(ProfileRuntime.Current.LocalStyles))
            {
                var item = new ListViewItem(style.DisplayName) { Tag = style };
                item.SubItems.Add(style.Origin == StyleOrigin.Local ? "Local" : style.Origin == StyleOrigin.BuiltIn ? "Built-in" : "Organization");
                item.SubItems.Add(style.Properties.Count.ToString()); _styles.Items.Add(item);
            }
            if (_styles.Items.Count > 0)
            {
                var match = _styles.Items.Cast<ListViewItem>().FirstOrDefault(value => ((StyleRecipe)value.Tag).StyleId == selected);
                (match ?? _styles.Items[0]).Selected = true;
            }
        }
        finally { _styles.EndUpdate(); }
        UpdateSelection();
    }

    private void UpdateSelection()
    {
        var style = Selected; _apply.Enabled = style is not null; _delete.Enabled = style?.IsDeletable == true;
        _properties.Text = style is null ? string.Empty : string.Join(Environment.NewLine, style.Properties.Select(value => value.Key + " = " + value.Value));
    }

    private void ApplySelected()
    {
        var style = Selected; if (style is null) return;
        CallbackBoundary.Run(style.Origin == StyleOrigin.BuiltIn ? "style.apply_builtin" : "style.apply",
            () => CommandDispatcher.ApplyStyle(style.StyleId));
        RefreshStyles();
    }

    private void CaptureNew()
    {
        using (var dialog = new StyleCaptureDialog())
        {
            var owner = ExcelWindowOwner.TryCreate();
            if ((owner is null ? dialog.ShowDialog() : dialog.ShowDialog(owner)) != DialogResult.OK) return;
            CallbackBoundary.Run("style.capture", () => CommandDispatcher.CaptureLocalStyle(dialog.StyleName, dialog.SelectedProperties));
        }
        RefreshStyles();
    }

    private void DeleteSelected()
    {
        var style = Selected; if (style?.IsDeletable != true) return;
        CallbackBoundary.Run("style.delete_local", () => CommandDispatcher.DeleteLocalStyle(style.StyleId)); RefreshStyles();
    }

    private void OnListKeyDown(object? sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter) { ApplySelected(); e.Handled = true; }
        else if (e.Control && e.KeyCode == Keys.N) { CaptureNew(); e.Handled = true; }
        else if (e.KeyCode == Keys.Delete) { DeleteSelected(); e.Handled = true; }
    }
    private StyleRecipe? Selected => _styles.SelectedItems.Count == 0 ? null : (StyleRecipe)_styles.SelectedItems[0].Tag;
}

internal sealed class StyleCaptureDialog : Form
{
    private readonly TextBox _name = new TextBox();
    private readonly CheckedListBox _properties = new CheckedListBox();
    public StyleCaptureDialog()
    {
        Text = "Capture Local Style"; StartPosition = FormStartPosition.CenterParent; Size = new Size(440, 480); MinimizeBox = false; MaximizeBox = false;
        var nameLabel = new Label { Text = "Style name", Dock = DockStyle.Top, Height = 22 };
        _name.Dock = DockStyle.Top; _name.AccessibleName = "Style name";
        _properties.Dock = DockStyle.Fill; _properties.CheckOnClick = true; _properties.AccessibleName = "Formatting properties to capture";
        foreach (var property in StylePropertyCatalog.Supported) _properties.Items.Add(property, property != "font_name" && property != "font_italic");
        var ok = new Button { Text = "Capture", AutoSize = true, DialogResult = DialogResult.OK };
        ok.Click += (_, e) => { if (string.IsNullOrWhiteSpace(_name.Text) || _properties.CheckedItems.Count == 0) { MessageBox.Show("Enter a name and select at least one property.", "ExcelAccel", MessageBoxButtons.OK, MessageBoxIcon.Warning); DialogResult = DialogResult.None; } };
        var cancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 42, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(4) };
        buttons.Controls.Add(cancel); buttons.Controls.Add(ok);
        Controls.Add(_properties); Controls.Add(_name); Controls.Add(nameLabel); Controls.Add(buttons); AcceptButton = ok; CancelButton = cancel;
    }
    public string StyleName => _name.Text.Trim();
    public IReadOnlyList<string> SelectedProperties => _properties.CheckedItems.Cast<string>().ToArray();
}
