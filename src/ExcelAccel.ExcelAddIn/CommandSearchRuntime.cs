using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Discovery;
using ExcelAccel.Application.Profiles;
using ExcelAccel.ExcelAddIn.Reliability;

namespace ExcelAccel.ExcelAddIn;

internal static class CommandSearchRuntime
{
    private static CommandSearchDialog? _dialog;

    public static CommandResult Open()
    {
        if (_dialog is not null && !_dialog.IsDisposed)
        {
            _dialog.Activate();
            _dialog.FocusSearch();
            return CommandResult.Success("command.search.open", "Command Search is focused.");
        }

        _dialog = new CommandSearchDialog();
        _dialog.FormClosed += (_, __) => _dialog = null;
        var owner = ExcelWindowOwner.TryCreate();
        if (owner is null) _dialog.Show(); else _dialog.Show(owner);
        _dialog.FocusSearch();
        return CommandResult.Success("command.search.open", "Command Search opened.");
    }

    public static void Reset()
    {
        if (_dialog is not null && !_dialog.IsDisposed) _dialog.Close();
        _dialog = null;
    }
}

internal sealed class CommandSearchDialog : Form
{
    private readonly CommandSearchIndex _index = new CommandSearchIndex(BuiltInCommandRegistry.All);
    private readonly TextBox _query = new TextBox();
    private readonly ListView _results = new ListView();
    private readonly Label _status = new Label();
    private readonly Button _run = new Button();
    private readonly Button _favorite = new Button();

    public CommandSearchDialog()
    {
        Text = "ExcelAccel Command Search";
        StartPosition = FormStartPosition.CenterParent;
        MinimumSize = new Size(720, 420);
        Size = new Size(880, 560);
        ShowInTaskbar = false;
        KeyPreview = true;
        AccessibleName = "ExcelAccel Command Search";
        AccessibleDescription = "Search, inspect availability, favorite, and run ExcelAccel commands.";

        _query.Dock = DockStyle.Top;
        _query.Font = new Font(SystemFonts.MessageBoxFont.FontFamily, 12f);
        _query.AccessibleName = "Search commands";
        _query.AccessibleDescription = "Type a command name, alias, category, description, or shortcut.";
        _query.TextChanged += (_, __) => RefreshResults();
        _query.KeyDown += OnQueryKeyDown;

        _results.Dock = DockStyle.Fill;
        _results.View = View.Details;
        _results.FullRowSelect = true;
        _results.HideSelection = false;
        _results.MultiSelect = false;
        _results.AccessibleName = "Command results";
        _results.Columns.Add("Command", 250);
        _results.Columns.Add("Category", 100);
        _results.Columns.Add("Shortcut", 170);
        _results.Columns.Add("Status", 280);
        _results.SelectedIndexChanged += (_, __) => UpdateSelection();
        _results.DoubleClick += (_, __) => RunSelected();
        _results.KeyDown += OnResultsKeyDown;

        _status.Dock = DockStyle.Top;
        _status.Height = 42;
        _status.AutoEllipsis = true;
        _status.AccessibleName = "Selected command status";

        _run.Text = "Run";
        _run.AutoSize = true;
        _run.AccessibleName = "Run selected command";
        _run.Click += (_, __) => RunSelected();
        _favorite.Text = "Add Favorite";
        _favorite.AutoSize = true;
        _favorite.AccessibleName = "Toggle selected command favorite";
        _favorite.Click += (_, __) => ToggleFavorite();
        var close = new Button { Text = "Close", AutoSize = true, DialogResult = DialogResult.Cancel };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 42, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(4) };
        buttons.Controls.Add(close); buttons.Controls.Add(_run); buttons.Controls.Add(_favorite);

        Controls.Add(_results);
        Controls.Add(_status);
        Controls.Add(buttons);
        Controls.Add(_query);
        CancelButton = close;
        RefreshResults();
    }

    public void FocusSearch() { _query.Focus(); _query.SelectAll(); }

    private void RefreshResults()
    {
        var selectedId = SelectedRow?.Command?.Id ?? SelectedRow?.Favorite?.FavoriteId;
        var availability = CommandDispatcher.CaptureAvailability();
        IReadOnlyList<CommandSearchResult> found;
        try { found = _index.Search(_query.Text, availability); }
        catch (ArgumentException exception) { _status.Text = exception.Message; return; }
        var favorites = ProfileRuntime.Current.Favorites;

        _results.BeginUpdate();
        try
        {
            _results.Items.Clear();
            foreach (var result in found)
            {
                var favorite = favorites.FirstOrDefault(value => value.CommandId == result.Command.Id && value.ContractVersion == result.Command.ContractVersion);
                AddRow(new SearchRow(result.Command, favorite, result.Availability));
            }
            foreach (var favorite in favorites)
            {
                if (BuiltInCommandRegistry.All.Any(value => value.Id == favorite.CommandId && value.ContractVersion == favorite.ContractVersion)) continue;
                if (!MatchesStaleFavorite(favorite, _query.Text)) continue;
                var resolution = FavoriteCatalog.Resolve(favorite, BuiltInCommandRegistry.All, availability);
                AddRow(new SearchRow(null, favorite, resolution.Availability));
            }
            if (_results.Items.Count > 0)
            {
                var match = _results.Items.Cast<ListViewItem>().FirstOrDefault(value =>
                {
                    var row = (SearchRow)value.Tag;
                    return string.Equals(row.Command?.Id ?? row.Favorite?.FavoriteId, selectedId, StringComparison.Ordinal);
                });
                (match ?? _results.Items[0]).Selected = true;
            }
        }
        finally { _results.EndUpdate(); }
        UpdateSelection();
    }

    private void AddRow(SearchRow row)
    {
        var name = row.Command?.DisplayName ?? $"Missing favorite: {row.Favorite!.CommandId}";
        if (row.Favorite is not null) name = "★ " + name;
        var item = new ListViewItem(name) { Tag = row };
        item.SubItems.Add(row.Command?.Category ?? "Unavailable");
        item.SubItems.Add(row.Command?.ShortcutLabel ?? string.Empty);
        item.SubItems.Add(row.Availability.Allowed ? "Available" : $"Unavailable: {row.Availability.Message}");
        _results.Items.Add(item);
    }

    private void UpdateSelection()
    {
        var row = SelectedRow;
        _run.Enabled = row is not null && row.Availability.Allowed && row.Command is not null;
        _favorite.Enabled = row?.Command is not null && !row.Command.Id.StartsWith("favorite.", StringComparison.Ordinal);
        _favorite.Text = row?.Favorite is null ? "Add Favorite" : "Remove Favorite";
        _status.Text = row is null ? "No command selected." : row.Availability.Allowed
            ? row.Command?.Description ?? "Available."
            : $"{row.Availability.Message} {row.Availability.Remediation}".Trim();
    }

    private void RunSelected()
    {
        var row = SelectedRow;
        if (row?.Command is null || !row.Availability.Allowed) return;
        if (row.Favorite is not null)
        {
            CallbackBoundary.Run("favorite.invoke", () => FavoriteCatalog.Invoke(row.Favorite, BuiltInCommandRegistry.All,
                CommandDispatcher.CaptureAvailability(), CommandDispatcher.InvokeRegistered));
        }
        else CallbackBoundary.Run(row.Command.Id, () => CommandDispatcher.InvokeRegistered(row.Command.Id, null, InvocationSource.Search));
        RefreshResults();
    }

    private void ToggleFavorite()
    {
        var row = SelectedRow;
        if (row?.Command is null) return;
        CallbackBoundary.Run(row.Favorite is null ? "favorite.add" : "favorite.remove", () =>
        {
            if (row.Favorite is null)
            {
                var added = ProfileRuntime.AddFavorite(new FavoriteDefinition("favorite." + row.Command.Id, row.Command.Id, row.Command.ContractVersion));
                return CommandResult.Success("favorite.add", added ? "Favorite added to the local profile." : "The favorite already exists.");
            }
            var removed = ProfileRuntime.RemoveFavorite(row.Favorite.FavoriteId);
            return CommandResult.Success("favorite.remove", removed ? "Favorite removed from the local profile." : "The favorite was already absent.");
        });
        RefreshResults();
    }

    private void OnQueryKeyDown(object? sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Down && _results.Items.Count > 0) { _results.Focus(); _results.Items[0].Selected = true; e.Handled = true; }
        else if (e.KeyCode == Keys.Enter) { RunSelected(); e.Handled = true; }
        else if (e.Control && e.KeyCode == Keys.D) { ToggleFavorite(); e.Handled = true; }
    }

    private void OnResultsKeyDown(object? sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter) { RunSelected(); e.Handled = true; }
        else if (e.Control && e.KeyCode == Keys.D) { ToggleFavorite(); e.Handled = true; }
        else if (e.Control && e.Shift && e.KeyCode == Keys.D && SelectedRow?.Favorite is not null) { ToggleFavorite(); e.Handled = true; }
    }

    private SearchRow? SelectedRow => _results.SelectedItems.Count == 0 ? null : (SearchRow)_results.SelectedItems[0].Tag;
    private static bool MatchesStaleFavorite(FavoriteDefinition favorite, string query) =>
        string.IsNullOrWhiteSpace(query) || favorite.FavoriteId.IndexOf(query.Trim(), StringComparison.OrdinalIgnoreCase) >= 0 ||
        favorite.CommandId.IndexOf(query.Trim(), StringComparison.OrdinalIgnoreCase) >= 0;

    private sealed class SearchRow
    {
        public SearchRow(CommandDescriptor? command, FavoriteDefinition? favorite, CanExecuteResult availability)
        { Command = command; Favorite = favorite; Availability = availability; }
        public CommandDescriptor? Command { get; }
        public FavoriteDefinition? Favorite { get; }
        public CanExecuteResult Availability { get; }
    }
}
