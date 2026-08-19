using System;
using System.Collections.Generic;

namespace ExcelAccel.Application.Navigation;

public sealed class NavigationSession
{
    public const int MaximumHistory = 100;
    public const int MaximumBookmarks = 50;
    private readonly List<NavigationLocation> _history = new List<NavigationLocation>();
    private readonly List<NavigationLocation> _bookmarks = new List<NavigationLocation>();
    private int _historyIndex = -1;
    private int _bookmarkIndex = -1;

    public int HistoryCount => _history.Count;
    public int BookmarkCount => _bookmarks.Count;

    public void Record(NavigationLocation location)
    {
        if (_historyIndex >= 0 && _history[_historyIndex].Equals(location)) return;
        if (_historyIndex < _history.Count - 1) _history.RemoveRange(_historyIndex + 1, _history.Count - _historyIndex - 1);
        _history.Add(location ?? throw new ArgumentNullException(nameof(location)));
        if (_history.Count > MaximumHistory) _history.RemoveAt(0);
        _historyIndex = _history.Count - 1;
    }

    public bool TryBack(out NavigationLocation? location) => TryMoveHistory(-1, out location);
    public bool TryForward(out NavigationLocation? location) => TryMoveHistory(1, out location);

    public bool AddBookmark(NavigationLocation location)
    {
        if (_bookmarks.Exists(value => value.Equals(location))) return false;
        _bookmarks.Add(location ?? throw new ArgumentNullException(nameof(location)));
        if (_bookmarks.Count > MaximumBookmarks) _bookmarks.RemoveAt(0);
        _bookmarkIndex = _bookmarks.Count - 1;
        return true;
    }

    public bool TryMoveBookmark(int delta, out NavigationLocation? location)
    {
        location = null;
        if (_bookmarks.Count == 0 || (delta != -1 && delta != 1)) return false;
        _bookmarkIndex = (_bookmarkIndex + delta + _bookmarks.Count) % _bookmarks.Count;
        location = _bookmarks[_bookmarkIndex];
        return true;
    }

    public void ClearBookmarks()
    {
        _bookmarks.Clear();
        _bookmarkIndex = -1;
    }

    public void ClearWorkbook(string workbookId)
    {
        if (string.IsNullOrWhiteSpace(workbookId)) return;
        _history.RemoveAll(value => string.Equals(value.WorkbookId, workbookId, StringComparison.Ordinal));
        _bookmarks.RemoveAll(value => string.Equals(value.WorkbookId, workbookId, StringComparison.Ordinal));
        _historyIndex = _history.Count - 1;
        _bookmarkIndex = _bookmarks.Count - 1;
    }

    public void Clear()
    {
        _history.Clear();
        _bookmarks.Clear();
        _historyIndex = -1;
        _bookmarkIndex = -1;
    }

    private bool TryMoveHistory(int delta, out NavigationLocation? location)
    {
        location = null;
        var candidate = _historyIndex + delta;
        if (candidate < 0 || candidate >= _history.Count) return false;
        _historyIndex = candidate;
        location = _history[candidate];
        return true;
    }
}
