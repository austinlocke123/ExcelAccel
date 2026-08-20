using System;
using System.Linq;

namespace ExcelAccel.Application.Navigation;

public sealed class NavigationService
{
    private readonly NavigationSession _session;
    public NavigationService(NavigationSession session) => _session = session ?? throw new ArgumentNullException(nameof(session));

    public bool MoveSheet(INavigationPort port, int delta, bool wrap)
    {
        if (delta != -1 && delta != 1) throw new ArgumentOutOfRangeException(nameof(delta));
        var origin = port.CaptureLocation();
        var sheets = port.GetVisibleWorksheetNames();
        var index = sheets.ToList().FindIndex(value => string.Equals(value, origin.WorksheetName, StringComparison.Ordinal));
        if (index < 0 || sheets.Count < 2) return false;
        var targetIndex = index + delta;
        if (targetIndex < 0 || targetIndex >= sheets.Count)
        {
            if (!wrap) return false;
            targetIndex = (targetIndex + sheets.Count) % sheets.Count;
        }

        return Navigate(port, origin, new NavigationLocation(origin.WorkbookId, sheets[targetIndex], origin.Address));
    }

    public bool Move(INavigationPort port, NavigationTargetKind kind)
    {
        var origin = port.CaptureLocation();
        return Navigate(port, origin, port.ResolveTarget(kind));
    }

    /// <summary>
    /// Navigates to an explicit location, revalidating it through the port and
    /// recording the prior location so session Back returns to it. Used by trace
    /// navigation, which is a separate action from analysis.
    /// </summary>
    public bool GoTo(INavigationPort port, NavigationLocation target)
    {
        if (port is null) throw new ArgumentNullException(nameof(port));
        if (target is null) throw new ArgumentNullException(nameof(target));
        return Navigate(port, port.CaptureLocation(), target);
    }

    public void AddBookmark(INavigationPort port) => _session.AddBookmark(port.CaptureLocation());
    public void ClearBookmarks() => _session.ClearBookmarks();
    public bool Back(INavigationPort port) => NavigateFromSession(port, _session.TryBack);
    public bool Forward(INavigationPort port) => NavigateFromSession(port, _session.TryForward);
    public bool NextBookmark(INavigationPort port) => NavigateBookmark(port, 1);
    public bool PreviousBookmark(INavigationPort port) => NavigateBookmark(port, -1);

    private bool Navigate(INavigationPort port, NavigationLocation origin, NavigationLocation target)
    {
        if (!port.TryNavigate(target)) return false;
        _session.Record(origin);
        _session.Record(target);
        return true;
    }

    private bool NavigateFromSession(INavigationPort port, TryGetLocation get)
    {
        if (!get(out var target) || target is null) return false;
        return port.TryNavigate(target);
    }

    private bool NavigateBookmark(INavigationPort port, int delta)
    {
        for (var attempt = 0; attempt < _session.BookmarkCount; attempt++)
        {
            if (!_session.TryMoveBookmark(delta, out var target) || target is null) return false;
            if (port.TryNavigate(target)) return true;
        }
        return false;
    }

    private delegate bool TryGetLocation(out NavigationLocation? location);
}
