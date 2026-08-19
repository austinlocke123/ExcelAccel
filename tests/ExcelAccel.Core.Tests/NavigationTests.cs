using System.Collections.Generic;
using ExcelAccel.Application.Navigation;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class NavigationTests
{
    [Fact]
    public void VisibleSheetNavigationWrapsOnlyWhenProfileAllows()
    {
        var session = new NavigationSession();
        var service = new NavigationService(session);
        var port = new FakeNavigationPort("Sheet3");

        Assert.False(service.MoveSheet(port, 1, wrap: false));
        Assert.True(service.MoveSheet(port, 1, wrap: true));
        Assert.Equal("Sheet1", port.Current.WorksheetName);
    }

    [Fact]
    public void NavigationHistoryIsBoundedAndDropsForwardBranch()
    {
        var session = new NavigationSession();
        for (var index = 0; index < NavigationSession.MaximumHistory + 10; index++)
            session.Record(new NavigationLocation("Book.xlsx", "Sheet1", $"A{index + 1}"));
        Assert.Equal(NavigationSession.MaximumHistory, session.HistoryCount);

        Assert.True(session.TryBack(out _));
        session.Record(new NavigationLocation("Book.xlsx", "Sheet1", "Z1"));
        Assert.False(session.TryForward(out _));
    }

    [Fact]
    public void BookmarksAreSessionOnlyBoundedAndSkipStaleTargets()
    {
        var session = new NavigationSession();
        var service = new NavigationService(session);
        var port = new FakeNavigationPort("Sheet1");
        service.AddBookmark(port);
        port.Current = new NavigationLocation("Book.xlsx", "Sheet2", "A1");
        service.AddBookmark(port);
        port.RefusedSheet = "Sheet1";

        Assert.True(service.NextBookmark(port));
        Assert.Equal("Sheet2", port.Current.WorksheetName);
        service.ClearBookmarks();
        Assert.Equal(0, session.BookmarkCount);
    }

    [Fact]
    public void A1NavigationRecordsOriginAndSupportsBack()
    {
        var session = new NavigationSession();
        var service = new NavigationService(session);
        var port = new FakeNavigationPort("Sheet1") { Current = new NavigationLocation("Book.xlsx", "Sheet1", "D9") };

        Assert.True(service.Move(port, NavigationTargetKind.A1));
        Assert.Equal("A1", port.Current.Address);
        Assert.True(service.Back(port));
        Assert.Equal("D9", port.Current.Address);
    }

    private sealed class FakeNavigationPort : INavigationPort
    {
        public FakeNavigationPort(string sheet) => Current = new NavigationLocation("Book.xlsx", sheet, "A1");
        public NavigationLocation Current { get; set; }
        public string? RefusedSheet { get; set; }
        public NavigationLocation CaptureLocation() => Current;
        public IReadOnlyList<string> GetVisibleWorksheetNames() => new[] { "Sheet1", "Sheet2", "Sheet3" };
        public NavigationLocation ResolveTarget(NavigationTargetKind target) => new NavigationLocation(Current.WorkbookId, Current.WorksheetName, target == NavigationTargetKind.A1 ? "A1" : "B2");
        public bool TryNavigate(NavigationLocation target)
        {
            if (target.WorksheetName == RefusedSheet) return false;
            Current = target;
            return true;
        }
    }
}
