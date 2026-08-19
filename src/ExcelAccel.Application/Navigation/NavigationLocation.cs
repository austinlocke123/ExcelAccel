using System;

namespace ExcelAccel.Application.Navigation;

public sealed class NavigationLocation : IEquatable<NavigationLocation>
{
    public NavigationLocation(string workbookId, string worksheetName, string address)
    {
        WorkbookId = Require(workbookId, nameof(workbookId));
        WorksheetName = Require(worksheetName, nameof(worksheetName));
        Address = Require(address, nameof(address));
    }

    public string WorkbookId { get; }
    public string WorksheetName { get; }
    public string Address { get; }
    public bool Equals(NavigationLocation? other) => other is not null && WorkbookId == other.WorkbookId && WorksheetName == other.WorksheetName && Address == other.Address;
    public override bool Equals(object? obj) => Equals(obj as NavigationLocation);
    public override int GetHashCode() => (WorkbookId + "\0" + WorksheetName + "\0" + Address).GetHashCode();
    private static string Require(string value, string name) => string.IsNullOrWhiteSpace(value) ? throw new ArgumentException("A navigation identity component is required.", name) : value;
}

public enum NavigationTargetKind
{
    A1,
    UsedFirst,
    UsedLast,
    RegionEdgeUp,
    RegionEdgeDown,
    RegionEdgeLeft,
    RegionEdgeRight,
}

public interface INavigationPort
{
    NavigationLocation CaptureLocation();
    System.Collections.Generic.IReadOnlyList<string> GetVisibleWorksheetNames();
    NavigationLocation ResolveTarget(NavigationTargetKind target);
    bool TryNavigate(NavigationLocation target);
}
