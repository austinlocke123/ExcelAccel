using ExcelAccel.Application.Navigation;

namespace ExcelAccel.ExcelAddIn;

internal static class NavigationRuntime
{
    public static NavigationSession Session { get; } = new NavigationSession();
    public static NavigationService Service { get; } = new NavigationService(Session);
    public static void Reset() => Session.Clear();
}
