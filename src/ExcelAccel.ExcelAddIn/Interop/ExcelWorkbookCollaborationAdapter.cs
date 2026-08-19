using System;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelAccel.Core.Collaboration;

namespace ExcelAccel.ExcelAddIn.Interop;

internal static class ExcelWorkbookCollaborationAdapter
{
    public static WorkbookCollaborationState Capture(object workbook)
    {
        if (workbook is null)
        {
            throw new ArgumentNullException(nameof(workbook));
        }

        var autoSaveOn = TryReadBoolean(workbook, "AutoSaveOn");
        var legacyShared = TryReadBoolean(workbook, "MultiUserEditing");
        var path = TryReadString(workbook, "Path");
        var location = ClassifyLocation(path);

        return WorkbookCollaborationClassifier.Classify(
            new WorkbookCollaborationProbe(
                autoSaveOn,
                legacyShared,
                location,
                remoteChangeEventsHooked: false));
    }

    private static bool? TryReadBoolean(object target, string propertyName)
    {
        try
        {
            var value = target.GetType().InvokeMember(
                propertyName,
                BindingFlags.GetProperty,
                binder: null,
                target,
                args: null,
                CultureInfo.InvariantCulture);
            return Convert.ToBoolean(value, CultureInfo.InvariantCulture);
        }
        catch (MissingMethodException)
        {
            return null;
        }
        catch (COMException)
        {
            return null;
        }
    }

    private static string? TryReadString(object target, string propertyName)
    {
        try
        {
            var value = target.GetType().InvokeMember(
                propertyName,
                BindingFlags.GetProperty,
                binder: null,
                target,
                args: null,
                CultureInfo.InvariantCulture);
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }
        catch (MissingMethodException)
        {
            return null;
        }
        catch (COMException)
        {
            return null;
        }
    }

    private static WorkbookLocationKind ClassifyLocation(string? path)
    {
        if (path is null)
        {
            return WorkbookLocationKind.Unknown;
        }

        if (path.Length == 0)
        {
            return WorkbookLocationKind.Unsaved;
        }

        if (Uri.TryCreate(path, UriKind.Absolute, out var uri) &&
            (uri.Scheme.Equals(Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) ||
             uri.Scheme.Equals(Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)))
        {
            return WorkbookLocationKind.CloudUrl;
        }

        return WorkbookLocationKind.LocalOrSyncedPath;
    }
}
