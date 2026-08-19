using System;
using System.Runtime.InteropServices;

namespace ExcelAccel.ExcelAddIn.Interop;

internal static class ComRelease
{
    public static void Owned(object? value)
    {
        if (value is null || !Marshal.IsComObject(value))
        {
            return;
        }

        try
        {
            Marshal.ReleaseComObject(value);
        }
        catch (ArgumentException)
        {
            // The proxy was already released. Cleanup must never escape a callback.
        }
        catch (InvalidComObjectException)
        {
            // The proxy was already disconnected. Cleanup must never escape a callback.
        }
    }
}
