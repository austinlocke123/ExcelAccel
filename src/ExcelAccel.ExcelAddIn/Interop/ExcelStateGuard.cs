using System;
using ExcelAccel.ExcelAddIn.Reliability;

namespace ExcelAccel.ExcelAddIn.Interop;

internal sealed class ExcelStateGuard : IDisposable
{
    private readonly object _applicationObject;
    private readonly bool _originalScreenUpdating;
    private readonly bool _ownsScreenUpdating;
    private bool _disposed;

    private ExcelStateGuard(object applicationObject, bool originalScreenUpdating, bool ownsScreenUpdating)
    {
        _applicationObject = applicationObject;
        _originalScreenUpdating = originalScreenUpdating;
        _ownsScreenUpdating = ownsScreenUpdating;
    }

    public static ExcelStateGuard SuppressScreenUpdating(object applicationObject)
    {
        if (applicationObject is null)
        {
            throw new ArgumentNullException(nameof(applicationObject));
        }

        dynamic application = applicationObject;
        bool original = application.ScreenUpdating;
        bool owns = original;
        if (owns)
        {
            application.ScreenUpdating = false;
        }

        return new ExcelStateGuard(applicationObject, original, owns);
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        if (!_ownsScreenUpdating)
        {
            return;
        }

        try
        {
            dynamic application = _applicationObject;
            application.ScreenUpdating = _originalScreenUpdating;
        }
        catch (Exception exception)
        {
            DiagnosticLog.Error("excel.state.restore", exception);
        }
    }
}
