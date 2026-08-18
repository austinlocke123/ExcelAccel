using System;
using ExcelAccel.Core.Reliability;

namespace ExcelAccel.ExcelAddIn.Interop;

internal sealed class ExcelApplicationStateAdapter : IApplicationStatePort
{
    private readonly object _applicationObject;

    public ExcelApplicationStateAdapter(object applicationObject)
    {
        _applicationObject = applicationObject ?? throw new ArgumentNullException(nameof(applicationObject));
    }

    public bool ScreenUpdating
    {
        get
        {
            dynamic application = _applicationObject;
            return application.ScreenUpdating;
        }
        set
        {
            dynamic application = _applicationObject;
            application.ScreenUpdating = value;
        }
    }

    public bool EnableEvents
    {
        get
        {
            dynamic application = _applicationObject;
            return application.EnableEvents;
        }
        set
        {
            dynamic application = _applicationObject;
            application.EnableEvents = value;
        }
    }
}
