using System;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace ExcelAccel.ExcelAddIn.Reliability;

internal sealed class ExcelWindowOwner : IWin32Window
{
    private ExcelWindowOwner(IntPtr handle) => Handle = handle;
    public IntPtr Handle { get; }
    public static ExcelWindowOwner? TryCreate()
    {
        try
        {
            dynamic application = ExcelDnaUtil.Application;
            var handle = new IntPtr(Convert.ToInt64(application.Hwnd));
            return handle == IntPtr.Zero ? null : new ExcelWindowOwner(handle);
        }
        catch { return null; }
    }
}
