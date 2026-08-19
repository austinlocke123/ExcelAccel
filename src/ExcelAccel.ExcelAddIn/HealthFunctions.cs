using System;
using ExcelDna.Integration;
using ExcelAccel.ExcelAddIn.Reliability;

namespace ExcelAccel.ExcelAddIn;

public static class HealthFunctions
{
    [ExcelFunction(
        Name = "EXCELACCEL.VERSION",
        Description = "Returns the loaded ExcelAccel add-in version without reading or changing the workbook.",
        IsThreadSafe = true,
        IsExceptionSafe = true)]
    public static object Version()
    {
        try
        {
            return typeof(HealthFunctions).Assembly.GetName().Version?.ToString() ?? "unknown";
        }
        catch (Exception exception)
        {
            DiagnosticLog.Error("udf.version", exception);
            return ExcelError.ExcelErrorValue;
        }
    }
}
