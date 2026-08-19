using System;
using System.Globalization;
using Microsoft.CSharp.RuntimeBinder;
using ExcelAccel.Application.Commands;

namespace ExcelAccel.ExcelInterop;

internal static class ExcelCommandReadiness
{
    public static void RequireReady(object applicationObject)
    {
        if (applicationObject is null)
        {
            throw new ArgumentNullException(nameof(applicationObject));
        }

        try
        {
            dynamic application = applicationObject;
            var ready = Convert.ToBoolean(application.Ready, CultureInfo.InvariantCulture);
            if (!ready)
            {
                throw new CommandRefusedException(
                    RefusalCodes.EditMode,
                    "Excel is editing a cell, calculating, or otherwise not ready for an add-in command.",
                    "Finish or cancel the current Excel operation and try again.");
            }
        }
        catch (CommandRefusedException)
        {
            throw;
        }
        catch (RuntimeBinderException exception)
        {
            throw new CommandRefusedException(
                RefusalCodes.CommandUnavailable,
                $"Excel readiness could not be determined safely: {exception.Message}",
                "Finish the current Excel operation and try again.");
        }
    }
}
