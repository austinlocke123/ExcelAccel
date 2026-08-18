namespace ExcelAccel.Core.Reliability;

public interface IApplicationStatePort
{
    bool ScreenUpdating { get; set; }

    bool EnableEvents { get; set; }
}
