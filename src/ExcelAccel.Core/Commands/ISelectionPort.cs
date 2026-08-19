namespace ExcelAccel.Core.Commands;

public interface ISelectionPort
{
    SelectionSnapshot CaptureSelection();

    void SetNumberFormat(string formatCode);
}
