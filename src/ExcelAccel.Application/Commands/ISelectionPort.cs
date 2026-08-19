using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.Commands;

public interface ISelectionPort
{
    SelectionSnapshot CaptureSelection();

    void SetNumberFormat(string formatCode);
}
