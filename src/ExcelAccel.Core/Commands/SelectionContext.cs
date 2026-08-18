using System;

namespace ExcelAccel.Core.Commands;

public sealed class SelectionContext : IEquatable<SelectionContext>
{
    public SelectionContext(string workbookId, string worksheetName, string address)
    {
        WorkbookId = Require(workbookId, nameof(workbookId));
        WorksheetName = Require(worksheetName, nameof(worksheetName));
        Address = Require(address, nameof(address));
    }

    public string WorkbookId { get; }

    public string WorksheetName { get; }

    public string Address { get; }

    public bool Equals(SelectionContext? other) =>
        other is not null &&
        StringComparer.OrdinalIgnoreCase.Equals(WorkbookId, other.WorkbookId) &&
        StringComparer.Ordinal.Equals(WorksheetName, other.WorksheetName) &&
        StringComparer.Ordinal.Equals(Address, other.Address);

    public override bool Equals(object? obj) => Equals(obj as SelectionContext);

    public override int GetHashCode()
    {
        unchecked
        {
            var hash = StringComparer.OrdinalIgnoreCase.GetHashCode(WorkbookId);
            hash = (hash * 397) ^ StringComparer.Ordinal.GetHashCode(WorksheetName);
            return (hash * 397) ^ StringComparer.Ordinal.GetHashCode(Address);
        }
    }

    public override string ToString() => $"{WorkbookId}|{WorksheetName}|{Address}";

    private static string Require(string value, string parameterName)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw new ArgumentException("A non-empty value is required.", parameterName);
        }

        return value;
    }
}
