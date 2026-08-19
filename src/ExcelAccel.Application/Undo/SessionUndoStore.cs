using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.Undo;

public sealed class PropertyReceipt
{
    public PropertyReceipt(string receiptId, string commandId, int commandVersion, SelectionContext target,
        string propertyId, string beforeValue, string afterValue, string planHash, DateTimeOffset createdUtc, DateTimeOffset expiresUtc)
    {
        ReceiptId = Require(receiptId, nameof(receiptId));
        CommandId = Require(commandId, nameof(commandId));
        if (commandVersion < 1) throw new ArgumentOutOfRangeException(nameof(commandVersion));
        CommandVersion = commandVersion;
        Target = target ?? throw new ArgumentNullException(nameof(target));
        PropertyId = Require(propertyId, nameof(propertyId));
        BeforeValue = beforeValue ?? string.Empty;
        AfterValue = afterValue ?? string.Empty;
        PlanHash = Require(planHash, nameof(planHash));
        CreatedUtc = createdUtc;
        ExpiresUtc = expiresUtc > createdUtc ? expiresUtc : throw new ArgumentOutOfRangeException(nameof(expiresUtc));
    }
    public string ReceiptId { get; }
    public string CommandId { get; }
    public int CommandVersion { get; }
    public SelectionContext Target { get; }
    public string PropertyId { get; }
    public string BeforeValue { get; }
    public string AfterValue { get; }
    public string PlanHash { get; }
    public DateTimeOffset CreatedUtc { get; }
    public DateTimeOffset ExpiresUtc { get; }
    private static string Require(string value, string name) => string.IsNullOrWhiteSpace(value) ? throw new ArgumentException("A receipt field is required.", name) : value;
}

public interface IPropertyReceiptSink { void Add(PropertyReceipt receipt); }
public interface IPropertyBatchReceiptSink { void Add(PropertyBatchReceipt receipt); }
public interface IPropertyReceiptPort
{
    bool TryRead(SelectionContext target, string propertyId, out string value);
    bool TryWrite(SelectionContext target, string propertyId, string value);
}

public enum UndoOutcome { Success, Empty, Expired, Stale, WriteFailed, VerificationFailed }

public sealed class PropertyChange
{
    public PropertyChange(string propertyId, string beforeValue, string afterValue)
    {
        PropertyId = string.IsNullOrWhiteSpace(propertyId) ? throw new ArgumentException("A property ID is required.", nameof(propertyId)) : propertyId;
        BeforeValue = beforeValue ?? string.Empty;
        AfterValue = afterValue ?? string.Empty;
    }
    public string PropertyId { get; }
    public string BeforeValue { get; }
    public string AfterValue { get; }
}

public sealed class PropertyBatchReceipt
{
    public const int MaximumChanges = 32;
    public PropertyBatchReceipt(string receiptId, string commandId, int commandVersion, SelectionContext target,
        IEnumerable<PropertyChange> changes, string planHash, DateTimeOffset createdUtc, DateTimeOffset expiresUtc)
    {
        ReceiptId = Require(receiptId, nameof(receiptId));
        CommandId = Require(commandId, nameof(commandId));
        if (commandVersion < 1) throw new ArgumentOutOfRangeException(nameof(commandVersion));
        CommandVersion = commandVersion;
        Target = target ?? throw new ArgumentNullException(nameof(target));
        var normalized = (changes ?? throw new ArgumentNullException(nameof(changes))).OrderBy(value => value.PropertyId, StringComparer.Ordinal).ToArray();
        if (normalized.Length < 1 || normalized.Length > MaximumChanges) throw new ArgumentException($"A batch receipt requires 1 through {MaximumChanges} changes.", nameof(changes));
        if (normalized.Select(value => value.PropertyId).Distinct(StringComparer.Ordinal).Count() != normalized.Length) throw new ArgumentException("Batch receipt property IDs must be unique.", nameof(changes));
        Changes = normalized;
        PlanHash = Require(planHash, nameof(planHash));
        CreatedUtc = createdUtc;
        ExpiresUtc = expiresUtc > createdUtc ? expiresUtc : throw new ArgumentOutOfRangeException(nameof(expiresUtc));
    }
    public string ReceiptId { get; }
    public string CommandId { get; }
    public int CommandVersion { get; }
    public SelectionContext Target { get; }
    public IReadOnlyList<PropertyChange> Changes { get; }
    public string PlanHash { get; }
    public DateTimeOffset CreatedUtc { get; }
    public DateTimeOffset ExpiresUtc { get; }
    private static string Require(string value, string name) => string.IsNullOrWhiteSpace(value) ? throw new ArgumentException("A receipt field is required.", name) : value;
}

public sealed class UndoResult
{
    public UndoResult(UndoOutcome outcome, string message, string receiptId = "") { Outcome = outcome; Message = message; ReceiptId = receiptId; }
    public UndoOutcome Outcome { get; }
    public bool Succeeded => Outcome == UndoOutcome.Success;
    public string Message { get; }
    public string ReceiptId { get; }
}

public sealed class SessionUndoStore : IPropertyReceiptSink, IPropertyBatchReceiptSink
{
    public const int MaximumReceiptsPerWorkbook = 20;
    public const int MaximumValueCharacters = 1_000_000;
    public const int MaximumValueCharactersPerWorkbook = 4_000_000;
    private readonly object _sync = new object();
    private readonly Dictionary<string, LinkedList<PropertyBatchReceipt>> _byWorkbook = new Dictionary<string, LinkedList<PropertyBatchReceipt>>(StringComparer.Ordinal);

    public int Count(string workbookId)
    {
        lock (_sync) return _byWorkbook.TryGetValue(workbookId, out var receipts) ? receipts.Count : 0;
    }

    public void Add(PropertyReceipt receipt)
    {
        if (receipt is null) throw new ArgumentNullException(nameof(receipt));
        Add(new PropertyBatchReceipt(receipt.ReceiptId, receipt.CommandId, receipt.CommandVersion, receipt.Target,
            new[] { new PropertyChange(receipt.PropertyId, receipt.BeforeValue, receipt.AfterValue) }, receipt.PlanHash,
            receipt.CreatedUtc, receipt.ExpiresUtc));
    }

    public void Add(PropertyBatchReceipt receipt)
    {
        if (receipt is null) throw new ArgumentNullException(nameof(receipt));
        if (receipt.Changes.Any(value => value.BeforeValue.Length > MaximumValueCharacters || value.AfterValue.Length > MaximumValueCharacters))
            throw new ArgumentException("Receipt values exceed the in-memory bound.", nameof(receipt));
        lock (_sync)
        {
            if (!_byWorkbook.TryGetValue(receipt.Target.WorkbookId, out var receipts))
                _byWorkbook.Add(receipt.Target.WorkbookId, receipts = new LinkedList<PropertyBatchReceipt>());
            receipts.AddLast(receipt);
            while (receipts.Count > MaximumReceiptsPerWorkbook || ReceiptCharacters(receipts) > MaximumValueCharactersPerWorkbook)
                receipts.RemoveFirst();
        }
    }

    public UndoResult TryUndo(string workbookId, IPropertyReceiptPort port, DateTimeOffset now)
    {
        if (port is null) throw new ArgumentNullException(nameof(port));
        PropertyBatchReceipt? receipt;
        lock (_sync)
        {
            if (!_byWorkbook.TryGetValue(workbookId, out var receipts) || receipts.Last is null)
                return new UndoResult(UndoOutcome.Empty, "No eligible ExcelAccel receipt exists for this workbook.");
            receipt = receipts.Last.Value;
            receipts.RemoveLast();
        }
        if (now > receipt.ExpiresUtc) return new UndoResult(UndoOutcome.Expired, "The latest receipt expired and was discarded.", receipt.ReceiptId);
        foreach (var change in receipt.Changes)
            if (!port.TryRead(receipt.Target, change.PropertyId, out var current) || !ValuesMatch(change.PropertyId, current, change.AfterValue))
                return new UndoResult(UndoOutcome.Stale, "Undo refused because the target or a receipt property changed after the command.", receipt.ReceiptId);

        var restored = new List<PropertyChange>();
        foreach (var change in receipt.Changes.Reverse())
        {
            if (port.TryWrite(receipt.Target, change.PropertyId, change.BeforeValue)) { restored.Add(change); continue; }
            // A failed port write may still have partially mutated Excel. Include
            // that property in the compensating set rather than assuming failure
            // means no write occurred.
            var compensating = restored.Concat(new[] { change }).ToArray();
            var rollbackComplete = compensating.All(value => port.TryWrite(receipt.Target, value.PropertyId, value.AfterValue)) &&
                compensating.All(value => port.TryRead(receipt.Target, value.PropertyId, out var observedAfter) && ValuesMatch(value.PropertyId, observedAfter, value.AfterValue));
            return new UndoResult(UndoOutcome.WriteFailed, rollbackComplete
                ? "Undo could not write the complete before-state; already restored properties were returned to the post-command state."
                : "Undo failed and could not fully return already restored properties to the post-command state; inspect the target.", receipt.ReceiptId);
        }
        foreach (var change in receipt.Changes)
            if (!port.TryRead(receipt.Target, change.PropertyId, out var observed) || !ValuesMatch(change.PropertyId, observed, change.BeforeValue))
                return new UndoResult(UndoOutcome.VerificationFailed, "Undo postcondition verification failed for one or more receipt properties; inspect the target.", receipt.ReceiptId);
        return new UndoResult(UndoOutcome.Success, $"Restored {receipt.Changes.Count} unchanged ExcelAccel propert{(receipt.Changes.Count == 1 ? "y" : "ies")}.", receipt.ReceiptId);
    }

    public void ClearWorkbook(string workbookId) { lock (_sync) _byWorkbook.Remove(workbookId); }
    public void ClearAll() { lock (_sync) _byWorkbook.Clear(); }

    private static int ReceiptCharacters(IEnumerable<PropertyBatchReceipt> receipts) =>
        receipts.Sum(receipt => receipt.Changes.Sum(change => checked(change.BeforeValue.Length + change.AfterValue.Length)));

    private static bool ValuesMatch(string propertyId, string first, string second) =>
        string.Equals(first, second,
            string.Equals(propertyId, "cell_contents_v1", StringComparison.Ordinal) ||
            string.Equals(propertyId, "cell_format_block_v1", StringComparison.Ordinal)
                ? StringComparison.Ordinal
                : StringComparison.OrdinalIgnoreCase);
}
