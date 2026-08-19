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
public interface IPropertyReceiptPort
{
    bool TryRead(SelectionContext target, string propertyId, out string value);
    bool TryWrite(SelectionContext target, string propertyId, string value);
}

public enum UndoOutcome { Success, Empty, Expired, Stale, WriteFailed, VerificationFailed }

public sealed class UndoResult
{
    public UndoResult(UndoOutcome outcome, string message, string receiptId = "") { Outcome = outcome; Message = message; ReceiptId = receiptId; }
    public UndoOutcome Outcome { get; }
    public bool Succeeded => Outcome == UndoOutcome.Success;
    public string Message { get; }
    public string ReceiptId { get; }
}

public sealed class SessionUndoStore : IPropertyReceiptSink
{
    public const int MaximumReceiptsPerWorkbook = 20;
    public const int MaximumValueCharacters = 65_536;
    private readonly object _sync = new object();
    private readonly Dictionary<string, LinkedList<PropertyReceipt>> _byWorkbook = new Dictionary<string, LinkedList<PropertyReceipt>>(StringComparer.Ordinal);

    public int Count(string workbookId)
    {
        lock (_sync) return _byWorkbook.TryGetValue(workbookId, out var receipts) ? receipts.Count : 0;
    }

    public void Add(PropertyReceipt receipt)
    {
        if (receipt is null) throw new ArgumentNullException(nameof(receipt));
        if (receipt.BeforeValue.Length > MaximumValueCharacters || receipt.AfterValue.Length > MaximumValueCharacters)
            throw new ArgumentException("Receipt values exceed the in-memory bound.", nameof(receipt));
        lock (_sync)
        {
            if (!_byWorkbook.TryGetValue(receipt.Target.WorkbookId, out var receipts))
                _byWorkbook.Add(receipt.Target.WorkbookId, receipts = new LinkedList<PropertyReceipt>());
            receipts.AddLast(receipt);
            while (receipts.Count > MaximumReceiptsPerWorkbook) receipts.RemoveFirst();
        }
    }

    public UndoResult TryUndo(string workbookId, IPropertyReceiptPort port, DateTimeOffset now)
    {
        if (port is null) throw new ArgumentNullException(nameof(port));
        PropertyReceipt? receipt;
        lock (_sync)
        {
            if (!_byWorkbook.TryGetValue(workbookId, out var receipts) || receipts.Last is null)
                return new UndoResult(UndoOutcome.Empty, "No eligible ExcelAccel receipt exists for this workbook.");
            receipt = receipts.Last.Value;
            receipts.RemoveLast();
        }
        if (now > receipt.ExpiresUtc) return new UndoResult(UndoOutcome.Expired, "The latest receipt expired and was discarded.", receipt.ReceiptId);
        if (!port.TryRead(receipt.Target, receipt.PropertyId, out var current) || !string.Equals(current, receipt.AfterValue, StringComparison.OrdinalIgnoreCase))
            return new UndoResult(UndoOutcome.Stale, "Undo refused because the target or property changed after the command.", receipt.ReceiptId);
        if (!port.TryWrite(receipt.Target, receipt.PropertyId, receipt.BeforeValue))
            return new UndoResult(UndoOutcome.WriteFailed, "Undo could not write the qualified before-state.", receipt.ReceiptId);
        if (!port.TryRead(receipt.Target, receipt.PropertyId, out var observed) || !string.Equals(observed, receipt.BeforeValue, StringComparison.OrdinalIgnoreCase))
            return new UndoResult(UndoOutcome.VerificationFailed, "Undo postcondition verification failed; inspect the target.", receipt.ReceiptId);
        return new UndoResult(UndoOutcome.Success, "Restored the latest unchanged ExcelAccel property.", receipt.ReceiptId);
    }

    public void ClearWorkbook(string workbookId) { lock (_sync) _byWorkbook.Remove(workbookId); }
    public void ClearAll() { lock (_sync) _byWorkbook.Clear(); }
}
