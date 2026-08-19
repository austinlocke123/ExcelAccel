using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Profiles;

namespace ExcelAccel.Application.QuickKeys;

public enum QuickKeyOutcome
{
    NotHandled = 0,
    AwaitingNextStroke = 1,
    Invoke = 2,
    Cancelled = 3,
    TimedOut = 4,
}

public sealed class QuickKeyResult
{
    public QuickKeyResult(QuickKeyOutcome outcome, string commandId = "")
    {
        Outcome = outcome;
        CommandId = commandId ?? string.Empty;
    }

    public QuickKeyOutcome Outcome { get; }
    public string CommandId { get; }
}

public sealed class QuickKeyConflict
{
    public QuickKeyConflict(string sequence, string reason)
    {
        Sequence = sequence;
        Reason = reason;
    }

    public string Sequence { get; }
    public string Reason { get; }
}

public static class QuickKeyValidator
{
    private static readonly HashSet<string> Reserved = new HashSet<string>(StringComparer.Ordinal)
    {
        "CTRL+A", "CTRL+C", "CTRL+F", "CTRL+H", "CTRL+K", "CTRL+P", "CTRL+S",
        "CTRL+V", "CTRL+X", "CTRL+Y", "CTRL+Z", "ALT+F4", "SHIFT+F10", "F1",
        "F2", "F4", "F5", "F7", "F9", "F10", "F11", "F12",
    };

    public static IReadOnlyList<QuickKeyConflict> Validate(IEnumerable<QuickKeyBinding> bindings)
    {
        var normalized = (bindings ?? throw new ArgumentNullException(nameof(bindings)))
            .Select(binding => new { Binding = binding, Sequence = Normalize(binding.Sequence) })
            .ToArray();
        var conflicts = new List<QuickKeyConflict>();

        foreach (var item in normalized)
        {
            var strokes = Split(item.Sequence);
            if (strokes.Length == 0 || strokes.Length > 3)
            {
                conflicts.Add(new QuickKeyConflict(item.Binding.Sequence, "Sequences require one through three strokes."));
                continue;
            }

            if (strokes.Any(stroke => Reserved.Contains(stroke)))
            {
                conflicts.Add(new QuickKeyConflict(item.Binding.Sequence, "The sequence contains a reserved Excel, Windows, or accessibility shortcut."));
            }
        }

        foreach (var group in normalized.GroupBy(value => value.Sequence, StringComparer.Ordinal).Where(group => group.Count() > 1))
        {
            conflicts.Add(new QuickKeyConflict(group.Key, "The sequence is assigned more than once."));
        }

        for (var left = 0; left < normalized.Length; left++)
        {
            for (var right = left + 1; right < normalized.Length; right++)
            {
                var first = normalized[left].Sequence;
                var second = normalized[right].Sequence;
                if (first != second && (second.StartsWith(first + ",", StringComparison.Ordinal) || first.StartsWith(second + ",", StringComparison.Ordinal)))
                {
                    conflicts.Add(new QuickKeyConflict($"{first} / {second}", "One assignment is a prefix of another."));
                }
            }
        }

        return conflicts;
    }

    public static string Normalize(string sequence) => string.Join(",", Split(sequence));

    internal static string[] Split(string sequence) =>
        (sequence ?? string.Empty)
            .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(stroke => string.Join("+", stroke
                .Split(new[] { '+' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(token => token.Trim().ToUpperInvariant())
                .Where(token => token.Length > 0)))
            .Where(stroke => stroke.Length > 0)
            .ToArray();
}

public sealed class QuickKeyEngine
{
    private readonly IReadOnlyDictionary<string, string> _bindings;
    private readonly TimeSpan _timeout;
    private string _pending = string.Empty;
    private DateTimeOffset _expiresAt;

    public QuickKeyEngine(IEnumerable<QuickKeyBinding> bindings, TimeSpan timeout)
    {
        if (timeout < TimeSpan.FromMilliseconds(250) || timeout > TimeSpan.FromSeconds(10))
        {
            throw new ArgumentOutOfRangeException(nameof(timeout));
        }

        var bindingArray = (bindings ?? throw new ArgumentNullException(nameof(bindings))).ToArray();
        var conflicts = QuickKeyValidator.Validate(bindingArray);
        if (conflicts.Count != 0)
        {
            throw new ArgumentException(conflicts[0].Reason, nameof(bindings));
        }

        _bindings = bindingArray.ToDictionary(
            binding => QuickKeyValidator.Normalize(binding.Sequence),
            binding => binding.CommandId,
            StringComparer.Ordinal);
        _timeout = timeout;
    }

    public QuickKeyResult ProcessStroke(string stroke, DateTimeOffset now, bool excelEditMode)
    {
        if (excelEditMode)
        {
            Reset();
            return new QuickKeyResult(QuickKeyOutcome.NotHandled);
        }

        var normalizedStroke = QuickKeyValidator.Normalize(stroke);
        if (normalizedStroke.Contains(","))
        {
            Reset();
            return new QuickKeyResult(QuickKeyOutcome.NotHandled);
        }

        if (string.Equals(normalizedStroke, "ESCAPE", StringComparison.Ordinal))
        {
            var wasPending = _pending.Length != 0;
            Reset();
            return new QuickKeyResult(wasPending ? QuickKeyOutcome.Cancelled : QuickKeyOutcome.NotHandled);
        }

        if (_pending.Length != 0 && now > _expiresAt)
        {
            Reset();
            return new QuickKeyResult(QuickKeyOutcome.TimedOut);
        }

        var candidate = _pending.Length == 0 ? normalizedStroke : _pending + "," + normalizedStroke;
        if (_bindings.TryGetValue(candidate, out var commandId))
        {
            Reset();
            return new QuickKeyResult(QuickKeyOutcome.Invoke, commandId);
        }

        if (_bindings.Keys.Any(sequence => sequence.StartsWith(candidate + ",", StringComparison.Ordinal)))
        {
            _pending = candidate;
            _expiresAt = now.Add(_timeout);
            return new QuickKeyResult(QuickKeyOutcome.AwaitingNextStroke);
        }

        Reset();
        return new QuickKeyResult(QuickKeyOutcome.NotHandled);
    }

    public void Reset()
    {
        _pending = string.Empty;
        _expiresAt = default;
    }
}
