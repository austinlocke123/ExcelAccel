using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAccel.Core.Commands;

public sealed class CommandPlan
{
    public CommandPlan(
        string commandId,
        CommandImpact impact,
        SelectionContext context,
        IEnumerable<string> changedProperties,
        long affectedCellCount,
        string summary)
    {
        if (string.IsNullOrWhiteSpace(commandId))
        {
            throw new ArgumentException("A command ID is required.", nameof(commandId));
        }

        if (affectedCellCount < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(affectedCellCount));
        }

        CommandId = commandId;
        Impact = impact;
        Context = context ?? throw new ArgumentNullException(nameof(context));
        ChangedProperties = (changedProperties ?? throw new ArgumentNullException(nameof(changedProperties)))
            .Where(property => !string.IsNullOrWhiteSpace(property))
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        AffectedCellCount = affectedCellCount;
        Summary = summary ?? string.Empty;

        if (Impact == CommandImpact.ReadOnly && ChangedProperties.Count != 0)
        {
            throw new ArgumentException("A read-only command cannot declare changed properties.", nameof(changedProperties));
        }

        if (Impact != CommandImpact.ReadOnly && ChangedProperties.Count == 0)
        {
            throw new ArgumentException("A mutation must declare every changed property.", nameof(changedProperties));
        }
    }

    public string CommandId { get; }

    public CommandImpact Impact { get; }

    public SelectionContext Context { get; }

    public IReadOnlyList<string> ChangedProperties { get; }

    public long AffectedCellCount { get; }

    public string Summary { get; }
}
