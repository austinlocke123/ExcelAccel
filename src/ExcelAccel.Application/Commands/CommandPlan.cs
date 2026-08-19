using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.Commands;

public sealed class CommandPlan
{
    public CommandPlan(
        string commandId,
        CommandImpact impact,
        SelectionContext context,
        IEnumerable<string> changedProperties,
        long affectedCellCount,
        string summary,
        string preconditionFingerprint = "",
        int contractVersion = 1,
        bool requiresPreview = false,
        IEnumerable<KeyValuePair<string, string>>? arguments = null)
    {
        if (string.IsNullOrWhiteSpace(commandId))
        {
            throw new ArgumentException("A command ID is required.", nameof(commandId));
        }

        if (affectedCellCount < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(affectedCellCount));
        }

        if (contractVersion < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(contractVersion));
        }

        CommandId = commandId;
        Impact = impact;
        Context = context ?? throw new ArgumentNullException(nameof(context));
        ChangedProperties = (changedProperties ?? throw new ArgumentNullException(nameof(changedProperties)))
            .Where(property => !string.IsNullOrWhiteSpace(property))
            .Distinct(StringComparer.Ordinal)
            .OrderBy(property => property, StringComparer.Ordinal)
            .ToArray();
        AffectedCellCount = affectedCellCount;
        Summary = summary ?? string.Empty;
        PreconditionFingerprint = preconditionFingerprint ?? throw new ArgumentNullException(nameof(preconditionFingerprint));
        ContractVersion = contractVersion;
        RequiresPreview = requiresPreview;
        var normalizedArguments = new SortedDictionary<string, string>(StringComparer.Ordinal);
        if (arguments is not null)
        {
            foreach (var argument in arguments)
            {
                if (string.IsNullOrWhiteSpace(argument.Key))
                {
                    throw new ArgumentException("Argument keys cannot be empty.", nameof(arguments));
                }

                if (normalizedArguments.ContainsKey(argument.Key))
                {
                    throw new ArgumentException($"Duplicate argument key '{argument.Key}'.", nameof(arguments));
                }

                normalizedArguments.Add(argument.Key, argument.Value ?? string.Empty);
            }
        }
        Arguments = normalizedArguments;

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

    public string PreconditionFingerprint { get; }

    public int ContractVersion { get; }

    public bool RequiresPreview { get; }

    public IReadOnlyDictionary<string, string> Arguments { get; }

    public string CanonicalJson => CanonicalPlanSerializer.Serialize(this);

    public string PlanHash => CanonicalPlanSerializer.Hash(this);
}
