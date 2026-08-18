using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAccel.Core.Commands;

public sealed class CommandDescriptor
{
    public CommandDescriptor(
        string id,
        int contractVersion,
        string displayName,
        CommandImpact impact,
        IEnumerable<string> changedProperties,
        bool hasFixedParameters,
        string keyboardRoute)
    {
        if (string.IsNullOrWhiteSpace(id))
        {
            throw new ArgumentException("A command ID is required.", nameof(id));
        }

        if (contractVersion < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(contractVersion));
        }

        if (string.IsNullOrWhiteSpace(keyboardRoute))
        {
            throw new ArgumentException("Every command requires a keyboard route.", nameof(keyboardRoute));
        }

        Id = id;
        ContractVersion = contractVersion;
        DisplayName = displayName ?? throw new ArgumentNullException(nameof(displayName));
        Impact = impact;
        ChangedProperties = (changedProperties ?? throw new ArgumentNullException(nameof(changedProperties))).ToArray();
        HasFixedParameters = hasFixedParameters;
        KeyboardRoute = keyboardRoute;
    }

    public string Id { get; }

    public int ContractVersion { get; }

    public string DisplayName { get; }

    public CommandImpact Impact { get; }

    public IReadOnlyList<string> ChangedProperties { get; }

    public bool HasFixedParameters { get; }

    public string KeyboardRoute { get; }
}
