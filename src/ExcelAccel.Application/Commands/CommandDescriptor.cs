using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Core.Commands;

namespace ExcelAccel.Application.Commands;

public sealed class CommandDescriptor
{
    public CommandDescriptor(
        string id,
        int contractVersion,
        string displayName,
        CommandImpact impact,
        IEnumerable<string> changedProperties,
        bool hasFixedParameters,
        string keyboardRoute,
        string capabilityId = "CAP-CMD-001",
        CommandContextRequirement contextRequirement = CommandContextRequirement.Selection,
        PreviewPolicy previewPolicy = PreviewPolicy.None,
        UndoPolicy undoPolicy = UndoPolicy.None,
        IEnumerable<string>? acceptanceIds = null,
        string category = "General",
        string? description = null,
        IEnumerable<string>? aliases = null,
        string? shortcutLabel = null,
        bool inheritsReferencedPolicy = false)
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
        ChangedProperties = (changedProperties ?? throw new ArgumentNullException(nameof(changedProperties)))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.Ordinal)
            .OrderBy(value => value, StringComparer.Ordinal)
            .ToArray();
        HasFixedParameters = hasFixedParameters;
        KeyboardRoute = keyboardRoute;
        CapabilityId = string.IsNullOrWhiteSpace(capabilityId)
            ? throw new ArgumentException("A capability ID is required.", nameof(capabilityId))
            : capabilityId;
        ContextRequirement = contextRequirement;
        PreviewPolicy = previewPolicy;
        UndoPolicy = undoPolicy;
        AcceptanceIds = (acceptanceIds ?? Array.Empty<string>())
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.Ordinal)
            .OrderBy(value => value, StringComparer.Ordinal)
            .ToArray();
        Category = string.IsNullOrWhiteSpace(category) ? "General" : category.Trim();
        Description = string.IsNullOrWhiteSpace(description) ? DisplayName : description!.Trim();
        Aliases = (aliases ?? Array.Empty<string>())
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => value.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        ShortcutLabel = string.IsNullOrWhiteSpace(shortcutLabel) ? KeyboardRoute : shortcutLabel!.Trim();
        InheritsReferencedPolicy = inheritsReferencedPolicy;

        if (Impact == CommandImpact.ReadOnly && ChangedProperties.Count != 0)
        {
            throw new ArgumentException("A read-only command cannot declare changed properties.", nameof(changedProperties));
        }
    }

    public string Id { get; }

    public int ContractVersion { get; }

    public string DisplayName { get; }

    public CommandImpact Impact { get; }

    public IReadOnlyList<string> ChangedProperties { get; }

    public bool HasFixedParameters { get; }

    public string KeyboardRoute { get; }

    public string CapabilityId { get; }

    public CommandContextRequirement ContextRequirement { get; }

    public PreviewPolicy PreviewPolicy { get; }

    public UndoPolicy UndoPolicy { get; }

    public IReadOnlyList<string> AcceptanceIds { get; }

    public string Category { get; }

    public string Description { get; }

    public IReadOnlyList<string> Aliases { get; }

    public string ShortcutLabel { get; }

    public bool InheritsReferencedPolicy { get; }
}
