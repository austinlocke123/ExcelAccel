using System;
using System.Collections.Generic;
using System.Linq;
using ExcelAccel.Application.Commands;

namespace ExcelAccel.Application.Profiles;

public sealed class FavoriteDefinition
{
    public const int MaximumArguments = 16;

    public FavoriteDefinition(string favoriteId, string commandId, int contractVersion,
        IEnumerable<KeyValuePair<string, string>>? arguments = null)
    {
        FavoriteId = RequireToken(favoriteId, nameof(favoriteId), 128);
        CommandId = RequireToken(commandId, nameof(commandId), 128);
        if (contractVersion < 1) throw new ArgumentOutOfRangeException(nameof(contractVersion));
        ContractVersion = contractVersion;
        var normalized = new SortedDictionary<string, string>(StringComparer.Ordinal);
        foreach (var argument in arguments ?? Array.Empty<KeyValuePair<string, string>>())
        {
            var key = RequireToken(argument.Key, nameof(arguments), 64);
            var value = argument.Value ?? string.Empty;
            if (value.Length > 1024) throw new ArgumentException("Favorite argument values may not exceed 1,024 characters.", nameof(arguments));
            if (normalized.ContainsKey(key)) throw new ArgumentException($"Duplicate favorite argument '{key}'.", nameof(arguments));
            normalized.Add(key, value);
        }
        if (normalized.Count > MaximumArguments) throw new ArgumentException($"Favorites may contain at most {MaximumArguments} arguments.", nameof(arguments));
        Arguments = normalized;
    }

    public string FavoriteId { get; }
    public string CommandId { get; }
    public int ContractVersion { get; }
    public IReadOnlyDictionary<string, string> Arguments { get; }

    public bool SemanticallyEquals(FavoriteDefinition other) =>
        other is not null && string.Equals(FavoriteId, other.FavoriteId, StringComparison.Ordinal) &&
        string.Equals(CommandId, other.CommandId, StringComparison.Ordinal) && ContractVersion == other.ContractVersion &&
        Arguments.Count == other.Arguments.Count && Arguments.All(item => other.Arguments.TryGetValue(item.Key, out var value) && string.Equals(item.Value, value, StringComparison.Ordinal));

    private static string RequireToken(string value, string parameterName, int maximumLength)
    {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("A nonempty invariant token is required.", parameterName);
        var result = value.Trim();
        if (result.Length > maximumLength) throw new ArgumentException($"The token may not exceed {maximumLength} characters.", parameterName);
        return result;
    }
}

public enum FavoriteResolutionStatus { Available = 0, Unavailable = 1, MissingCommand = 2, IncompatibleVersion = 3, InvalidArguments = 4 }

public sealed class FavoriteResolution
{
    public FavoriteResolution(FavoriteDefinition favorite, FavoriteResolutionStatus status, CommandDescriptor? command,
        CanExecuteResult availability, string remediation)
    {
        Favorite = favorite; Status = status; Command = command; Availability = availability; Remediation = remediation ?? string.Empty;
    }
    public FavoriteDefinition Favorite { get; }
    public FavoriteResolutionStatus Status { get; }
    public CommandDescriptor? Command { get; }
    public CanExecuteResult Availability { get; }
    public string Remediation { get; }
    public bool CanInvoke => Status == FavoriteResolutionStatus.Available && Availability.Allowed;
}

public static class FavoriteCatalog
{
    public static IReadOnlyList<FavoriteDefinition> Add(IReadOnlyList<FavoriteDefinition> current, FavoriteDefinition favorite)
    {
        if (current is null) throw new ArgumentNullException(nameof(current));
        if (favorite is null) throw new ArgumentNullException(nameof(favorite));
        var existing = current.FirstOrDefault(value => string.Equals(value.FavoriteId, favorite.FavoriteId, StringComparison.Ordinal));
        if (existing is not null)
        {
            if (!existing.SemanticallyEquals(favorite)) throw new InvalidOperationException($"Favorite ID '{favorite.FavoriteId}' already refers to different content.");
            return current;
        }
        return current.Concat(new[] { favorite }).OrderBy(value => value.FavoriteId, StringComparer.Ordinal).ToArray();
    }

    public static IReadOnlyList<FavoriteDefinition> Remove(IReadOnlyList<FavoriteDefinition> current, string favoriteId)
    {
        if (current is null) throw new ArgumentNullException(nameof(current));
        if (string.IsNullOrWhiteSpace(favoriteId)) throw new ArgumentException("A favorite ID is required.", nameof(favoriteId));
        return current.Where(value => !string.Equals(value.FavoriteId, favoriteId, StringComparison.Ordinal)).ToArray();
    }

    public static FavoriteResolution Resolve(FavoriteDefinition favorite, IReadOnlyList<CommandDescriptor> registry,
        Func<CommandDescriptor, CanExecuteResult> availability)
    {
        if (favorite is null) throw new ArgumentNullException(nameof(favorite));
        var command = registry.FirstOrDefault(value => string.Equals(value.Id, favorite.CommandId, StringComparison.Ordinal));
        if (command is null)
            return Unavailable(favorite, FavoriteResolutionStatus.MissingCommand, null, RefusalCodes.CommandUnavailable,
                "The referenced command is not installed.", "Install or enable a compatible command, or remove this favorite.");
        if (command.ContractVersion != favorite.ContractVersion)
            return Unavailable(favorite, FavoriteResolutionStatus.IncompatibleVersion, command, RefusalCodes.ContractMismatch,
                "The favorite targets a different command contract version.", "Recreate the favorite for the installed command version.");
        if (favorite.Arguments.Count > 0 && !command.HasFixedParameters)
            return Unavailable(favorite, FavoriteResolutionStatus.InvalidArguments, command, RefusalCodes.ContractMismatch,
                "The command does not accept fixed favorite arguments.", "Remove the arguments or recreate the favorite.");
        var result = availability(command);
        return new FavoriteResolution(favorite, result.Allowed ? FavoriteResolutionStatus.Available : FavoriteResolutionStatus.Unavailable,
            command, result, result.Remediation);
    }

    public static CommandResult Invoke(FavoriteDefinition favorite, IReadOnlyList<CommandDescriptor> registry,
        Func<CommandDescriptor, CanExecuteResult> availability,
        Func<string, IReadOnlyDictionary<string, string>, InvocationSource, CommandResult> route)
    {
        if (route is null) throw new ArgumentNullException(nameof(route));
        var resolution = Resolve(favorite, registry, availability);
        if (!resolution.CanInvoke)
            return CommandResult.Refused("favorite.invoke", $"{resolution.Availability.Message} {resolution.Remediation}".Trim(), resolution.Availability.RefusalCode);
        return route(favorite.CommandId, favorite.Arguments, InvocationSource.Favorite);
    }

    private static FavoriteResolution Unavailable(FavoriteDefinition favorite, FavoriteResolutionStatus status,
        CommandDescriptor? command, string code, string message, string remediation) =>
        new FavoriteResolution(favorite, status, command, CanExecuteResult.Refuse(code, message, remediation), remediation);
}
