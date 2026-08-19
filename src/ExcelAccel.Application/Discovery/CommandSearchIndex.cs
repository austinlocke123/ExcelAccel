using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelAccel.Application.Commands;

namespace ExcelAccel.Application.Discovery;

public sealed class CommandSearchResult
{
    public CommandSearchResult(CommandDescriptor command, int score, CanExecuteResult availability)
    {
        Command = command ?? throw new ArgumentNullException(nameof(command));
        Score = score;
        Availability = availability ?? throw new ArgumentNullException(nameof(availability));
    }

    public CommandDescriptor Command { get; }
    public int Score { get; }
    public CanExecuteResult Availability { get; }
}

public sealed class CommandSearchIndex
{
    public const int MaximumRegistrySize = 2048;
    public const int MaximumQueryLength = 128;
    public const int MaximumResults = 100;
    private readonly IReadOnlyList<IndexedCommand> _commands;

    public CommandSearchIndex(IEnumerable<CommandDescriptor> commands)
    {
        var descriptors = (commands ?? throw new ArgumentNullException(nameof(commands))).ToArray();
        if (descriptors.Length > MaximumRegistrySize)
            throw new ArgumentException($"The registry exceeds {MaximumRegistrySize} commands.", nameof(commands));
        if (descriptors.Select(value => value.Id).Distinct(StringComparer.Ordinal).Count() != descriptors.Length)
            throw new ArgumentException("Command IDs must be unique.", nameof(commands));

        _commands = descriptors.Select(value => new IndexedCommand(value)).ToArray();
    }

    public IReadOnlyList<CommandSearchResult> Search(
        string? query,
        Func<CommandDescriptor, CanExecuteResult> availability,
        int maximumResults = 50)
    {
        if (availability is null) throw new ArgumentNullException(nameof(availability));
        if (maximumResults < 1 || maximumResults > MaximumResults) throw new ArgumentOutOfRangeException(nameof(maximumResults));
        var normalizedQuery = Normalize(query ?? string.Empty);
        if (normalizedQuery.Length > MaximumQueryLength)
            throw new ArgumentException($"Queries may contain at most {MaximumQueryLength} normalized characters.", nameof(query));
        var terms = normalizedQuery.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

        return _commands
            .Select(value => new { Indexed = value, Score = value.Score(terms) })
            .Where(value => value.Score >= 0)
            .OrderByDescending(value => value.Score)
            .ThenBy(value => value.Indexed.Command.DisplayName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(value => value.Indexed.Command.Id, StringComparer.Ordinal)
            .Take(maximumResults)
            .Select(value => new CommandSearchResult(value.Indexed.Command, value.Score, availability(value.Indexed.Command)))
            .ToArray();
    }

    private static string Normalize(string value)
    {
        var normalized = value.Normalize(NormalizationForm.FormKC).Trim().ToLowerInvariant();
        var builder = new StringBuilder(normalized.Length);
        var priorWhitespace = false;
        foreach (var character in normalized)
        {
            var whitespace = char.IsWhiteSpace(character);
            if (whitespace)
            {
                if (!priorWhitespace) builder.Append(' ');
            }
            else builder.Append(character);
            priorWhitespace = whitespace;
        }
        return builder.ToString();
    }

    private sealed class IndexedCommand
    {
        private readonly IReadOnlyList<SearchField> _fields;

        public IndexedCommand(CommandDescriptor command)
        {
            Command = command;
            var fields = new List<SearchField>
            {
                new SearchField(command.DisplayName, 1000),
                new SearchField(command.ShortcutLabel, 800),
                new SearchField(command.Category, 650),
                new SearchField(command.Id, 550),
                new SearchField(command.Description, 450),
            };
            fields.AddRange(command.Aliases.Select(value => new SearchField(value, 900)));
            _fields = fields;
        }

        public CommandDescriptor Command { get; }

        public int Score(IReadOnlyList<string> terms)
        {
            if (terms.Count == 0) return 0;
            var total = 0;
            foreach (var term in terms)
            {
                var best = -1;
                foreach (var field in _fields) best = Math.Max(best, field.Score(term));
                if (best < 0) return -1;
                total += best;
            }
            return total;
        }
    }

    private sealed class SearchField
    {
        private readonly string _value;
        private readonly int _weight;
        public SearchField(string value, int weight) { _value = Normalize(value); _weight = weight; }

        public int Score(string term)
        {
            if (string.Equals(_value, term, StringComparison.Ordinal)) return _weight + 100;
            if (_value.StartsWith(term, StringComparison.Ordinal)) return _weight + 75;
            if (_value.Split(' ').Any(token => token.StartsWith(term, StringComparison.Ordinal))) return _weight + 60;
            return _value.IndexOf(term, StringComparison.Ordinal) >= 0 ? _weight + 40 : -1;
        }
    }
}
