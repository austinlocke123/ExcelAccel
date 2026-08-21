using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Application.Styles;
using Xunit;

namespace ExcelAccel.Core.Tests;

/// <summary>
/// The validator <c>docs/commands/RIBBON_LAYOUT.md</c> describes. It did not
/// exist: <c>RibbonRoutes</c> is hand-maintained and <see cref="RibbonRoutes.For"/>
/// falls back to the Command Search route on an unknown id, so a descriptor could
/// advertise a keyboard path that does nothing and every existing test would still
/// pass.
/// </summary>
/// <remarks>
/// The ribbon XML is read from source rather than from <c>ExcelAccelRibbon</c>,
/// because the test project deliberately does not reference the net48 host and
/// <c>ArchitectureBoundaryTests</c> asserts that exact project graph. Reading a
/// source file to check a repository invariant is the pattern that test already
/// uses for the .csproj dependency graph.
/// </remarks>
public sealed class RibbonRouteTests
{
    [Fact]
    public void EveryRibbonButtonCarriesAnActionAndAKeyTip()
    {
        foreach (var button in Buttons(RibbonXml()))
        {
            var id = button.Element.Attribute("id")?.Value ?? "(no id)";
            Assert.False(
                string.IsNullOrWhiteSpace(button.Element.Attribute("onAction")?.Value),
                $"Ribbon control '{id}' has no onAction, so pressing it does nothing.");
            Assert.False(
                string.IsNullOrWhiteSpace(button.Element.Attribute("keytip")?.Value),
                $"Ribbon control '{id}' has no keytip, so it is unreachable from the keyboard.");
        }
    }

    /// <summary>
    /// Excel resolves a KeyTip the moment it is unambiguous, so a single-letter
    /// KeyTip makes every longer KeyTip starting with that letter unreachable.
    /// </summary>
    [Fact]
    public void NoKeyTipIsAPrefixOfAnotherWithinTheSameMenu()
    {
        foreach (var scope in KeyTipScopes(RibbonXml()))
        {
            var duplicates = scope.Value
                .GroupBy(entry => entry.KeyTip, StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() > 1)
                .Select(group => group.Key)
                .ToArray();
            Assert.True(
                duplicates.Length == 0,
                $"Duplicate KeyTip(s) {string.Join(", ", duplicates)} inside '{scope.Key}'.");

            foreach (var candidate in scope.Value)
            {
                var shadowed = scope.Value
                    .Where(other =>
                        !ReferenceEquals(other, candidate)
                        && other.KeyTip.Length > candidate.KeyTip.Length
                        && other.KeyTip.StartsWith(candidate.KeyTip, StringComparison.OrdinalIgnoreCase))
                    .Select(other => other.KeyTip)
                    .ToArray();
                Assert.True(
                    shadowed.Length == 0,
                    $"KeyTip '{candidate.KeyTip}' in '{scope.Key}' is a prefix of {string.Join(", ", shadowed)}, "
                    + "so Excel resolves the shorter one first and the longer ones are unreachable.");
            }
        }
    }

    [Fact]
    public void EveryTaggedButtonNamesARegisteredCommand()
    {
        var registered = BuiltInCommandRegistry.All.Select(command => command.Id).ToHashSet(StringComparer.Ordinal);

        foreach (var button in Buttons(RibbonXml()))
        {
            var tag = button.Element.Attribute("tag")?.Value;
            if (string.IsNullOrWhiteSpace(tag) || !tag.Contains('.'))
            {
                // Untagged controls use a fixed-id callback, and style buttons
                // carry a style id rather than a command id.
                continue;
            }

            Assert.True(
                registered.Contains(tag),
                $"Ribbon control '{button.Element.Attribute("id")?.Value}' is tagged '{tag}', which is not a registered command.");
        }
    }

    /// <summary>
    /// The point of the whole file: a descriptor's advertised route must be the
    /// path a user actually presses, because Command Search and the shortcut
    /// cheat sheet print it verbatim.
    /// </summary>
    [Fact]
    public void EveryRibbonHostedCommandAdvertisesTheRouteItsButtonActuallyHas()
    {
        var mismatches = new List<string>();
        foreach (var button in Buttons(RibbonXml()))
        {
            var tag = button.Element.Attribute("tag")?.Value;
            if (string.IsNullOrWhiteSpace(tag) || !RibbonRoutes.Has(tag))
            {
                continue;
            }

            var declared = RibbonRoutes.For(tag);
            if (!string.Equals(declared, button.Route, StringComparison.Ordinal))
            {
                mismatches.Add($"{tag}: RibbonRoutes says '{declared}', ribbon path is '{button.Route}'");
            }
        }

        Assert.True(mismatches.Count == 0, string.Join(Environment.NewLine, mismatches));
    }

    /// <summary>
    /// A route for something that no longer exists is dead weight, and worse, it
    /// makes <see cref="RibbonRoutes.For"/> look like it resolved when a caller
    /// typo'd an id. Built-in style buttons legitimately hold routes under a
    /// style id rather than a command id, so they are the one exemption.
    /// </summary>
    [Fact]
    public void EveryRouteInTheTableBelongsToARegisteredCommandOrBuiltInStyle()
    {
        var known = BuiltInCommandRegistry.All.Select(command => command.Id)
            .Concat(BuiltInStyleCatalog.All.Select(style => style.StyleId))
            .ToHashSet(StringComparer.Ordinal);

        var orphans = RibbonRoutes.All.Keys.Where(id => !known.Contains(id)).ToArray();

        Assert.True(
            orphans.Length == 0,
            $"RibbonRoutes carries route(s) for nothing that exists: {string.Join(", ", orphans)}.");
    }

    /// <summary>
    /// The drift this package exists to catch. A descriptor may compose its route
    /// however it likes, and four Model Check descriptors built theirs by
    /// concatenating a different command's route, advertising a path that does
    /// nothing in Command Search and the cheat sheet.
    /// </summary>
    [Fact]
    public void EveryDescriptorRouteMatchesTheRouteTable()
    {
        var mismatches = BuiltInCommandRegistry.All
            .Where(command => RibbonRoutes.Has(command.Id))
            .Where(command => !string.Equals(command.KeyboardRoute, RibbonRoutes.For(command.Id), StringComparison.Ordinal)
                || !string.Equals(command.ShortcutLabel, RibbonRoutes.For(command.Id), StringComparison.Ordinal))
            .Select(command =>
                $"{command.Id}: descriptor says '{command.KeyboardRoute}' (label '{command.ShortcutLabel}'), table says '{RibbonRoutes.For(command.Id)}'")
            .ToArray();

        Assert.True(mismatches.Length == 0, string.Join(Environment.NewLine, mismatches));
    }

    private static XElement RibbonXml()
    {
        var source = File.ReadAllText(Path.Combine(
            RepositoryRoot(), "src", "ExcelAccel.ExcelAddIn", "ExcelAccelRibbon.cs"));
        var start = source.IndexOf("<customUI", StringComparison.Ordinal);
        var end = source.IndexOf("</customUI>", StringComparison.Ordinal);
        Assert.True(start >= 0 && end > start, "The ribbon XML could not be located in ExcelAccelRibbon.cs.");

        var xml = source.Substring(start, end - start + "</customUI>".Length)
            .Replace("\"\"", "\"", StringComparison.Ordinal);
        return XElement.Parse(xml);
    }

    private static string RepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "ExcelAccel.sln")))
        {
            directory = directory.Parent;
        }

        Assert.NotNull(directory);
        return directory!.FullName;
    }

    private static IEnumerable<RibbonButton> Buttons(XElement root)
    {
        foreach (var tab in root.Descendants().Where(element => element.Name.LocalName == "tab"))
        {
            // A tab KeyTip is pressed as separate letters; menu KeyTips are
            // printed as one token. This mirrors how RibbonRoutes writes them.
            var prefix = "Alt, " + string.Join(", ", (tab.Attribute("keytip")?.Value ?? string.Empty).ToCharArray());
            foreach (var button in Walk(tab, prefix))
            {
                yield return button;
            }
        }
    }

    private static IEnumerable<RibbonButton> Walk(XElement parent, string prefix)
    {
        foreach (var child in parent.Elements())
        {
            switch (child.Name.LocalName)
            {
                case "group":
                    foreach (var button in Walk(child, prefix))
                    {
                        yield return button;
                    }

                    break;
                case "menu":
                    var menuPrefix = prefix + ", " + child.Attribute("keytip")?.Value;
                    foreach (var button in Walk(child, menuPrefix))
                    {
                        yield return button;
                    }

                    break;
                case "button":
                case "toggleButton":
                    yield return new RibbonButton(child, prefix + ", " + child.Attribute("keytip")?.Value);
                    break;
                default:
                    foreach (var button in Walk(child, prefix))
                    {
                        yield return button;
                    }

                    break;
            }
        }
    }

    /// <summary>
    /// KeyTips only have to be unique within the scope Excel resolves them in:
    /// everything directly on the tab, and everything inside one menu.
    /// </summary>
    private static IReadOnlyDictionary<string, List<KeyTipEntry>> KeyTipScopes(XElement root)
    {
        var scopes = new Dictionary<string, List<KeyTipEntry>>(StringComparer.Ordinal);

        void Collect(XElement parent, string scopeName)
        {
            foreach (var child in parent.Elements())
            {
                var name = child.Name.LocalName;
                var keytip = child.Attribute("keytip")?.Value;
                if (name == "group")
                {
                    Collect(child, scopeName);
                }
                else if (name == "menu")
                {
                    Add(scopeName, keytip, child);
                    Collect(child, child.Attribute("id")?.Value ?? "menu");
                }
                else if (name is "button" or "toggleButton")
                {
                    Add(scopeName, keytip, child);
                }
                else
                {
                    Collect(child, scopeName);
                }
            }
        }

        void Add(string scopeName, string? keytip, XElement element)
        {
            if (string.IsNullOrWhiteSpace(keytip))
            {
                return;
            }

            if (!scopes.TryGetValue(scopeName, out var entries))
            {
                entries = new List<KeyTipEntry>();
                scopes.Add(scopeName, entries);
            }

            entries.Add(new KeyTipEntry(keytip!, element.Attribute("id")?.Value ?? "(no id)"));
        }

        foreach (var tab in root.Descendants().Where(element => element.Name.LocalName == "tab"))
        {
            Collect(tab, tab.Attribute("id")?.Value ?? "tab");
        }

        return scopes;
    }

    private sealed record RibbonButton(XElement Element, string Route);

    private sealed record KeyTipEntry(string KeyTip, string ControlId);
}
