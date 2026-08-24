using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace ExcelAccel.Application.AutoColor;

/// <summary>
/// The font colours of a bounded set of cells, carried as one undo value.
/// </summary>
/// <remarks>
/// AutoColor recolours far more cells than <see cref="Undo.PropertyBatchReceipt"/>
/// can hold as individual changes, since that caps at 32. Rather than raise the
/// cap, this follows the shape the product already uses twice: one coarse
/// property carrying a whole block, like <c>cell_contents_v1</c> for formula
/// blocks and <c>cell_format_block_v1</c> for formats-only paste. A recolour of
/// any size is therefore one change on one receipt, and a single undo restores
/// all of it or none.
///
/// The serialized form groups addresses under their colour, because a real model
/// holds a handful of distinct font colours across thousands of cells. That keeps
/// a large block far inside the store's per-value character limit while staying
/// exact: no colour is approximated and no cell is dropped.
/// </remarks>
public static class FontColorBlock
{
    public const string ReceiptPropertyId = "cell_font_color_block_v1";

    /// <summary>
    /// The execution ceiling for a recolour, matching the existing formatting
    /// ceiling. It exists because the undo value has to fit: at worst every cell
    /// carries a distinct colour, which is about 16 characters per cell, so
    /// 50,000 cells stays comfortably inside the store's 1,000,000-character
    /// limit for a single value.
    /// </summary>
    public const int MaximumCells = 50_000;

    private const char ColorSeparator = ';';
    private const char ColorAssignment = '=';
    private const char AddressSeparator = ',';

    /// <summary>
    /// Writes the block deterministically: colours in ordinal order, addresses in
    /// ordinal order within each colour. Two captures of the same state produce
    /// byte-identical values, which is what lets undo compare them exactly.
    /// </summary>
    public static string Serialize(IEnumerable<KeyValuePair<string, string>> cells)
    {
        if (cells is null) throw new ArgumentNullException(nameof(cells));

        var normalized = new SortedDictionary<string, SortedSet<string>>(StringComparer.Ordinal);
        // Addresses are tracked across every colour, not within one. The same
        // cell listed under two colours would otherwise serialize twice and come
        // back as one cell with two conflicting colours, and undo would write
        // whichever it read last.
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var count = 0;
        foreach (var cell in cells)
        {
            var address = RequireAddress(cell.Key);
            var color = RequireColor(cell.Value);
            if (!seen.Add(address))
            {
                throw new ArgumentException($"Cell '{address}' appears more than once.", nameof(cells));
            }

            if (!normalized.TryGetValue(color, out var addresses))
            {
                addresses = new SortedSet<string>(StringComparer.Ordinal);
                normalized.Add(color, addresses);
            }

            addresses.Add(address);
            count++;
            if (count > MaximumCells)
            {
                throw new ArgumentException(
                    $"A font-colour block is limited to {MaximumCells:N0} cells.", nameof(cells));
            }
        }

        var builder = new StringBuilder();
        foreach (var entry in normalized)
        {
            if (builder.Length > 0) builder.Append(ColorSeparator);
            builder.Append(entry.Key).Append(ColorAssignment).Append(string.Join(AddressSeparator.ToString(), entry.Value));
        }

        return builder.ToString();
    }

    /// <summary>
    /// Reads a block back. Every address is returned with its colour, in the same
    /// deterministic order <see cref="Serialize"/> wrote.
    /// </summary>
    public static IReadOnlyList<KeyValuePair<string, string>> Deserialize(string? value)
    {
        var text = value ?? string.Empty;
        var result = new List<KeyValuePair<string, string>>();
        if (text.Trim().Length == 0)
        {
            return result;
        }

        foreach (var group in text.Split(ColorSeparator))
        {
            if (group.Length == 0) continue;
            var split = group.IndexOf(ColorAssignment);
            if (split <= 0 || split + 1 >= group.Length)
            {
                throw new FormatException("A font-colour block group requires 'colour=addresses'.");
            }

            var color = RequireColor(group.Substring(0, split));
            foreach (var address in group.Substring(split + 1).Split(AddressSeparator))
            {
                if (address.Length == 0) continue;
                result.Add(new KeyValuePair<string, string>(RequireAddress(address), color));
            }
        }

        return result;
    }

    /// <summary>
    /// A conservative worst-case size for a block of this many cells, used to
    /// keep the execution ceiling and the store's value limit reasoned about
    /// together rather than guessed.
    /// </summary>
    public static long WorstCaseCharacters(int cellCount) =>
        cellCount <= 0 ? 0 : (long)cellCount * 16;

    private static string RequireAddress(string? address)
    {
        var value = (address ?? string.Empty).Trim().ToUpperInvariant();
        if (value.Length == 0)
        {
            throw new ArgumentException("A cell address is required.", nameof(address));
        }

        if (value.IndexOf(ColorSeparator) >= 0
            || value.IndexOf(ColorAssignment) >= 0
            || value.IndexOf(AddressSeparator) >= 0)
        {
            throw new ArgumentException($"A cell address cannot contain a block separator: '{value}'.", nameof(address));
        }

        return value;
    }

    private static string RequireColor(string? color)
    {
        var value = (color ?? string.Empty).Trim().ToUpperInvariant();
        if (value.Length != 7 || value[0] != '#' || !value.Skip(1).All(IsHex))
        {
            throw new ArgumentException($"A #RRGGBB colour is required, not '{color}'.", nameof(color));
        }

        return value;
    }

    private static bool IsHex(char value) =>
        (value >= '0' && value <= '9') || (value >= 'A' && value <= 'F');
}
