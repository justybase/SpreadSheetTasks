using System;
using System.Collections.Generic;

namespace SpreadSheetTasks
{
    /// <summary>
    /// A resolved A1 cell or rectangular range with an attached style.
    /// Bounds are absolute worksheet coordinates (zero-based).
    /// </summary>
    internal readonly struct StyleRange
    {
        public readonly int FirstRow;
        public readonly int FirstColumn;
        public readonly int LastRow;
        public readonly int LastColumn;
        public readonly CellStyle Style;

        public StyleRange(int firstRow, int firstColumn, int lastRow, int lastColumn, CellStyle style)
        {
            FirstRow = firstRow;
            FirstColumn = firstColumn;
            LastRow = lastRow;
            LastColumn = lastColumn;
            Style = style;
        }
    }

    /// <summary>
    /// Parses A1 cell and range references such as <c>A1</c>, <c>$B$2</c> or <c>A1:C10</c>.
    /// </summary>
    internal static class A1Reference
    {
        private const int MinRow = 1;
        private const long MaxRow = 1_048_576;
        private const long MaxColumn = 16_384;

        /// <summary>
        /// Converts a mapping of A1 references to styles into a list of resolved ranges.
        /// Enumeration order is preserved; later ranges override earlier ones.
        /// </summary>
        public static List<StyleRange>? ParseStyleRanges(IReadOnlyDictionary<string, CellStyle>? styles)
        {
            if (styles is null || styles.Count == 0)
            {
                return null;
            }
            var result = new List<StyleRange>(styles.Count);
            foreach (var kvp in styles)
            {
                if (kvp.Value is null)
                {
                    throw new ArgumentNullException(nameof(styles), $"Style for '{kvp.Key}' must not be null.");
                }
                result.Add(ParseStyleRange(kvp.Key, kvp.Value));
            }
            return result;
        }

        /// <summary>Parses one A1 reference with an attached style.</summary>
        public static StyleRange ParseStyleRange(string reference, CellStyle style)
        {
            ArgumentNullException.ThrowIfNull(style);
            if (string.IsNullOrWhiteSpace(reference))
            {
                throw new ArgumentException("Style range keys must be A1 references such as 'A1:D1'.", nameof(reference));
            }

            int colon = reference.IndexOf(':');
            if (colon >= 0 && reference.IndexOf(':', colon + 1) >= 0)
            {
                throw new ArgumentException($"Invalid A1 range '{reference}'.", nameof(reference));
            }

            string startPart = colon < 0 ? reference : reference[..colon];
            string endPart = colon < 0 ? reference : reference[(colon + 1)..];

            (int firstRow, int firstColumn) = ParseCell(startPart, reference);
            (int lastRow, int lastColumn) = colon < 0 ? (firstRow, firstColumn) : ParseCell(endPart, reference);

            if (lastRow < firstRow || lastColumn < firstColumn)
            {
                throw new ArgumentException($"Range end precedes its start in '{reference}'.", nameof(reference));
            }
            return new StyleRange(firstRow, firstColumn, lastRow, lastColumn, style);
        }

        private static (int row, int column) ParseCell(string cell, string reference)
        {
            int i = 0;
            bool invalid = false;

            if (i < cell.Length && cell[i] == '$')
            {
                i++;
            }

            int letterStart = i;
            long column = 0;
            while (i < cell.Length && char.IsAsciiLetter(cell[i]))
            {
                if (column <= MaxColumn)
                {
                    column = column * 26 + (char.ToUpperInvariant(cell[i]) - 'A' + 1);
                }
                if (column > MaxColumn)
                {
                    invalid = true;
                }
                i++;
            }
            if (i == letterStart)
            {
                invalid = true;
            }

            if (i < cell.Length && cell[i] == '$')
            {
                i++;
            }

            int digitStart = i;
            long row = 0;
            while (i < cell.Length && char.IsAsciiDigit(cell[i]))
            {
                row = row * 10 + (cell[i] - '0');
                i++;
            }
            if (i == digitStart || i != cell.Length || row < MinRow || row > MaxRow)
            {
                invalid = true;
            }

            if (invalid)
            {
                throw new ArgumentException(
                    $"Invalid A1 range '{reference}'; expected a cell such as 'A1' or a rectangle such as 'A1:C10'.",
                    nameof(reference));
            }
            return ((int)row - 1, (int)column - 1);
        }
    }
}
