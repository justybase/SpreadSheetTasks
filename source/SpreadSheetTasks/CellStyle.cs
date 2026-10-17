using System;
using System.Collections.Generic;

namespace SpreadSheetTasks
{
    /// <summary>
    /// A reusable Excel cell style: solid fill color, font color, bold/italic/underline,
    /// number format and (XLSX only for now) horizontal/vertical alignment and wrapping.
    /// <para>
    /// Every property is optional. <c>null</c> means "leave this property unchanged",
    /// which makes styles composable: a range style only overrides the properties it sets,
    /// both over a header style and over another overlapping range style.
    /// </para>
    /// </summary>
    /// <remarks>
    /// Colors are six-digit RGB hex strings with or without the leading <c>#</c>,
    /// for example <c>"FFF2CC"</c> or <c>"#FFF2CC"</c>.
    /// <see cref="Horizontal"/>, <see cref="Vertical"/> and <see cref="Wrap"/> are currently
    /// only supported by <see cref="XlsxWriter"/>; <see cref="XlsbWriter"/> throws
    /// <see cref="NotSupportedException"/> when they are set.
    /// </remarks>
    public sealed record CellStyle
    {
        private static readonly Dictionary<string, string> _horizontalAlignments = new(StringComparer.OrdinalIgnoreCase)
        {
            ["general"] = "general",
            ["left"] = "left",
            ["center"] = "center",
            ["right"] = "right",
            ["fill"] = "fill",
            ["justify"] = "justify",
            ["centerContinuous"] = "centerContinuous",
            ["distributed"] = "distributed",
        };

        private static readonly Dictionary<string, string> _verticalAlignments = new(StringComparer.OrdinalIgnoreCase)
        {
            ["top"] = "top",
            ["center"] = "center",
            ["bottom"] = "bottom",
            ["justify"] = "justify",
            ["distributed"] = "distributed",
        };

        private string? _fill;
        private string? _fontColor;
        private string? _horizontal;
        private string? _vertical;

        /// <summary>Creates an empty style that leaves every property unchanged.</summary>
        public CellStyle()
        {
        }

        /// <summary>Creates a style with all properties set at once.</summary>
        public CellStyle(
            string? fill = null,
            string? fontColor = null,
            bool? bold = null,
            bool? italic = null,
            bool? underline = null,
            string? numberFormat = null,
            string? horizontal = null,
            string? vertical = null,
            bool? wrap = null)
        {
            Fill = fill;
            FontColor = fontColor;
            Bold = bold;
            Italic = italic;
            Underline = underline;
            NumberFormat = numberFormat;
            Horizontal = horizontal;
            Vertical = vertical;
            Wrap = wrap;
        }

        /// <summary>
        /// Solid fill color as a six-digit RGB hex string, e.g. <c>"FFF2CC"</c> or <c>"#FFF2CC"</c>.
        /// </summary>
        public string? Fill
        {
            get => _fill;
            init => _fill = NormalizeColor(value, nameof(Fill));
        }

        /// <summary>
        /// Font color as a six-digit RGB hex string, e.g. <c>"FFFFFF"</c> or <c>"#FFFFFF"</c>.
        /// </summary>
        public string? FontColor
        {
            get => _fontColor;
            init => _fontColor = NormalizeColor(value, nameof(FontColor));
        }

        /// <summary>Bold font. <c>null</c> leaves the current value unchanged.</summary>
        public bool? Bold { get; init; }

        /// <summary>Italic font. <c>null</c> leaves the current value unchanged.</summary>
        public bool? Italic { get; init; }

        /// <summary>Single underline. <c>null</c> leaves the current value unchanged.</summary>
        public bool? Underline { get; init; }

        /// <summary>
        /// Excel number format string applied by the style, e.g. <see cref="F.THOUSANDS_SEP"/>.
        /// A <see cref="FormattedCell"/> format on the cell itself takes precedence.
        /// </summary>
        public string? NumberFormat { get; init; }

        /// <summary>
        /// Horizontal alignment: <c>general</c>, <c>left</c>, <c>center</c>, <c>right</c>,
        /// <c>fill</c>, <c>justify</c>, <c>centerContinuous</c> or <c>distributed</c>.
        /// XLSX only; <see cref="XlsbWriter"/> throws when set.
        /// </summary>
        public string? Horizontal
        {
            get => _horizontal;
            init => _horizontal = NormalizeAlignment(value, _horizontalAlignments, nameof(Horizontal));
        }

        /// <summary>
        /// Vertical alignment: <c>top</c>, <c>center</c>, <c>bottom</c>, <c>justify</c>
        /// or <c>distributed</c>. XLSX only; <see cref="XlsbWriter"/> throws when set.
        /// </summary>
        public string? Vertical
        {
            get => _vertical;
            init => _vertical = NormalizeAlignment(value, _verticalAlignments, nameof(Vertical));
        }

        /// <summary>Wrap text. XLSX only; <see cref="XlsbWriter"/> throws when set.</summary>
        public bool? Wrap { get; init; }

        /// <summary>
        /// Returns a style where every non-null property of <paramref name="overlay"/>
        /// replaces the same property of <paramref name="baseStyle"/>.
        /// </summary>
        /// <param name="baseStyle">The base style; may be null.</param>
        /// <param name="overlay">The style whose set properties win.</param>
        public static CellStyle Merge(CellStyle? baseStyle, CellStyle overlay)
        {
            ArgumentNullException.ThrowIfNull(overlay);
            if (baseStyle is null)
            {
                return overlay;
            }
            return new CellStyle
            {
                Fill = overlay.Fill ?? baseStyle.Fill,
                FontColor = overlay.FontColor ?? baseStyle.FontColor,
                Bold = overlay.Bold ?? baseStyle.Bold,
                Italic = overlay.Italic ?? baseStyle.Italic,
                Underline = overlay.Underline ?? baseStyle.Underline,
                NumberFormat = overlay.NumberFormat ?? baseStyle.NumberFormat,
                Horizontal = overlay.Horizontal ?? baseStyle.Horizontal,
                Vertical = overlay.Vertical ?? baseStyle.Vertical,
                Wrap = overlay.Wrap ?? baseStyle.Wrap,
            };
        }

        /// <summary>True when the style sets no property at all.</summary>
        internal bool IsEmpty =>
            Fill is null && FontColor is null && Bold is null && Italic is null
            && Underline is null && NumberFormat is null
            && Horizontal is null && Vertical is null && Wrap is null;

        internal bool HasAlignment => Horizontal is not null || Vertical is not null || Wrap is not null;

        internal FontKey ToFontKey() => new(FontColor, Bold, Italic, Underline);

        private static string? NormalizeColor(string? value, string propertyName)
        {
            if (value is null)
            {
                return null;
            }
            if (value.Length == 7 && value[0] == '#')
            {
                value = value[1..];
            }
            if (value.Length != 6)
            {
                throw new ArgumentException(
                    $"Invalid RGB color '{value}'; expected a six-digit hex value such as '#008000'.", propertyName);
            }
            Span<char> normalized = stackalloc char[6];
            for (int i = 0; i < normalized.Length; i++)
            {
                char c = char.ToUpperInvariant(value[i]);
                if (!Uri.IsHexDigit(c))
                {
                    throw new ArgumentException(
                        $"Invalid RGB color '{value}'; expected a six-digit hex value such as '#008000'.", propertyName);
                }
                normalized[i] = c;
            }
            return new string(normalized);
        }

        private static string? NormalizeAlignment(string? value, Dictionary<string, string> allowed, string propertyName)
        {
            if (value is null)
            {
                return null;
            }
            if (!allowed.TryGetValue(value, out string? normalized))
            {
                throw new ArgumentException(
                    $"Invalid {propertyName.ToLowerInvariant()} alignment '{value}'. Allowed values: {string.Join(", ", allowed.Values)}.",
                    propertyName);
            }
            return normalized;
        }
    }

    /// <summary>Identity of a font record: only the properties that change the font definition.</summary>
    internal readonly record struct FontKey(string? Color, bool? Bold, bool? Italic, bool? Underline);
}
