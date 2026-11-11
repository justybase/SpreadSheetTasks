using System;
using System.Buffers;
using System.Text;

namespace SpreadSheetTasks;

/// <summary>
/// XML character-data decoding shared by the fast XLSX scanners.
/// <para>
/// Element text in SpreadsheetML is a small subset of XML: the five predefined entities,
/// decimal/hex character references, and XML end-of-line normalization (CRLF and lone CR
/// become LF). Anything else is reported as unsupported so the caller can hand the part to
/// <see cref="System.Xml.XmlReader"/>, which applies the full XML rules and raises the
/// exception a damaged file deserves.
/// </para>
/// </summary>
internal static class FastXmlText
{
    /// <summary>Bytes that need more than a straight UTF-8 decode.</summary>
    private static readonly SearchValues<byte> XmlTextSpecialBytes = SearchValues.Create("<&\r"u8);

    /// <summary>
    /// Decodes text that a caller already knows has no markup besides character data.
    /// Returns false when the span contains a raw '&lt;', which cannot occur in XML text and
    /// means the surrounding element boundaries were guessed wrongly.
    /// </summary>
    internal static bool TryDecodeText(ReadOnlySpan<byte> value, out string text)
    {
        if (value.IndexOfAny(XmlTextSpecialBytes) < 0)
        {
            text = Encoding.UTF8.GetString(value);
            return true;
        }

        if (value.IndexOf((byte)'<') >= 0)
        {
            text = string.Empty;
            return false;
        }
        return TryDecode(value, out text);
    }

    /// <summary>
    /// Decodes element text. Returns false for an unknown/malformed reference or a character
    /// reference XML does not allow (NUL, other control characters, surrogates, U+FFFE/U+FFFF).
    /// </summary>
    internal static bool TryDecode(ReadOnlySpan<byte> value, out string text)
    {
        text = string.Empty;

        // The overwhelmingly common case: no references and no CR to normalize.
        if (value.IndexOf((byte)'&') < 0 && value.IndexOf((byte)'\r') < 0)
        {
            text = Encoding.UTF8.GetString(value);
            return true;
        }

        var builder = new StringBuilder(value.Length);
        int i = 0;
        while (i < value.Length)
        {
            int relative = value[i..].IndexOfAny("&\r"u8);
            if (relative < 0)
            {
                builder.Append(Encoding.UTF8.GetString(value[i..]));
                break;
            }

            if (relative > 0)
            {
                builder.Append(Encoding.UTF8.GetString(value.Slice(i, relative)));
                i += relative;
            }

            if (value[i] == (byte)'\r')
            {
                builder.Append('\n');
                i++;
                if (i < value.Length && value[i] == (byte)'\n')
                {
                    i++;
                }
                continue;
            }

            if (!TryDecodeEntity(value, ref i, builder))
            {
                return false;
            }
        }

        text = builder.ToString();
        return true;
    }

    internal static bool IsAllXmlWhitespace(ReadOnlySpan<byte> value)
    {
        foreach (byte b in value)
        {
            if (b != ' ' && b != '\t' && b != '\r' && b != '\n')
            {
                return false;
            }
        }
        return true;
    }

    internal static bool IsAllXmlWhitespace(string value)
    {
        foreach (char c in value)
        {
            if (c != ' ' && c != '\t' && c != '\r' && c != '\n')
            {
                return false;
            }
        }
        return true;
    }

    private static bool TryDecodeEntity(ReadOnlySpan<byte> value, ref int position, StringBuilder builder)
    {
        // position points at '&'. The longest valid reference is "#x10FFFF;".
        int start = position + 1;
        int limit = Math.Min(value.Length, start + 12);
        int semicolon = value.Slice(start, limit - start).IndexOf((byte)';');
        if (semicolon < 0)
        {
            return false;
        }

        int tokenLength = semicolon;
        ReadOnlySpan<byte> token = value.Slice(start, tokenLength);
        if (token.SequenceEqual("amp"u8))
        {
            builder.Append('&');
        }
        else if (token.SequenceEqual("lt"u8))
        {
            builder.Append('<');
        }
        else if (token.SequenceEqual("gt"u8))
        {
            builder.Append('>');
        }
        else if (token.SequenceEqual("quot"u8))
        {
            builder.Append('"');
        }
        else if (token.SequenceEqual("apos"u8))
        {
            builder.Append('\'');
        }
        else if (token.Length > 1 && token[0] == (byte)'#')
        {
            int radix = 10;
            ReadOnlySpan<byte> digits = token[1..];
            if (digits.Length > 0 && digits[0] == (byte)'x')
            {
                radix = 16;
                digits = digits[1..];
            }
            else if (digits.Length > 0 && digits[0] == (byte)'X')
            {
                return false; // XML only accepts a lower-case x.
            }

            if (!TryParseCodePoint(digits, radix, out int codePoint))
            {
                return false;
            }
            if (codePoint <= 0xFFFF)
            {
                builder.Append((char)codePoint);
            }
            else
            {
                builder.Append(char.ConvertFromUtf32(codePoint));
            }
        }
        else
        {
            return false;
        }

        position = start + tokenLength + 1;
        return true;
    }

    private static bool TryParseCodePoint(ReadOnlySpan<byte> digits, int radix, out int codePoint)
    {
        codePoint = 0;
        if (digits.Length == 0 || digits.Length > 8)
        {
            return false;
        }

        long value = 0;
        foreach (byte digit in digits)
        {
            int d;
            if (digit >= (byte)'0' && digit <= (byte)'9')
            {
                d = digit - '0';
            }
            else if (radix == 16 && digit >= (byte)'a' && digit <= (byte)'f')
            {
                d = digit - 'a' + 10;
            }
            else if (radix == 16 && digit >= (byte)'A' && digit <= (byte)'F')
            {
                d = digit - 'A' + 10;
            }
            else
            {
                return false;
            }

            value = value * radix + d;
            if (value > 0x10FFFF)
            {
                return false;
            }
        }

        if (!IsXmlCodePoint((int)value))
        {
            return false;
        }
        codePoint = (int)value;
        return true;
    }

    private static bool IsXmlCodePoint(int value)
        => value is 0x9 or 0xA or 0xD
            or (>= 0x20 and <= 0xD7FF)
            or (>= 0xE000 and <= 0xFFFD)
            or (>= 0x10000 and <= 0x10FFFF);
}
