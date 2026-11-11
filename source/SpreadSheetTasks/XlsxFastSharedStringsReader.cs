using System;
using System.Buffers.Text;
using System.Text;
using System.Text.Unicode;

namespace SpreadSheetTasks;

public sealed partial class XlsxOrXlsbReadOrEdit
{
    /// <summary>
    /// Shared strings are commonly the largest XML part of a text-heavy workbook. This scanner
    /// reads the plain markup Excel writes (and that the writers in this library write) directly
    /// from UTF-8 bytes. It returns false for anything else so <see cref="FillSharedStrings"/>
    /// can fall back to <see cref="System.Xml.XmlReader"/>, which decides what the file means.
    /// </summary>
    private static bool TryReadSharedStringsFast(ReadOnlySpan<byte> data, out string[]? strings)
    {
        strings = null;
        if (data.IsEmpty || !Utf8.IsValid(data))
        {
            return false;
        }

        int position = 0;
        if (data.Length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF)
        {
            position = 3;
        }

        if (!SstSkipWhitespace(data, ref position))
        {
            return false;
        }

        // XML declaration and comments/PIs before the root element.
        while (position < data.Length && data[position] == '<')
        {
            if (SstStartsWith(data, position, "<?xml"u8))
            {
                int end = data[position..].IndexOf("?>"u8);
                if (end < 0)
                {
                    return false;
                }
                if (!SstDeclarationIsUtf8(data.Slice(position, end + 2)))
                {
                    return false;
                }
                position += end + 2;
            }
            else if (SstStartsWith(data, position, "<!--"u8))
            {
                int end = data[position..].IndexOf("-->"u8);
                if (end < 0)
                {
                    return false;
                }
                position += end + 3;
            }
            else if (SstStartsWith(data, position, "<?"u8))
            {
                int end = data[position..].IndexOf("?>"u8);
                if (end < 0)
                {
                    return false;
                }
                position += end + 2;
            }
            else
            {
                break;
            }

            if (!SstSkipWhitespace(data, ref position))
            {
                return false;
            }
        }

        if (position >= data.Length || data[position] != '<' || !SstTagNameEquals(data, position, "sst"u8))
        {
            return false;
        }
        if (!SstReadStartTag(data, position, out int sstEnd, out bool sstSelfClosing))
        {
            return false;
        }
        if (sstSelfClosing)
        {
            strings = Array.Empty<string>();
            return true;
        }

        int capacity = 1024;
        if (SstTryGetAttribute(data, position, sstEnd, "uniqueCount"u8, out ReadOnlySpan<byte> uniqueCountValue)
            && Utf8Parser.TryParse(uniqueCountValue, out int uniqueCount, out int consumed)
            && consumed == uniqueCountValue.Length
            && uniqueCount > 0)
        {
            capacity = Math.Min(uniqueCount, MaxSharedStringsInitialCapacity);
        }
        else
        {
            // A missing or unusable count is normal in hand-written files.
            capacity = 4;
        }

        bool rootPreservesWhitespace = SstGetXmlSpace(data, position, sstEnd) ?? false;
        string[] result = new string[capacity];
        int count = 0;
        position = sstEnd;

        while (true)
        {
            if (!SstSkipWhitespace(data, ref position))
            {
                return false;
            }
            if (position >= data.Length)
            {
                return false;
            }
            if (data[position] != '<')
            {
                return false;
            }

            if (SstStartsWith(data, position, "<!--"u8))
            {
                int end = data[position..].IndexOf("-->"u8);
                if (end < 0)
                {
                    return false;
                }
                position += end + 3;
                continue;
            }
            if (SstStartsWith(data, position, "<?"u8))
            {
                int end = data[position..].IndexOf("?>"u8);
                if (end < 0)
                {
                    return false;
                }
                position += end + 2;
                continue;
            }
            if (SstStartsWith(data, position, "</sst"u8))
            {
                if (!SstReadStartTag(data, position, out int afterSst, out _))
                {
                    return false;
                }
                position = afterSst;
                if (!SstOnlyTrailingMarkup(data, ref position))
                {
                    return false;
                }
                if (count != result.Length)
                {
                    Array.Resize(ref result, count);
                }
                strings = result;
                return true;
            }

            // Excel's writer emits the canonical <si><t>text</t></si>. A vectorized search for
            // the closing pair avoids the general tag scanner for what is nearly every item.
            if (SstStartsWith(data, position, "<si><t>"u8))
            {
                int textStart = position + 7;
                int close = data[textStart..].IndexOf("</t></si>"u8);
                if (close >= 0)
                {
                    ReadOnlySpan<byte> text = data.Slice(textStart, close);
                    // A raw '<' cannot appear in element text; TryDecodeText rejects it and
                    // sends such markup to the general scanner below.
                    if (FastXmlText.TryDecodeText(text, out string simpleItem))
                    {
                        // The exact <si><t> shape inherits xml:space from the root only.
                        if (!rootPreservesWhitespace && FastXmlText.IsAllXmlWhitespace(simpleItem))
                        {
                            simpleItem = string.Empty;
                        }
                        if (count == result.Length)
                        {
                            Array.Resize(ref result, Math.Max(4, result.Length * 2));
                        }
                        result[count++] = DecodeXstring(simpleItem);
                        position = textStart + close + 9;
                        continue;
                    }
                }
            }

            if (!SstTagNameEquals(data, position, "si"u8)
                || !SstReadStartTag(data, position, out int siEnd, out bool siSelfClosing))
            {
                return false;
            }

            if (siSelfClosing)
            {
                if (count == result.Length)
                {
                    Array.Resize(ref result, Math.Max(4, result.Length * 2));
                }
                result[count++] = string.Empty;
                position = siEnd;
                continue;
            }

            if (!SstFindCloseTag(data, siEnd, "si"u8, out int siCloseStart, out int siCloseEnd))
            {
                return false;
            }

            bool siPreservesWhitespace = SstGetXmlSpace(data, position, siEnd) ?? rootPreservesWhitespace;
            if (!TryReadSharedStringItem(data.Slice(siEnd, siCloseStart - siEnd), siPreservesWhitespace, out string item))
            {
                return false;
            }
            if (count == result.Length)
            {
                Array.Resize(ref result, Math.Max(4, result.Length * 2));
            }
            result[count++] = item;
            position = siCloseEnd;
        }
    }

    private const int MaxSharedStringsInitialCapacity = 1_000_000;

    /// <summary>
    /// Upper bound for the part we are willing to read into memory for the byte scanner.
    /// Larger parts keep the streaming XmlReader path, which does not hold the whole part.
    /// </summary>
    private const int MaxSharedStringsBytes = 256 * 1024 * 1024;

    private static bool TryReadSharedStringItem(ReadOnlySpan<byte> content, bool preservesWhitespace, out string value)
    {
        // The common item is a single <t> run. Keeping its decoded string directly avoids a
        // StringBuilder allocation and copy for every shared string in the workbook.
        var state = new SharedStringTextState();
        if (!AppendSharedStringText(content, preservesWhitespace, ref state))
        {
            value = string.Empty;
            return false;
        }

        value = DecodeXstring(state.GetValue());
        return true;
    }

    private struct SharedStringTextState
    {
        private string? _single;
        private StringBuilder? _builder;

        internal void Append(string text)
        {
            if (_single is null)
            {
                _single = text;
                return;
            }

            _builder ??= new StringBuilder(_single.Length + text.Length).Append(_single);
            _builder.Append(text);
        }

        internal readonly string GetValue()
            => _builder is not null ? _builder.ToString() : _single ?? string.Empty;
    }

    private static bool AppendSharedStringText(ReadOnlySpan<byte> content, bool preservesWhitespace, ref SharedStringTextState state)
    {
        int position = 0;
        while (true)
        {
            SstSkipWhitespace(content, ref position);
            if (position >= content.Length)
            {
                return true;
            }
            if (content[position] != '<')
            {
                return false; // Text outside a <t> element is not a shared string.
            }

            if (SstStartsWith(content, position, "<!--"u8))
            {
                int end = content[position..].IndexOf("-->"u8);
                if (end < 0)
                {
                    return false;
                }
                position += end + 3;
                continue;
            }
            if (SstStartsWith(content, position, "<?"u8))
            {
                int end = content[position..].IndexOf("?>"u8);
                if (end < 0)
                {
                    return false;
                }
                position += end + 2;
                continue;
            }
            if (position + 1 >= content.Length)
            {
                return false;
            }
            if (content[position + 1] == (byte)'!')
            {
                return false; // CDATA, DOCTYPE or another declaration.
            }
            if (content[position + 1] == (byte)'/')
            {
                return false; // The closing </si> was sliced off before this call.
            }

            if (!SstReadStartTag(content, position, out int tagEnd, out bool selfClosing))
            {
                return false;
            }

            if (SstTagNameEquals(content, position, "t"u8))
            {
                if (selfClosing)
                {
                    position = tagEnd;
                    continue;
                }

                if (!SstFindCloseTag(content, tagEnd, "t"u8, out int closeStart, out int closeEnd))
                {
                    return false;
                }

                bool preserve = SstGetXmlSpace(content, position, tagEnd) ?? preservesWhitespace;
                ReadOnlySpan<byte> text = content.Slice(tagEnd, closeStart - tagEnd);
                if (!preserve && FastXmlText.IsAllXmlWhitespace(text))
                {
                    position = closeEnd;
                    continue;
                }
                if (!FastXmlText.TryDecodeText(text, out string decoded))
                {
                    return false;
                }
                if (!preserve && FastXmlText.IsAllXmlWhitespace(decoded))
                {
                    decoded = string.Empty;
                }
                state.Append(decoded);
                position = closeEnd;
                continue;
            }

            if (SstTagNameEquals(content, position, "r"u8))
            {
                if (selfClosing)
                {
                    position = tagEnd;
                    continue;
                }

                if (!SstFindCloseTag(content, tagEnd, "r"u8, out int closeStart, out int closeEnd))
                {
                    return false;
                }

                bool preserve = SstGetXmlSpace(content, position, tagEnd) ?? preservesWhitespace;
                if (!AppendSharedStringText(content.Slice(tagEnd, closeStart - tagEnd), preserve, ref state))
                {
                    return false;
                }
                position = closeEnd;
                continue;
            }

            if (SstTagNameEquals(content, position, "rPh"u8))
            {
                // Phonetic guides are metadata; they are not the cell's displayed text.
                if (selfClosing)
                {
                    position = tagEnd;
                    continue;
                }
                if (!SstFindCloseTag(content, tagEnd, "rPh"u8, out _, out int closeEnd))
                {
                    return false;
                }
                position = closeEnd;
                continue;
            }

            if (SstTagNameEquals(content, position, "rPr"u8) || SstTagNameEquals(content, position, "phoneticPr"u8))
            {
                if (selfClosing)
                {
                    position = tagEnd;
                    continue;
                }

                ReadOnlySpan<byte> name = SstTagNameEquals(content, position, "rPr"u8) ? "rPr"u8 : "phoneticPr"u8;
                if (!SstFindCloseTag(content, tagEnd, name, out _, out int closeEnd))
                {
                    return false;
                }
                position = closeEnd;
                continue;
            }

            return false;
        }
    }

    private static bool SstOnlyTrailingMarkup(ReadOnlySpan<byte> data, ref int position)
    {
        while (true)
        {
            if (!SstSkipWhitespace(data, ref position))
            {
                return true;
            }
            if (SstStartsWith(data, position, "<!--"u8))
            {
                int end = data[position..].IndexOf("-->"u8);
                if (end < 0)
                {
                    return false;
                }
                position += end + 3;
                continue;
            }
            if (SstStartsWith(data, position, "<?"u8))
            {
                int end = data[position..].IndexOf("?>"u8);
                if (end < 0)
                {
                    return false;
                }
                position += end + 2;
                continue;
            }
            return false;
        }
    }

    private static bool SstSkipWhitespace(ReadOnlySpan<byte> data, ref int position)
    {
        while (position < data.Length)
        {
            byte b = data[position];
            if (b != ' ' && b != '\t' && b != '\r' && b != '\n')
            {
                return true;
            }
            position++;
        }
        return false;
    }

    private static bool SstStartsWith(ReadOnlySpan<byte> data, int position, ReadOnlySpan<byte> value)
        => position <= data.Length - value.Length && data.Slice(position, value.Length).SequenceEqual(value);

    private static bool SstTagNameEquals(ReadOnlySpan<byte> data, int tagStart, ReadOnlySpan<byte> name)
    {
        int position = tagStart + 1;
        if (position < data.Length && data[position] == (byte)'/')
        {
            position++;
        }

        if (position > data.Length - name.Length
            || !data.Slice(position, name.Length).SequenceEqual(name))
        {
            return false;
        }

        int after = position + name.Length;
        return after >= data.Length
            || data[after] == (byte)'>'
            || data[after] == (byte)'/'
            || data[after] == (byte)' '
            || data[after] == (byte)'\t'
            || data[after] == (byte)'\r'
            || data[after] == (byte)'\n';
    }

    private static bool SstReadStartTag(ReadOnlySpan<byte> data, int tagStart, out int tagEnd, out bool selfClosing)
    {
        tagEnd = tagStart;
        selfClosing = false;
        byte quote = 0;
        for (int i = tagStart + 1; i < data.Length; i++)
        {
            byte b = data[i];
            if (quote != 0)
            {
                if (b == quote)
                {
                    quote = 0;
                }
            }
            else if (b == (byte)'"' || b == (byte)'\'')
            {
                quote = b;
            }
            else if (b == (byte)'>')
            {
                selfClosing = i > tagStart + 1 && data[i - 1] == (byte)'/';
                tagEnd = i + 1;
                return true;
            }
        }
        return false;
    }

    private static bool SstFindCloseTag(ReadOnlySpan<byte> data, int from, ReadOnlySpan<byte> name, out int closeStart, out int closeEnd)
    {
        closeStart = 0;
        closeEnd = 0;
        int position = from;
        while (true)
        {
            int relative = data[position..].IndexOf((byte)'<');
            if (relative < 0)
            {
                return false;
            }
            position += relative;

            if (position + 2 + name.Length > data.Length)
            {
                return false;
            }
            if (data[position + 1] != (byte)'/' || !data.Slice(position + 2, name.Length).SequenceEqual(name))
            {
                position++;
                continue;
            }

            int afterName = position + 2 + name.Length;
            if (afterName < data.Length
                && data[afterName] != (byte)'>'
                && data[afterName] != (byte)' '
                && data[afterName] != (byte)'\t'
                && data[afterName] != (byte)'\r'
                && data[afterName] != (byte)'\n')
            {
                position++;
                continue;
            }

            if (!SstReadStartTag(data, position, out closeEnd, out _))
            {
                return false;
            }
            closeStart = position;
            return true;
        }
    }

    private static bool SstTryGetAttribute(
        ReadOnlySpan<byte> data,
        int tagStart,
        int tagEnd,
        ReadOnlySpan<byte> wanted,
        out ReadOnlySpan<byte> value)
    {
        int position = tagStart + 1;
        while (position < tagEnd && FastIsNameByte(data[position]))
        {
            position++;
        }

        while (position < tagEnd)
        {
            while (position < tagEnd && IsFastWhitespace(data[position]))
            {
                position++;
            }
            if (position >= tagEnd || data[position] == (byte)'/' || data[position] == (byte)'>')
            {
                break;
            }

            int nameStart = position;
            while (position < tagEnd && FastIsNameByte(data[position]))
            {
                position++;
            }
            int nameEnd = position;
            while (position < tagEnd && IsFastWhitespace(data[position]))
            {
                position++;
            }
            if (position >= tagEnd || data[position] != (byte)'=')
            {
                break;
            }
            position++;
            while (position < tagEnd && IsFastWhitespace(data[position]))
            {
                position++;
            }
            if (position >= tagEnd)
            {
                break;
            }

            byte quote = data[position];
            if (quote != (byte)'"' && quote != (byte)'\'')
            {
                break;
            }
            position++;
            int valueStart = position;
            while (position < tagEnd && data[position] != quote)
            {
                position++;
            }
            int valueEnd = position;
            if (position < tagEnd)
            {
                position++;
            }

            if (data.Slice(nameStart, nameEnd - nameStart).SequenceEqual(wanted))
            {
                value = data.Slice(valueStart, valueEnd - valueStart);
                return true;
            }
        }

        value = default;
        return false;
    }

    private static bool? SstGetXmlSpace(ReadOnlySpan<byte> data, int tagStart, int tagEnd)
    {
        if (!SstTryGetAttribute(data, tagStart, tagEnd, "xml:space"u8, out ReadOnlySpan<byte> value))
        {
            return null;
        }
        if (value.IndexOf((byte)'&') < 0)
        {
            if (value.SequenceEqual("preserve"u8))
            {
                return true;
            }
            if (value.SequenceEqual("default"u8))
            {
                return false;
            }
        }
        else if (FastXmlText.TryDecode(value, out string decoded))
        {
            // An escaped value such as pre&#115;erve is legal XML; decode it before deciding.
            if (decoded == "preserve")
            {
                return true;
            }
            if (decoded == "default")
            {
                return false;
            }
        }
        return null;
    }

    private static bool SstDeclarationIsUtf8(ReadOnlySpan<byte> declaration)
    {
        int encoding = declaration.IndexOf("encoding"u8);
        if (encoding < 0)
        {
            return true;
        }

        int position = encoding + 8;
        while (position < declaration.Length && IsFastWhitespace(declaration[position]))
        {
            position++;
        }
        if (position >= declaration.Length || declaration[position] != (byte)'=')
        {
            return false;
        }
        position++;
        while (position < declaration.Length && IsFastWhitespace(declaration[position]))
        {
            position++;
        }
        if (position >= declaration.Length
            || (declaration[position] != (byte)'"' && declaration[position] != (byte)'\''))
        {
            return false;
        }

        byte quote = declaration[position++];
        int valueStart = position;
        while (position < declaration.Length && declaration[position] != quote)
        {
            position++;
        }
        ReadOnlySpan<byte> value = declaration.Slice(valueStart, position - valueStart);
        return FastAsciiEqualsIgnoreCase(value, "utf-8"u8)
            || FastAsciiEqualsIgnoreCase(value, "utf8"u8);
    }
}
