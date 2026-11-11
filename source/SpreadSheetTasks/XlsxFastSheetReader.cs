using System;
using System.Buffers;
using System.Buffers.Text;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.Unicode;
using System.Xml;

namespace SpreadSheetTasks;

// Fast XLSX worksheet row scanner.
//
// SpreadsheetML sheets are a very small subset of XML. XmlTextReaderImpl spends most of its time
// parsing attributes and shifting its buffers, so this scanner reads the row bytes directly and
// falls back to XmlReader for anything it is not sure about (entities, rich text, namespaces,
// unexpected elements). The fallback is per row, so a single unusual row never affects the rest.
public sealed partial class XlsxOrXlsbReadOrEdit
{
    private const int FastInitialBufferSize = 256 * 1024;
    private const int FastMaxBufferSize = 64 * 1024 * 1024;
    private const int FastMaxColumns = 16384;
    private const int FastMaxRows = 1_048_576;

    private byte[]? _fastBuffer;
    private int _fastPos;
    private int _fastLen;
    private bool _fastEof;
    internal bool _fastActive;
    private bool _fastRowHeaderPending;
    private bool _fastRowSelfClosing;
    private int _fastRowNumber;
    private int _fastRowStart;
    private int _fastRowContentPos;
    private int _fastRowEnd;
    private string _fastRootOpenTag = string.Empty;
    private bool _fastRootPreservesWhitespace;
    private bool _fastSheetDataPreservesWhitespace;
    private bool _fastRowPreservesWhitespace;
    private bool[]? _fastDateStyles;
    private bool _fastInvalidUtf8;
    private bool _fastValidateUtf8;
    private int _fastValidated;

    // Diagnostics used by the parity/regression tests and by the adaptive fallback policy.
    internal bool _fastXlsxReaderUsed;
    internal int _fastRowsParsed;
    internal int _fastRowsFallback;
    private int _fastConsecutiveFallbacks;

    /// <summary>
    /// Reads the worksheet prologue and starts the fast scanner. Returns false (and leaves the
    /// stream at an unspecified position) when the part is not a plain UTF-8 worksheet that the
    /// scanner can handle; the caller then re-opens the stream and uses XmlReader.
    /// </summary>
    private bool TryInitFastSheetReader()
    {
        _fastActive = false;
        _fastRowHeaderPending = false;
        _fastEof = false;
        _fastPos = 0;
        _fastLen = 0;
        _fastValidated = 0;
        _fastInvalidUtf8 = false;
        _fastValidateUtf8 = false;
        _fastXlsxReaderUsed = false;
        _fastRowsParsed = 0;
        _fastRowsFallback = 0;
        _fastConsecutiveFallbacks = 0;
        _fastRootPreservesWhitespace = false;
        _fastSheetDataPreservesWhitespace = false;
        _fastRowPreservesWhitespace = false;
        if (_fastBuffer is null)
        {
            _fastBuffer = new byte[FastInitialBufferSize];
        }

        if (!FastEnsure(2))
        {
            return false;
        }

        // Only UTF-8 is scanned directly. A UTF-16 BOM means XmlReader must do the decoding.
        byte b0 = _fastBuffer![0];
        byte b1 = _fastBuffer[1];
        if ((b0 == 0xFF && b1 == 0xFE) || (b0 == 0xFE && b1 == 0xFF))
        {
            return false;
        }

        if (b0 == 0xEF && b1 == 0xBB)
        {
            if (!FastEnsure(3))
            {
                return false;
            }
            if (_fastBuffer[2] == 0xBF)
            {
                _fastPos = 3;
            }
        }

        // From here on the part is expected to be UTF-8. Invalid bytes are handed back to
        // XmlReader (which raises the file's own error) rather than being replacement-decoded.
        _fastValidateUtf8 = true;
        if (!ValidateFastUtf8(finalCheck: false))
        {
            return false;
        }

        string? rootTag = null;
        int p = _fastPos;
        while (true)
        {
            if (!FastFindByte((byte)'<', ref p))
            {
                return false;
            }
            if (!FastEnsureFrom(p + 1, 1))
            {
                return false;
            }

            byte c = _fastBuffer[p + 1];
            if (c == '?')
            {
                int q = p + 2;
                if (!FastFindSequence("?>"u8, ref q))
                {
                    return false;
                }
                if (FastIsXmlDeclaration(p, q + 2) && !FastDeclarationIsUtf8(p + 5, q))
                {
                    return false; // XmlReader must decode the declared encoding
                }
                p = q + 2;
                continue;
            }

            if (c == '!')
            {
                if (!FastEnsureFrom(p, 4))
                {
                    return false;
                }
                if (FastStartsWith(p, "<!--"u8))
                {
                    int q = p + 4;
                    if (!FastFindSequence("-->"u8, ref q))
                    {
                        return false;
                    }
                    p = q + 3;
                    continue;
                }
                return false; // DOCTYPE, CDATA or other declaration
            }

            if (c == '/')
            {
                if (!FastReadStartTagFrom(p, out int endTag, out _))
                {
                    return false;
                }
                p = endTag;
                continue;
            }

            int nameStart = p + 1;
            int nameEnd = nameStart;
            while (FastEnsureFrom(nameEnd, 1) && FastIsNameByte(_fastBuffer[nameEnd]))
            {
                nameEnd++;
            }

            ReadOnlySpan<byte> name = _fastBuffer.AsSpan(nameStart, nameEnd - nameStart);
            if (name.IndexOf((byte)':') >= 0)
            {
                return false; // prefixed elements cannot be reconstructed for the fallback root
            }

            if (!FastReadStartTagFrom(p, out int tagEnd, out bool selfClosing))
            {
                return false;
            }

            if (name.SequenceEqual("worksheet"u8))
            {
                if (rootTag is not null)
                {
                    return false;
                }
                rootTag = Encoding.UTF8.GetString(_fastBuffer, p, tagEnd - p);
                _fastRootPreservesWhitespace = FastGetXmlSpace(p, tagEnd) ?? false;
            }
            else if (name.SequenceEqual("dimension"u8))
            {
                if (FastTryGetAttribute(p, tagEnd, "ref"u8, out ReadOnlySpan<byte> reference))
                {
                    _actualSheetDimensions = Encoding.UTF8.GetString(reference);
                }
            }
            else if (name.SequenceEqual("sheetData"u8))
            {
                if (rootTag is null)
                {
                    return false;
                }

                _fastRootOpenTag = rootTag;
                _fastSheetDataPreservesWhitespace = FastGetXmlSpace(p, tagEnd) ?? _fastRootPreservesWhitespace;
                BuildFastDateStyles();
                _fastPos = tagEnd;
                _fastActive = true;
                _fastXlsxReaderUsed = true;
                if (selfClosing)
                {
                    _fastEof = true;
                    _fastLen = _fastPos;
                }
                return true;
            }

            p = tagEnd;
        }
    }

    private void BuildFastDateStyles()
    {
        if (_stylesCellXfsArray is null)
        {
            _fastDateStyles = null;
            return;
        }

        var result = new bool[_stylesCellXfsArray.Length];
        for (int i = 0; i < result.Length; i++)
        {
            int numFmtId = _stylesCellXfsArray[i].NumFmtId;
            result[i] = _numberFormatsTypeDictionary.TryGetValue(numFmtId, out Type type) && type == typeof(DateTime?);
        }
        _fastDateStyles = result;
    }

    private bool FastIsXmlDeclaration(int start, int end)
    {
        if (end - start < 5 || !FastStartsWith(start, "<?xml"u8))
        {
            return false;
        }
        return start + 5 >= end || IsFastWhitespace(_fastBuffer![start + 5]);
    }

    private bool FastDeclarationIsUtf8(int start, int end)
    {
        for (int i = start; i + 8 < end; i++)
        {
            if (!_fastBuffer.AsSpan(i, 8).SequenceEqual("encoding"u8))
            {
                continue;
            }

            int j = i + 8;
            while (j < end && IsFastWhitespace(_fastBuffer[j]))
            {
                j++;
            }
            if (j >= end || _fastBuffer[j] != '=')
            {
                continue;
            }
            j++;
            while (j < end && IsFastWhitespace(_fastBuffer[j]))
            {
                j++;
            }
            if (j >= end || (_fastBuffer[j] != '"' && _fastBuffer[j] != '\''))
            {
                return false;
            }

            byte quote = _fastBuffer[j++];
            int valueStart = j;
            while (j < end && _fastBuffer[j] != quote)
            {
                j++;
            }
            ReadOnlySpan<byte> value = _fastBuffer.AsSpan(valueStart, j - valueStart);
            return FastAsciiEqualsIgnoreCase(value, "utf-8"u8)
                || FastAsciiEqualsIgnoreCase(value, "utf8"u8);
        }

        // No encoding declaration means the XML default, UTF-8.
        return true;
    }

    private static bool FastAsciiEqualsIgnoreCase(ReadOnlySpan<byte> value, ReadOnlySpan<byte> expected)
    {
        if (value.Length != expected.Length)
        {
            return false;
        }
        for (int i = 0; i < value.Length; i++)
        {
            int a = value[i];
            int b = expected[i];
            if (a >= 'A' && a <= 'Z')
            {
                a += 32;
            }
            if (b >= 'A' && b <= 'Z')
            {
                b += 32;
            }
            if (a != b)
            {
                return false;
            }
        }
        return true;
    }

    private bool ReadXlsxFast()
    {
        if (_emptyRowCnt < _howManyEmptyRow)
        {
            _emptyRowCnt++;
            _prevRowNum++;
            return true;
        }

        int rowNum;
        if (_fastRowHeaderPending)
        {
            rowNum = _fastRowNumber;
        }
        else
        {
            FastReadRowResult result = FastReadRowHeader(out rowNum);
            if (result == FastReadRowResult.EndOfSheet)
            {
                _fastActive = false;
                _sheetStream.Dispose();
                _xlsxSheetReaderInitialized = false;
                _isFirstRow = true; // RESET
                return false;
            }
            if (result == FastReadRowResult.UseXmlReader)
            {
                SwitchToXmlReaderForRemainder();
                return ReadXlsx();
            }
        }

        _howManyEmptyRow = -1;
        _emptyRowCnt = -1;

        if (_prevRowNum != -1 && rowNum > _prevRowNum + 1)
        {
            for (int i = 0; i < innerRow.Length; i++)
            {
                innerRow[i].type = ExcelDataType.Null;
            }
            _howManyEmptyRow = rowNum - _prevRowNum - 1;
            _prevRowNum++;
            _emptyRowCnt = 1;
            return true;
        }

        _prevRowNum = rowNum;

        for (int i = 0; i < innerRow.Length; i++)
        {
            innerRow[i].type = ExcelDataType.Null;
        }

        _prevColNum = -1;
        _rowWidth = 0;

        bool parsedFast = FastReadRowCells();
        if (!parsedFast)
        {
            FastParseRowWithXml();
            _fastRowsFallback++;
            _fastConsecutiveFallbacks++;
        }
        else
        {
            _fastConsecutiveFallbacks = 0;
        }
        _fastRowsParsed++;

        _fastRowHeaderPending = false;
        _fastPos = _fastRowEnd;

        if (_isFirstRow)
        {
            _isFirstRow = false;
            _columnsCntFromFirstRow = _numberOfLastColumnWithData - _numberOfFirstColumnWithData + 1;
            Array.Resize<FieldInfo>(ref innerRow, _columnsCntFromFirstRow);
            FieldCount = _columnsCntFromFirstRow;
        }
        if (_rowWidth > FieldCount)
        {
            FieldCount = _rowWidth;
        }

        // A sheet of rich inline strings or another construct the scanner declines should not
        // pay the per-row fallback cost for every row. Once the fallback rate is clearly high,
        // stream the rest of the part through XmlReader.
        if (!parsedFast && ShouldSwitchToXmlReader())
        {
            SwitchToXmlReaderForRemainder();
        }
        return true;
    }

    private bool ShouldSwitchToXmlReader()
    {
        if (_fastConsecutiveFallbacks >= 32)
        {
            return true;
        }
        return _fastRowsParsed >= 512 && _fastRowsFallback * 4 >= _fastRowsParsed;
    }

    private enum FastReadRowResult
    {
        Row,
        EndOfSheet,
        UseXmlReader,
    }

    /// <summary>
    /// Locates the next &lt;row&gt; start tag. Leaves the position on the tag and the pending
    /// flags set, because the caller may discover a row gap and must be able to parse it later.
    /// </summary>
    private FastReadRowResult FastReadRowHeader(out int rowNum)
    {
        rowNum = 0;
        FastCompactIfUseful();

        int p = _fastPos;
        while (true)
        {
            if (!FastSkipWhitespace(ref p))
            {
                _fastPos = p;
                return FastReadRowResult.EndOfSheet;
            }
            if (_fastBuffer![p] != '<')
            {
                _fastPos = p;
                return FastReadRowResult.UseXmlReader;
            }
            if (!FastEnsureFrom(p + 1, 1))
            {
                _fastPos = p;
                return FastReadRowResult.EndOfSheet;
            }

            byte c = _fastBuffer[p + 1];
            if (c == '!')
            {
                if (!FastEnsureFrom(p, 4))
                {
                    _fastPos = p;
                    return FastReadRowResult.EndOfSheet;
                }
                if (FastStartsWith(p, "<!--"u8))
                {
                    int q = p + 4;
                    if (!FastFindSequence("-->"u8, ref q))
                    {
                        _fastPos = p;
                        return FastReadRowResult.EndOfSheet;
                    }
                    p = q + 3;
                    _fastPos = p;
                    continue;
                }
                _fastPos = p;
                return FastReadRowResult.UseXmlReader;
            }

            if (c == '?')
            {
                int q = p + 2;
                if (!FastFindSequence("?>"u8, ref q))
                {
                    _fastPos = p;
                    return FastReadRowResult.EndOfSheet;
                }
                p = q + 2;
                _fastPos = p;
                continue;
            }

            if (c == '/')
            {
                int q = p;
                if (!FastReadStartTagFrom(q, out int endTag, out _))
                {
                    _fastPos = p;
                    return FastReadRowResult.EndOfSheet;
                }
                if (FastNameEquals(p + 2, "sheetData"u8))
                {
                    _fastPos = endTag;
                    return FastReadRowResult.EndOfSheet;
                }
                _fastPos = p;
                return FastReadRowResult.UseXmlReader;
            }

            int nameStart = p + 1;
            int nameEnd = nameStart;
            while (FastEnsureFrom(nameEnd, 1) && FastIsNameByte(_fastBuffer[nameEnd]))
            {
                nameEnd++;
            }

            ReadOnlySpan<byte> name = _fastBuffer.AsSpan(nameStart, nameEnd - nameStart);
            if (name.SequenceEqual("row"u8))
            {
                if (!FastReadStartTagFrom(p, out int tagEnd, out bool selfClosing))
                {
                    _fastPos = p;
                    return FastReadRowResult.EndOfSheet;
                }

                _fastRowPreservesWhitespace = FastGetXmlSpace(p, tagEnd) ?? _fastSheetDataPreservesWhitespace;
                rowNum = FastTryGetAttribute(p, tagEnd, "r"u8, out ReadOnlySpan<byte> rowRef)
                    ? FastParseUnsignedInt(rowRef)
                    : _prevRowNum + 1;
                if (rowNum < 1 || rowNum > FastMaxRows)
                {
                    _fastPos = p;
                    return FastReadRowResult.UseXmlReader;
                }

                _fastRowNumber = rowNum;
                _fastRowStart = p;
                _fastRowContentPos = tagEnd;
                _fastRowSelfClosing = selfClosing;
                _fastRowHeaderPending = true;
                _fastPos = p;
                return FastReadRowResult.Row;
            }

            // XmlReader would skip unknown elements inside sheetData. Skip them here too.
            if (!FastSkipElement(p, out int after))
            {
                _fastPos = p;
                return FastReadRowResult.UseXmlReader;
            }
            p = after;
            _fastPos = p;
        }
    }

    private bool FastReadRowCells()
    {
        int contentEnd;
        if (_fastRowSelfClosing)
        {
            _fastRowEnd = _fastRowContentPos;
            return true;
        }

        int q = _fastRowContentPos;
        if (!FastFindSequence("</row"u8, ref q))
        {
            _fastRowEnd = _fastLen;
            return false;
        }

        // A comment or CDATA section may contain the literal "</row". If any declaration
        // appears before the candidate, walk the row with those constructs skipped so the
        // row boundary is the real closing tag.
        int rowEnd;
        bool hasSkippableMarkup =
            _fastBuffer!.AsSpan(_fastRowContentPos, q - _fastRowContentPos).IndexOf("<!"u8) >= 0
            || _fastBuffer.AsSpan(_fastRowContentPos, q - _fastRowContentPos).IndexOf("<?"u8) >= 0;
        if (hasSkippableMarkup)
        {
            if (!FastFindRowEndSkippingMarkup(_fastRowContentPos, out q, out rowEnd))
            {
                _fastRowEnd = _fastLen;
                return false;
            }
        }
        else if (!FastReadStartTagFrom(q, out rowEnd, out _))
        {
            _fastRowEnd = q;
            return false;
        }
        contentEnd = q;
        _fastRowEnd = rowEnd;

        // Comments and processing instructions inside a row always take the XmlReader path;
        // the boundary above only exists so the fallback gets the complete row.
        if (hasSkippableMarkup)
        {
            return false;
        }

        int p = _fastRowContentPos;
        int cellIndex = 0;
        while (true)
        {
            if (!FastSkipWhitespace(ref p))
            {
                return false;
            }
            if (p >= contentEnd)
            {
                break;
            }
            if (_fastBuffer![p] != '<')
            {
                return false;
            }
            if (!FastEnsureFrom(p + 1, 1))
            {
                return false;
            }
            if (_fastBuffer[p + 1] == '!' || _fastBuffer[p + 1] == '?')
            {
                return false;
            }
            if (_fastBuffer[p + 1] == '/')
            {
                return false; // unexpected end tag inside the row
            }

            int nameStart = p + 1;
            int nameEnd = nameStart;
            while (FastEnsureFrom(nameEnd, 1) && FastIsNameByte(_fastBuffer[nameEnd]))
            {
                nameEnd++;
            }
            if (!_fastBuffer.AsSpan(nameStart, nameEnd - nameStart).SequenceEqual("c"u8))
            {
                return false;
            }

            if (!FastReadStartTagFrom(p, out int cellContentStart, out bool selfClosing))
            {
                return false;
            }

            if (!FastReadCellAttributes(p, cellContentStart, out int col, out int styleId, out char type, out bool? cellXmlSpace))
            {
                return false;
            }
            bool cellPreservesWhitespace = cellXmlSpace ?? _fastRowPreservesWhitespace;
            if (col < 0)
            {
                col = cellIndex;
                if (col >= FastMaxColumns)
                {
                    return false; // wider than a worksheet allows (no r attribute to check)
                }
            }
            cellIndex++;

            _colNum = col;
            EnsureColumnBase(col);
            if (_isFirstRow && _numberOfLastColumnWithData < col)
            {
                _numberOfLastColumnWithData = col;
            }

            int tmpLen = ResizeRowIfNeeded();
            _rowWidth = Math.Max(_rowWidth, tmpLen + 1);
            ref FieldInfo valueX = ref innerRow[tmpLen];

            p = cellContentStart;
            if (selfClosing)
            {
                continue;
            }

            if (!FastParseCellContent(ref p, type, styleId, cellPreservesWhitespace, ref valueX))
            {
                return false;
            }
            _prevColNum = col;
        }

        return true;
    }

    /// <summary>
    /// Locates the real &lt;/row&gt; when a comment or CDATA section before it may contain the
    /// literal string. Comments, CDATA sections and processing instructions are skipped whole;
    /// attribute quotes are honored by the regular start-tag reader.
    /// </summary>
    private bool FastFindRowEndSkippingMarkup(int start, out int contentEnd, out int rowEnd)
    {
        contentEnd = start;
        rowEnd = start;
        int p = start;
        while (true)
        {
            if (!FastFindByte((byte)'<', ref p))
            {
                return false;
            }
            if (!FastEnsureFrom(p + 1, 1))
            {
                return false;
            }

            byte c = _fastBuffer![p + 1];
            if (c == '!')
            {
                if (FastStartsWith(p, "<!--"u8))
                {
                    int end = p + 4;
                    if (!FastFindSequence("-->"u8, ref end))
                    {
                        return false;
                    }
                    p = end + 3;
                    continue;
                }
                if (FastStartsWith(p, "<![CDATA["u8))
                {
                    int end = p + 9;
                    if (!FastFindSequence("]]>"u8, ref end))
                    {
                        return false;
                    }
                    p = end + 3;
                    continue;
                }
                return false;
            }
            if (c == '?')
            {
                int end = p + 2;
                if (!FastFindSequence("?>"u8, ref end))
                {
                    return false;
                }
                p = end + 2;
                continue;
            }
            if (c == '/')
            {
                if (FastNameEquals(p + 2, "row"u8))
                {
                    contentEnd = p;
                    return FastReadStartTagFrom(p, out rowEnd, out _);
                }
                if (!FastReadStartTagFrom(p, out int afterEndTag, out _))
                {
                    return false;
                }
                p = afterEndTag;
                continue;
            }

            if (!FastReadStartTagFrom(p, out int afterStartTag, out _))
            {
                return false;
            }
            p = afterStartTag;
        }
    }

    /// <summary>
    /// Matches the canonical <c>&lt;v&gt;value&lt;/v&gt;&lt;/c&gt;</c> cell body. Leaves
    /// <paramref name="position"/> untouched on failure so the general scanner can run.
    /// </summary>
    private bool TryReadSimpleValueChild(int position, out int valueStart, out int valueEnd, out int afterCell)
    {
        valueStart = valueEnd = afterCell = position;
        if (!FastEnsureFrom(position, 3))
        {
            return false;
        }
        if (_fastBuffer![position] != '<' || _fastBuffer[position + 1] != 'v' || _fastBuffer[position + 2] != '>')
        {
            return false;
        }

        int start = position + 3;
        int close = start;
        if (!FastFindSequence("</v>"u8, ref close))
        {
            return false;
        }

        int after = close + 4;
        if (!FastEnsureFrom(after, 4))
        {
            return false;
        }
        if (_fastBuffer[after] != '<' || _fastBuffer[after + 1] != '/'
            || _fastBuffer[after + 2] != 'c' || _fastBuffer[after + 3] != '>')
        {
            return false;
        }

        valueStart = start;
        valueEnd = close;
        afterCell = after + 4;
        return true;
    }

    /// <summary>Parses the children of one &lt;c&gt; element and applies the value to <paramref name="valueX"/>.</summary>
    private bool FastParseCellContent(ref int p, char type, int styleId, bool cellPreservesWhitespace, ref FieldInfo valueX)
    {
        int valueStart = -1;
        int valueEnd = -1;
        bool valueSelfClosing = false;
        bool sawValue = false;
        bool sawInline = false;
        int inlineStart = -1;
        int inlineEnd = -1;
        bool inlineEmpty = false;
        bool inlinePreserveWhitespace = false;

        // Excel writes almost every value cell as <v>...</v></c>. Resolve that exact shape
        // with one vectorized search instead of the general child scanner.
        if (TryReadSimpleValueChild(p, out int simpleValueStart, out int simpleValueEnd, out int simpleAfterCell))
        {
            valueStart = simpleValueStart;
            valueEnd = simpleValueEnd;
            sawValue = true;
            p = simpleAfterCell;
            goto ValueReady;
        }

        while (true)
        {
            if (!FastSkipWhitespace(ref p))
            {
                return false;
            }
            if (!FastEnsureFrom(p, 1))
            {
                return false;
            }
            if (_fastBuffer![p] != '<')
            {
                return false;
            }
            if (!FastEnsureFrom(p + 1, 1))
            {
                return false;
            }

            if (_fastBuffer[p + 1] == '/')
            {
                if (!FastReadStartTagFrom(p, out int afterCell, out _)
                    || !FastNameEquals(p + 2, "c"u8))
                {
                    return false;
                }
                p = afterCell;
                break;
            }
            if (_fastBuffer[p + 1] == '!' || _fastBuffer[p + 1] == '?')
            {
                return false;
            }

            int nameStart = p + 1;
            int nameEnd = nameStart;
            while (FastEnsureFrom(nameEnd, 1) && FastIsNameByte(_fastBuffer[nameEnd]))
            {
                nameEnd++;
            }
            ReadOnlySpan<byte> name = _fastBuffer.AsSpan(nameStart, nameEnd - nameStart);

            if (name.SequenceEqual("v"u8))
            {
                if (!FastReadStartTagFrom(p, out int contentStart, out bool selfClosing))
                {
                    return false;
                }
                sawValue = true;
                if (selfClosing)
                {
                    valueStart = contentStart;
                    valueEnd = contentStart;
                    valueSelfClosing = true;
                    p = contentStart;
                }
                else
                {
                    valueStart = contentStart;
                    int close = contentStart;
                    if (!FastFindSequence("</v>"u8, ref close))
                    {
                        return false;
                    }
                    valueEnd = close;
                    p = close + 4;
                }
            }
            else if (name.SequenceEqual("f"u8))
            {
                if (!FastSkipElement(p, out int afterFormula))
                {
                    return false;
                }
                p = afterFormula;
            }
            else if (name.SequenceEqual("is"u8))
            {
                if (!FastReadStartTagFrom(p, out int isContentStart, out bool isSelfClosing))
                {
                    return false;
                }
                sawInline = true;
                if (isSelfClosing)
                {
                    inlineEmpty = true;
                    p = isContentStart;
                    continue;
                }

                int isContentEnd = isContentStart;
                if (!FastFindSequence("</is"u8, ref isContentEnd))
                {
                    return false;
                }

                int q = isContentStart;
                if (!FastSkipWhitespace(ref q) || !FastEnsureFrom(q, 2))
                {
                    return false;
                }
                if (_fastBuffer[q] != '<' || !FastNameEquals(q + 1, "t"u8))
                {
                    return false; // rich text (runs) or unexpected content
                }
                if (!FastReadStartTagFrom(q, out int textStart, out bool textSelfClosing))
                {
                    return false;
                }
                inlinePreserveWhitespace = cellPreservesWhitespace;
                if (FastGetXmlSpace(p, isContentStart) is bool isSpace)
                {
                    inlinePreserveWhitespace = isSpace;
                }
                if (FastGetXmlSpace(q, textStart) is bool textSpace)
                {
                    inlinePreserveWhitespace = textSpace;
                }
                if (textSelfClosing)
                {
                    inlineEmpty = true;
                }
                else
                {
                    inlineStart = textStart;
                    int textEnd = textStart;
                    if (!FastFindSequence("</t>"u8, ref textEnd))
                    {
                        return false;
                    }
                    inlineEnd = textEnd;

                    // XmlReader concatenates the text of every child of <is>; a second text
                    // element or a rich run after the first one needs that general path.
                    int afterText = textEnd + 4;
                    if (!FastSkipWhitespace(ref afterText) || afterText != isContentEnd)
                    {
                        return false;
                    }
                }

                int afterIs = isContentEnd;
                if (!FastReadStartTagFrom(afterIs, out int afterIsTag, out _))
                {
                    return false;
                }
                p = afterIsTag;
            }
            else
            {
                return false; // unsupported cell child
            }
        }

    ValueReady:
        if (sawInline)
        {
            valueX.type = ExcelDataType.String;
            if (inlineEmpty || inlineStart < 0)
            {
                valueX.strValue = string.Empty;
            }
            else
            {
                ReadOnlySpan<byte> inlineValue = _fastBuffer.AsSpan(inlineStart, inlineEnd - inlineStart);
                // XmlTextReader runs with IgnoreWhitespace, so a whitespace-only <t> without
                // xml:space="preserve" reads as an empty string. Match it to keep both paths equal.
                if (!inlinePreserveWhitespace
                    && inlineValue.IndexOfAny("&\r"u8) < 0
                    && FastIsAllXmlWhitespace(inlineValue))
                {
                    valueX.strValue = string.Empty;
                }
                else
                {
                    if (!FastXmlText.TryDecodeText(inlineValue, out string decoded))
                    {
                        return false;
                    }
                    if (!inlinePreserveWhitespace && FastXmlText.IsAllXmlWhitespace(decoded))
                    {
                        decoded = string.Empty;
                    }
                    valueX.strValue = DecodeXstring(decoded);
                }
            }
            return true;
        }

        if (!sawValue)
        {
            valueX.type = ExcelDataType.Null;
            return true;
        }

        ReadOnlySpan<byte> value = _fastBuffer.AsSpan(valueStart, valueEnd - valueStart);

        if (type == 's')
        {
            if (!Utf8Parser.TryParse(value, out int index, out int consumed) || consumed != value.Length)
            {
                if (value.Length == 0)
                {
                    index = 0;
                }
                else
                {
                    return false;
                }
            }
            valueX.type = ExcelDataType.String;
            valueX.strValue = _sharedStringArray[index];
            return true;
        }

        if (type == 'f') // t="str": formula string result
        {
            if (valueSelfClosing)
            {
                valueX.type = ExcelDataType.Null;
            }
            else
            {
                if (!FastXmlText.TryDecodeText(value, out string formulaText))
                {
                    return false;
                }
                valueX.type = ExcelDataType.String;
                valueX.strValue = DecodeXstring(formulaText);
            }
            return true;
        }

        if (TreatAllColumnsAsText)
        {
            if (!FastXmlText.TryDecodeText(value, out string text))
            {
                return false;
            }
            valueX.type = ExcelDataType.String;
            valueX.strValue = text;
            return true;
        }

        if (type == 'b')
        {
            if (value.IndexOf((byte)'&') >= 0)
            {
                return false;
            }
            valueX.type = ExcelDataType.Boolean;
            valueX.boolValue = value.Length > 0 && value[0] == (byte)'1';
            return true;
        }

        if (type == 'e')
        {
            if (value.IndexOf((byte)'&') >= 0)
            {
                return false;
            }
            valueX.type = ExcelDataType.String;
            valueX.strValue = "error in cell";
            return true;
        }

        if (type == 'd' || type == 'i') // t="d" is not handled by the existing reader either
        {
            return false;
        }

        // No type attribute: number, boolean handled above, date based on the cell style.
        if (styleId >= 0 && (_stylesCellXfsArray is null || styleId >= _stylesCellXfsArray.Length))
        {
            return false; // the XmlReader path reproduces the original exception
        }

        if (value.Length == 0)
        {
            valueX.type = ExcelDataType.Int64;
            valueX.int64Value = 0;
            return true;
        }

        if (styleId >= 0 && _fastDateStyles is not null && styleId < _fastDateStyles.Length && _fastDateStyles[styleId]
            && TryFastParseDouble(value, out double dateSerial))
        {
            valueX.type = ExcelDataType.DateTime;
            valueX.dtValue = ConvertFromOADate(dateSerial);
            return true;
        }

        if (ContainsDoubleMarks(value))
        {
            if (!TryFastParseDouble(value, out double doubleValue))
            {
                return false;
            }
            valueX.type = ExcelDataType.Double;
            valueX.doubleValue = doubleValue;
            return true;
        }

        if (!Utf8Parser.TryParse(value, out long longValue, out int longConsumed) || longConsumed != value.Length)
        {
            return false;
        }
        valueX.type = ExcelDataType.Int64;
        valueX.int64Value = longValue;
        return true;
    }

    private static bool ContainsDoubleMarks(ReadOnlySpan<byte> value)
        => value.IndexOfAny(".eE"u8) >= 0;

    /// <summary>
    /// Parses a plain decimal without going through the general number parser when it is safe.
    /// The fast path is limited to a mantissa of at most 2^53 and a decimal scale of at most
    /// 22, where the mantissa and 10^scale are both exact doubles; one multiply or divide then
    /// rounds the exact decimal value once, so the result is bit-identical to the correctly
    /// rounded result of <see cref="Utf8Parser"/>. Everything else (long tails, huge
    /// exponents, NaN/Infinity, whitespace) falls back to it as well.
    /// </summary>
    private static bool TryFastParseDouble(ReadOnlySpan<byte> value, out double result)
    {
        if (TryParseSimpleDouble(value, out result))
        {
            return true;
        }
        return Utf8Parser.TryParse(value, out result, out int consumed) && consumed == value.Length;
    }

    private static bool TryParseSimpleDouble(ReadOnlySpan<byte> value, out double result)
    {
        result = 0;
        int i = 0;
        bool negative = false;
        if (i < value.Length && (value[i] == (byte)'-' || value[i] == (byte)'+'))
        {
            negative = value[i] == (byte)'-';
            i++;
        }

        ulong mantissa = 0;
        int digits = 0;
        int fractionDigits = 0;
        bool anyDigit = false;
        for (; i < value.Length; i++)
        {
            uint digit = (uint)(value[i] - (byte)'0');
            if (digit > 9)
            {
                break;
            }
            if (digits >= 19)
            {
                return false;
            }
            mantissa = (mantissa * 10) + digit;
            digits++;
            anyDigit = true;
        }

        if (i < value.Length && value[i] == (byte)'.')
        {
            i++;
            for (; i < value.Length; i++)
            {
                uint digit = (uint)(value[i] - (byte)'0');
                if (digit > 9)
                {
                    break;
                }
                if (digits >= 19)
                {
                    return false;
                }
                mantissa = (mantissa * 10) + digit;
                digits++;
                fractionDigits++;
                anyDigit = true;
            }
        }

        if (!anyDigit)
        {
            return false;
        }

        int exponent = 0;
        if (i < value.Length && (value[i] == (byte)'e' || value[i] == (byte)'E'))
        {
            i++;
            bool exponentNegative = false;
            if (i < value.Length && (value[i] == (byte)'-' || value[i] == (byte)'+'))
            {
                exponentNegative = value[i] == (byte)'-';
                i++;
            }
            if (i >= value.Length)
            {
                return false;
            }

            int e = 0;
            bool anyExponentDigit = false;
            for (; i < value.Length; i++)
            {
                uint digit = (uint)(value[i] - (byte)'0');
                if (digit > 9)
                {
                    return false;
                }
                if (e > 100_000)
                {
                    return false;
                }
                e = (e * 10) + (int)digit;
                anyExponentDigit = true;
            }
            if (!anyExponentDigit)
            {
                return false;
            }
            exponent = exponentNegative ? -e : e;
        }

        if (i != value.Length)
        {
            return false;
        }

        if (mantissa > (1UL << 53))
        {
            return false;
        }

        int scale = fractionDigits - exponent;
        if (scale < -22 || scale > 22)
        {
            return false;
        }

        double parsed = mantissa;
        if (scale > 0)
        {
            parsed /= DecimalPowersOfTen[scale];
        }
        else if (scale < 0)
        {
            parsed *= DecimalPowersOfTen[-scale];
        }
        result = negative ? -parsed : parsed;
        return true;
    }

    private static readonly double[] DecimalPowersOfTen =
    {
        1e0, 1e1, 1e2, 1e3, 1e4, 1e5, 1e6, 1e7, 1e8, 1e9, 1e10,
        1e11, 1e12, 1e13, 1e14, 1e15, 1e16, 1e17, 1e18, 1e19, 1e20, 1e21, 1e22,
    };

    private static int FastParseUnsignedInt(ReadOnlySpan<byte> value)
        => Utf8Parser.TryParse(value, out int result, out _) ? result : 0;

    private static int FastParseColumn(ReadOnlySpan<byte> reference)
    {
        int column = -1;
        for (int i = 0; i < reference.Length; i++)
        {
            byte c = reference[i];
            if (c < 'A' || c > 'Z')
            {
                break;
            }
            column = ((column + 1) * 26) + (c - 'A');
            if (column >= FastMaxColumns)
            {
                return -1; // wider than an Excel worksheet; let XmlReader decide
            }
        }
        return column;
    }

    private void FastParseRowWithXml()
    {
        string rowXml = Encoding.UTF8.GetString(_fastBuffer!, _fastRowStart, _fastRowEnd - _fastRowStart);
        // Keep the same element depth as the real part: ReadCellsFromXml expects
        // worksheet > sheetData > row > c > v (depths 0..4).
        string document = _fastRootOpenTag + "<sheetData>" + rowXml + "</sheetData></worksheet>";

        var settings = new XmlReaderSettings
        {
            IgnoreComments = true,
            IgnoreWhitespace = true,
            NameTable = new SheetNameTable(),
        };

        using XmlReader reader = XmlReader.Create(new StringReader(document), settings);
        if (!reader.ReadToFollowing("row"))
        {
            return;
        }

        // Remove anything the fast attempt may have written before it gave up.
        for (int i = 0; i < innerRow.Length; i++)
        {
            innerRow[i].type = ExcelDataType.Null;
        }
        _rowWidth = 0;
        ReadCellsFromXml(reader);
    }

    /// <summary>
    /// The scanner met a construct it does not understand. Chain the already-buffered bytes in
    /// front of the still-open deflate stream and continue with XmlReader, preserving the row
    /// state already emitted. The rest of the sheet keeps streaming; unlike a MemoryStream copy
    /// this never materialises the unread remainder.
    /// The caller re-enters <see cref="ReadXlsx"/>, which relies on
    /// <c>_xlsxSheetReaderInitialized</c> (not <c>_isFirstRow</c>) to avoid restarting the scan.
    /// </summary>
    private void SwitchToXmlReaderForRemainder()
    {
        _fastActive = false;
        _fastRowHeaderPending = false;

        byte[] prefix = Encoding.UTF8.GetBytes(_fastRootOpenTag + "<sheetData>");
        int buffered = _fastLen - _fastPos;

        // Copy the (small) read-ahead tail: _fastBuffer is reused by the next sheet while the
        // fallback reader may still be alive, so the wrapper must own its bytes.
        byte[] head = new byte[prefix.Length + buffered];
        Buffer.BlockCopy(prefix, 0, head, 0, prefix.Length);
        Buffer.BlockCopy(_fastBuffer!, _fastPos, head, prefix.Length, buffered);

        // Ownership of the deflate stream moves to the wrapper; do not dispose it here.
        _sheetStream = new FastConcatStream(head, _sheetStream);

        var settings = new XmlReaderSettings
        {
            IgnoreComments = true,
            IgnoreWhitespace = true,
            NameTable = new SheetNameTable(),
        };
        _xmlReader = XmlReader.Create(_sheetStream, settings);
        _xmlReader.Read();
        while (!_xmlReader.EOF && !_xmlReader.IsStartElement("sheetData"))
        {
            if (_xmlReader.Depth == 0)
            {
                _xmlReader.Read();
            }
            else
            {
                _xmlReader.Skip();
            }
        }
    }

    /// <summary>
    /// Read-only stream serving a fixed byte prefix followed by a tail stream. Disposing the
    /// wrapper disposes the tail, so it can take over the <see cref="_sheetStream"/> field.
    /// </summary>
    private sealed class FastConcatStream : Stream
    {
        private byte[]? _head;
        private int _headPos;
        private readonly Stream _tail;

        internal FastConcatStream(byte[] head, Stream tail)
        {
            _head = head;
            _tail = tail;
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            int read = 0;
            if (_head is not null)
            {
                int available = _head.Length - _headPos;
                if (available > 0)
                {
                    int take = Math.Min(available, count);
                    Buffer.BlockCopy(_head, _headPos, buffer, offset, take);
                    _headPos += take;
                    read += take;
                    if (read == count)
                    {
                        return read;
                    }
                }
                if (_headPos >= _head.Length)
                {
                    _head = null; // release the prefix once consumed
                }
            }
            read += _tail.Read(buffer, offset + read, count - read);
            return read;
        }

        public override int Read(Span<byte> buffer)
        {
            int read = 0;
            if (_head is not null)
            {
                int available = _head.Length - _headPos;
                if (available > 0)
                {
                    int take = Math.Min(available, buffer.Length);
                    _head.AsSpan(_headPos, take).CopyTo(buffer);
                    _headPos += take;
                    read += take;
                    if (read == buffer.Length)
                    {
                        return read;
                    }
                }
                if (_headPos >= _head.Length)
                {
                    _head = null;
                }
            }
            read += _tail.Read(buffer[read..]);
            return read;
        }

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _tail.Dispose();
            }
            base.Dispose(disposing);
        }
    }

    // ── Low-level byte scanning ────────────────────────────────────────────────

    private bool FastEnsure(int count) => FastEnsureFrom(_fastPos, count);

    private bool FastEnsureFrom(int position, int count)
    {
        while (_fastLen - position < count)
        {
            if (!FastReadMore())
            {
                return _fastLen - position >= count;
            }
        }
        return true;
    }

    private bool FastReadMore()
    {
        if (_fastEof)
        {
            return false;
        }
        if (_fastLen == _fastBuffer!.Length)
        {
            if (_fastBuffer.Length >= FastMaxBufferSize)
            {
                throw new InvalidDataException(
                    $"XLSX worksheet holds more than {FastMaxBufferSize} bytes without closing a row");
            }
            int newSize = Math.Min(_fastBuffer.Length * 2, FastMaxBufferSize);
            Array.Resize(ref _fastBuffer, newSize);
        }

        int start = _fastLen;
        int read = _sheetStream.Read(_fastBuffer, _fastLen, _fastBuffer.Length - _fastLen);
        if (read <= 0)
        {
            _fastEof = true;
            if (_fastValidateUtf8 && !ValidateFastUtf8(finalCheck: true))
            {
                throw new InvalidDataException("XLSX worksheet XML is not valid UTF-8.");
            }
            return false;
        }
        _fastLen += read;
        if (_fastValidateUtf8 && !ValidateFastUtf8(finalCheck: false))
        {
            throw new InvalidDataException("XLSX worksheet XML is not valid UTF-8.");
        }
        return true;
    }

    /// <summary>
    /// Validates the bytes appended since the last call. A character split across two reads is
    /// left unvalidated until the next read; at EOF it becomes an error.
    /// </summary>
    private bool ValidateFastUtf8(bool finalCheck)
    {
        if (_fastInvalidUtf8)
        {
            return false;
        }

        ReadOnlySpan<byte> rest = _fastBuffer.AsSpan(_fastValidated, _fastLen - _fastValidated);
        if (rest.IsEmpty)
        {
            return true;
        }
        if (Utf8.IsValid(rest))
        {
            _fastValidated = _fastLen;
            return true;
        }

        int position = _fastValidated;
        while (position < _fastLen)
        {
            OperationStatus status = Rune.DecodeFromUtf8(
                _fastBuffer.AsSpan(position, _fastLen - position), out _, out int consumed);
            if (status == OperationStatus.Done)
            {
                position += consumed;
                continue;
            }
            if (status == OperationStatus.NeedMoreData && !finalCheck)
            {
                break;
            }

            _fastInvalidUtf8 = true;
            return false;
        }

        _fastValidated = position;
        return true;
    }

    private void FastCompactIfUseful()
    {
        if (_fastRowHeaderPending || _fastPos == 0)
        {
            return;
        }

        if (_fastPos < _fastLen - _fastPos)
        {
            return;
        }

        int remaining = _fastLen - _fastPos;

        int consumed = _fastPos;
        if (remaining > 0)
        {
            Buffer.BlockCopy(_fastBuffer!, consumed, _fastBuffer!, 0, remaining);
        }
        _fastPos = 0;
        _fastLen = remaining;
        _fastValidated = Math.Max(0, _fastValidated - consumed);
    }

    private bool FastSkipWhitespace(ref int p)
    {
        while (true)
        {
            if (!FastEnsureFrom(p, 1))
            {
                return false;
            }
            byte b = _fastBuffer![p];
            if (b != ' ' && b != '\t' && b != '\r' && b != '\n')
            {
                return true;
            }
            p++;
        }
    }

    private bool FastFindByte(byte needle, ref int p)
    {
        while (true)
        {
            if (!FastEnsureFrom(p, 1))
            {
                return false;
            }
            int index = _fastBuffer.AsSpan(p, _fastLen - p).IndexOf(needle);
            if (index >= 0)
            {
                p += index;
                return true;
            }
            p = _fastLen;
            if (!FastReadMore())
            {
                return false;
            }
        }
    }

    private bool FastFindSequence(ReadOnlySpan<byte> needle, ref int p)
    {
        int scanned = p;
        while (true)
        {
            if (scanned >= _fastLen)
            {
                if (!FastReadMore())
                {
                    return false;
                }
                continue;
            }
            int index = _fastBuffer.AsSpan(scanned, _fastLen - scanned).IndexOf(needle);
            if (index >= 0)
            {
                p = scanned + index;
                return true;
            }
            scanned = Math.Max(scanned, _fastLen - needle.Length + 1);
            if (!FastReadMore())
            {
                return false;
            }
        }
    }

    private bool FastStartsWith(int p, ReadOnlySpan<byte> value)
    {
        if (!FastEnsureFrom(p, value.Length))
        {
            return false;
        }
        return _fastBuffer.AsSpan(p, value.Length).SequenceEqual(value);
    }

    private bool FastNameEquals(int p, ReadOnlySpan<byte> name)
    {
        if (!FastEnsureFrom(p, name.Length))
        {
            return false;
        }
        if (!_fastBuffer.AsSpan(p, name.Length).SequenceEqual(name))
        {
            return false;
        }
        if (!FastEnsureFrom(p + name.Length, 1))
        {
            return true;
        }
        byte next = _fastBuffer[p + name.Length];
        return next == '>' || next == '/' || next == ' ' || next == '\t' || next == '\r' || next == '\n';
    }

    private bool FastReadStartTagFrom(int tagStart, out int tagEnd, out bool selfClosing)
    {
        tagEnd = tagStart;
        selfClosing = false;
        byte quote = 0;
        int i = tagStart + 1;
        while (true)
        {
            if (!FastEnsureFrom(i, 1))
            {
                return false;
            }
            byte b = _fastBuffer![i];
            if (quote != 0)
            {
                if (b == quote)
                {
                    quote = 0;
                }
            }
            else if (b == '"' || b == '\'')
            {
                quote = b;
            }
            else if (b == '>')
            {
                selfClosing = i > tagStart + 1 && _fastBuffer[i - 1] == '/';
                tagEnd = i + 1;
                return true;
            }
            i++;
        }
    }

    private bool FastTryGetAttribute(int tagStart, int tagEnd, ReadOnlySpan<byte> wanted, out ReadOnlySpan<byte> value)
    {
        int i = tagStart + 1;
        while (i < tagEnd && FastIsNameByte(_fastBuffer![i]))
        {
            i++;
        }

        while (i < tagEnd)
        {
            while (i < tagEnd && IsFastWhitespace(_fastBuffer![i]))
            {
                i++;
            }
            if (i >= tagEnd || _fastBuffer[i] == '/' || _fastBuffer[i] == '>')
            {
                break;
            }

            int nameStart = i;
            while (i < tagEnd && FastIsNameByte(_fastBuffer[i]))
            {
                i++;
            }
            int nameEnd = i;
            while (i < tagEnd && IsFastWhitespace(_fastBuffer[i]))
            {
                i++;
            }
            if (i >= tagEnd || _fastBuffer[i] != '=')
            {
                break;
            }
            i++;
            while (i < tagEnd && IsFastWhitespace(_fastBuffer[i]))
            {
                i++;
            }
            if (i >= tagEnd)
            {
                break;
            }

            byte quote = _fastBuffer[i];
            if (quote != '"' && quote != '\'')
            {
                break;
            }
            i++;
            int valueStart = i;
            while (i < tagEnd && _fastBuffer[i] != quote)
            {
                i++;
            }
            int valueEnd = i;
            if (i < tagEnd)
            {
                i++;
            }

            if (_fastBuffer.AsSpan(nameStart, nameEnd - nameStart).SequenceEqual(wanted))
            {
                value = _fastBuffer.AsSpan(valueStart, valueEnd - valueStart);
                return true;
            }
        }

        value = default;
        return false;
    }

    /// <summary>
    /// Reads <c>xml:space</c> from the start tag at <paramref name="tagStart"/>:
    /// <c>true</c> for <c>preserve</c>, <c>false</c> for <c>default</c>, <c>null</c> when unspecified.
    /// </summary>
    private bool? FastGetXmlSpace(int tagStart, int tagEnd)
    {
        if (!FastTryGetAttribute(tagStart, tagEnd, "xml:space"u8, out ReadOnlySpan<byte> value))
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

    /// <summary>
    /// Reads the <c>r</c>, <c>s</c> and <c>t</c> attributes of one cell start tag in a single
    /// walk, together with <c>xml:space</c> (which is inherited by the inline string).
    /// Looking each one up separately walks every cell tag several times, which is pure
    /// overhead because all three cell attributes have one-character names and can be
    /// dispatched on the fly. Returns false when the tag holds something the scanner must not
    /// interpret.
    /// </summary>
    private bool FastReadCellAttributes(int tagStart, int tagEnd, out int column, out int styleId, out char type, out bool? xmlSpace)
    {
        column = -1;
        styleId = -1;
        type = '\0';
        xmlSpace = null;

        int i = tagStart + 1;
        while (i < tagEnd && FastIsNameByte(_fastBuffer![i]))
        {
            i++;
        }

        while (i < tagEnd)
        {
            while (i < tagEnd && IsFastWhitespace(_fastBuffer![i]))
            {
                i++;
            }
            if (i >= tagEnd || _fastBuffer[i] == '/' || _fastBuffer[i] == '>')
            {
                break;
            }

            int nameStart = i;
            while (i < tagEnd && FastIsNameByte(_fastBuffer[i]))
            {
                i++;
            }
            int nameLength = i - nameStart;
            while (i < tagEnd && IsFastWhitespace(_fastBuffer[i]))
            {
                i++;
            }
            if (i >= tagEnd || _fastBuffer[i] != '=')
            {
                break;
            }
            i++;
            while (i < tagEnd && IsFastWhitespace(_fastBuffer[i]))
            {
                i++;
            }
            if (i >= tagEnd)
            {
                break;
            }

            byte quote = _fastBuffer[i];
            if (quote != '"' && quote != '\'')
            {
                break;
            }
            i++;
            int valueStart = i;
            while (i < tagEnd && _fastBuffer[i] != quote)
            {
                i++;
            }
            int valueEnd = i;
            if (i < tagEnd)
            {
                i++;
            }

            ReadOnlySpan<byte> value = _fastBuffer.AsSpan(valueStart, valueEnd - valueStart);

            // xml:space is inherited by the inline string inside the cell.
            if (nameLength == 9 && value.Length <= 8
                && _fastBuffer.AsSpan(nameStart, nameLength).SequenceEqual("xml:space"u8))
            {
                if (value.IndexOf((byte)'&') >= 0)
                {
                    return false; // an escaped xml:space value needs the XML parser
                }
                if (value.SequenceEqual("preserve"u8))
                {
                    xmlSpace = true;
                }
                else if (value.SequenceEqual("default"u8))
                {
                    xmlSpace = false;
                }
                continue;
            }

            // Only r, s and t carry meaning on a cell; every other attribute is ignored.
            if (nameLength != 1)
            {
                continue;
            }
            if (value.IndexOf((byte)'&') >= 0)
            {
                return false; // escaped attribute values need the XML parser
            }

            switch (_fastBuffer[nameStart])
            {
                case (byte)'r':
                    column = FastParseColumn(value);
                    if (column < 0)
                    {
                        return false;
                    }
                    break;

                case (byte)'s':
                    if (!Utf8Parser.TryParse(value, out styleId, out int styleConsumed) || styleConsumed != value.Length)
                    {
                        return false;
                    }
                    break;

                case (byte)'t':
                    if (value.Length == 0)
                    {
                        return false;
                    }
                    char cellType = (char)value[0];
                    if (cellType != 's' && cellType != 'b' && cellType != 'e' && cellType != 'i' && cellType != 'n' && cellType != 'd')
                    {
                        return false;
                    }
                    if (cellType == 's' && value.SequenceEqual("str"u8))
                    {
                        cellType = 'f'; // formula string, same handling as t="str"
                    }
                    else if (cellType == 's' && value.Length != 1)
                    {
                        return false;
                    }
                    else if (cellType == 'i' && !value.SequenceEqual("inlineStr"u8))
                    {
                        return false;
                    }
                    type = cellType;
                    break;
            }
        }

        return true;
    }

    /// <summary>Skips an element and its subtree; used for elements the sheet schema does not define.</summary>
    private bool FastSkipElement(int tagStart, out int after)
    {
        after = tagStart;
        int nameStart = tagStart + 1;
        int nameEnd = nameStart;
        while (FastEnsureFrom(nameEnd, 1) && FastIsNameByte(_fastBuffer![nameEnd]))
        {
            nameEnd++;
        }
        ReadOnlySpan<byte> name = _fastBuffer.AsSpan(nameStart, nameEnd - nameStart);
        if (name.Length == 0)
        {
            return false;
        }

        if (!FastReadStartTagFrom(tagStart, out int pos, out bool selfClosing))
        {
            return false;
        }
        if (selfClosing)
        {
            after = pos;
            return true;
        }

        int depth = 1;
        while (depth > 0)
        {
            int q = pos;
            if (!FastFindByte((byte)'<', ref q))
            {
                return false;
            }
            if (!FastEnsureFrom(q + 1, 1))
            {
                return false;
            }

            byte c = _fastBuffer[q + 1];
            if (c == '!')
            {
                if (!FastEnsureFrom(q, 4))
                {
                    return false;
                }
                if (FastStartsWith(q, "<!--"u8))
                {
                    int commentEnd = q + 4;
                    if (!FastFindSequence("-->"u8, ref commentEnd))
                    {
                        return false;
                    }
                    pos = commentEnd + 3;
                    continue;
                }
                if (FastStartsWith(q, "<![CDATA["u8))
                {
                    int cdataEnd = q + 9;
                    if (!FastFindSequence("]]>"u8, ref cdataEnd))
                    {
                        return false;
                    }
                    pos = cdataEnd + 3;
                    continue;
                }
                return false;
            }
            if (c == '?')
            {
                int piEnd = q + 2;
                if (!FastFindSequence("?>"u8, ref piEnd))
                {
                    return false;
                }
                pos = piEnd + 2;
                continue;
            }
            if (c == '/')
            {
                if (FastNameEquals(q + 2, name))
                {
                    if (!FastReadStartTagFrom(q, out int end, out _))
                    {
                        return false;
                    }
                    depth--;
                    pos = end;
                }
                else
                {
                    if (!FastReadStartTagFrom(q, out int end, out _))
                    {
                        return false;
                    }
                    pos = end;
                }
                continue;
            }

            if (FastNameEquals(q + 1, name))
            {
                if (!FastReadStartTagFrom(q, out int end, out bool nestedSelfClosing))
                {
                    return false;
                }
                if (!nestedSelfClosing)
                {
                    depth++;
                }
                pos = end;
            }
            else
            {
                if (!FastReadStartTagFrom(q, out int end, out _))
                {
                    return false;
                }
                pos = end;
            }
        }

        after = pos;
        return true;
    }

    private static bool FastIsNameByte(byte b)
        => b != ' ' && b != '\t' && b != '\r' && b != '\n'
        && b != '=' && b != '/' && b != '>' && b != '"' && b != '\'';

    private static bool IsFastWhitespace(byte b)
        => b == ' ' || b == '\t' || b == '\r' || b == '\n';

    private static bool FastIsAllXmlWhitespace(ReadOnlySpan<byte> value)
    {
        foreach (byte b in value)
        {
            if (!IsFastWhitespace(b))
            {
                return false;
            }
        }
        return true;
    }
}




