using System;
using System.Buffers;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;


namespace SpreadSheetTasks
{
    public class XlsxWriter : ExcelWriter, IDisposable
    {
        private readonly bool _inMemoryMode;
        private readonly int _bufferSize;
        internal static readonly string[] _letters;
        private readonly bool _useScharedStrings;

        private readonly CompressionLevel _clvl;

        private const double _dateTimeWidth = 16.0;
        private const double _dateWidth = 10.140625;

        /// <summary>cellXfs index for date style (yyyy-mm-dd style numFmtId 14).</summary>
        internal const int XfStyleDate = 1;
        /// <summary>cellXfs index for datetime style (yyyy-mm-dd hh:mm:ss numFmtId 22).</summary>
        internal const int XfStyleDateTime = 2;
        /// <summary>cellXfs index for bold header (font 1).</summary>
        internal const int XfStyleBoldHeader = 3;

        /// <summary>Built-in number-format id for dates (numFmtId 14).</summary>
        private const int NumFmtIdDate = 14;
        /// <summary>Built-in number-format id for date-times (numFmtId 22).</summary>
        private const int NumFmtIdDateTime = 22;

        /// <summary>
        /// Creates a new XLSX writer that will write to a file at <paramref name="filePath"/>.
        /// The file is overwritten if it already exists.
        /// </summary>
        /// <param name="filePath">Path of the .xlsx file to create.</param>
        /// <param name="bufferSize">Size of the I/O write buffer.</param>
        /// <param name="InMemoryMode">If true, sheet content is written directly to the zip entry stream; if false, a temp file is used per sheet.</param>
        /// <param name="useScharedStrings">If true, string values are deduplicated via the sharedStrings part.</param>
        /// <param name="_clvl">Compression level for the zip archive.</param>
        public XlsxWriter(string filePath, int bufferSize = 4096, bool InMemoryMode = true, bool useScharedStrings = true, CompressionLevel _clvl = CompressionLevel.Optimal)
        {
            ArgumentNullException.ThrowIfNull(filePath);
            ArgumentException.ThrowIfNullOrWhiteSpace(filePath);

            FileStream fs = null!;
            try
            {
                fs = new FileStream(filePath, FileMode.Create);
                _newExcelFileStream = fs;
                _bufferSize = bufferSize;
                _inMemoryMode = InMemoryMode;
                TryToSpecifyWidthForMemoryMode = InMemoryMode;
                _useScharedStrings = useScharedStrings;
                if (_useScharedStrings)
                    _sstDic = new Dictionary<string, int>();
                this._clvl = _clvl;
                _excelArchiveFile = new ZipArchive(_newExcelFileStream, ZipArchiveMode.Create, leaveOpen: false);
                _excelStreamWasProvided = false;
            }
            catch
            {
                fs?.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Creates a new XLSX writer that writes to the supplied <see cref="Stream"/>.
        /// </summary>
        /// <param name="stream">Destination stream. The writer takes ownership of the stream position but does not close it unless <paramref name="leaveExcelArchiveOpen"/> is false.</param>
        /// <param name="bufferSize">Size of the I/O write buffer.</param>
        /// <param name="InMemoryMode">If true, sheet content is written directly to the zip entry stream.</param>
        /// <param name="useScharedStrings">If true, string values are deduplicated via the sharedStrings part.</param>
        /// <param name="_clvl">Compression level for the zip archive.</param>
        /// <param name="leaveExcelArchiveOpen">If true, the underlying <see cref="ZipArchive"/> is left open when <see cref="Save"/> runs. Defaults to true for the stream ctor; the file-path ctor always sets it to false.</param>
        public XlsxWriter(Stream stream, int bufferSize = 4096, bool InMemoryMode = true, bool useScharedStrings = true, CompressionLevel _clvl = CompressionLevel.Optimal, bool leaveExcelArchiveOpen = true)
        {
            ArgumentNullException.ThrowIfNull(stream);
            _excelStreamWasProvided = true;
            _newExcelFileStream = stream;
            _bufferSize = bufferSize;
            _inMemoryMode = InMemoryMode;
            TryToSpecifyWidthForMemoryMode = InMemoryMode;
            _useScharedStrings = useScharedStrings;
            if (_useScharedStrings)
                _sstDic = new Dictionary<string, int>();
            this._clvl = _clvl;
            _excelArchiveFile = new ZipArchive(_newExcelFileStream, ZipArchiveMode.Create, leaveOpen: leaveExcelArchiveOpen);
        }


        static XlsxWriter()
        {
            var lettersX = new List<string>();

            for (int i = 65; i < 91; i++)
            {
                lettersX.Add(((char)i).ToString());
            }

            var tempList = new List<string>();
            foreach (var item in lettersX)
            {

                for (int i = 65; i < 91; i++)
                {
                    tempList.Add(item + ((char)i).ToString());
                }
            }

            var tempList2 = new List<string>();
            foreach (var item in tempList)
            {
                if (item.CompareTo("XY") > 0)
                {
                    break;
                }

                for (int i = 65; i < 91; i++)
                {
                    tempList2.Add(item + ((char)i).ToString());
                }
            }

            lettersX.AddRange(tempList);
            lettersX.AddRange(tempList2);
            _letters = lettersX.ToArray();
        }
        private static string GetTempFileFullPath()
        {
            return $"{Path.GetTempPath()}\\{Path.GetRandomFileName()}";
        }
        public override void Save()
        {
            DoOnCompress();
            if (!_inMemoryMode)
            {
                for (int i = 0; i < sheetCnt + 1; i++)
                {
                    _excelArchiveFile.CreateEntryFromFile(_sheetList[i].pathOnDisc, _sheetList[i].pathInArchive, _clvl);
                    File.Delete(_sheetList[i].pathOnDisc);
                }
            }
            base.Save();
        }
        public override void AddSheet(string sheetName, bool hidden = false)
        {
            string archveSheetName = "sheet" + (sheetCnt + 2);
            _sheetList.Add((sheetName, String.Format(@"xl/worksheets/{0}.xml", archveSheetName), XlsxWriter.GetTempFileFullPath(), hidden, archveSheetName, (sheetCnt + 2),null));
            //_sheetList.Add((sheetName, String.Format(@"xl/worksheets/{0}.xml", sheetName), getTempFileFullPath(), hidden, sheetName));
        }
        public override void WriteSheet(IDataReader dataReader, Boolean headers = true, int maxRows = -1, int startingRow = 0, int startingColumn = 0, bool doAutofilter = false)
        {
            sheetCnt++;
            this._areHeaders = headers;
            _dataColReader = new DataColReader(dataReader, headers, maxRows);
            if (_inMemoryMode)
            {
                var e1 = _excelArchiveFile.CreateEntry(_sheetList[sheetCnt].pathInArchive, _clvl);
                using StreamWriter daneZakladkiX = new FormattingStreamWriter(e1.Open(), Encoding.UTF8, bufferSize: _bufferSize, CultureInfo.InvariantCulture.NumberFormat);
                _rowsCount = WriteSheet(daneZakladkiX, startingRow, startingColumn, doAutofilter: doAutofilter) - 1;
            }
            else
            {
                using StreamWriter sheedData = new FormattingStreamWriter(_sheetList[sheetCnt].pathOnDisc, false, Encoding.UTF8, bufferSize: _bufferSize, CultureInfo.InvariantCulture.NumberFormat);
                _rowsCount = WriteSheet(sheedData, startingRow, startingColumn, doAutofilter: doAutofilter) - 1;
            }
        }
        public override void WriteSheet(DataTable dataTable, Boolean headers = true, int maxRows = -1, int startingRow = 0, int startingColumn = 0, bool doAutofilter = false)
        {
            sheetCnt++;
            this._areHeaders = headers;
            _dataColReader = new DataColReader(dataTable, headers, maxRows);
            if (_inMemoryMode)
            {
                var e1 = _excelArchiveFile.CreateEntry(_sheetList[sheetCnt].pathInArchive, _clvl);
                using StreamWriter daneZakladkiX = new FormattingStreamWriter(e1.Open(), Encoding.UTF8, bufferSize: _bufferSize, CultureInfo.InvariantCulture.NumberFormat);
                _rowsCount = WriteSheet(daneZakladkiX, startingRow, startingColumn,doAutofilter: doAutofilter) - 1;
            }
            else
            {
                using StreamWriter daneZakladkiX = new FormattingStreamWriter(_sheetList[sheetCnt].pathOnDisc, false, Encoding.UTF8, bufferSize: _bufferSize, CultureInfo.InvariantCulture.NumberFormat);
                _rowsCount = WriteSheet(daneZakladkiX, startingRow, startingColumn, doAutofilter: doAutofilter) - 1;
            }
        }

        private static readonly CultureInfo _invariantCulture = CultureInfo.InvariantCulture;
        private const int BUFFER_SIZE = 65_536;
        private const int BUFFER_SIZE_HALF = BUFFER_SIZE / 2;
        private readonly char[] _buffer = new char[BUFFER_SIZE];
        private int _currentBufferOffset = 0;
        private int _currentStartingColumn;

        private void WriteByteToBuffer(byte val)
        {
            val.TryFormat(_buffer.AsSpan(_currentBufferOffset), out int len, default, _invariantCulture);
            _currentBufferOffset += len;
        }
        private void WritesByteToBuffer(sbyte val)
        {
            val.TryFormat(_buffer.AsSpan(_currentBufferOffset), out int len, default, _invariantCulture);
            _currentBufferOffset += len;
        }

        private void WriteInt16ToBuffer(Int16 val)
        {
            val.TryFormat(_buffer.AsSpan(_currentBufferOffset), out int len, default, _invariantCulture);
            _currentBufferOffset += len;
        }

        private void WriteInt32ToBuffer(Int32 val)
        {
            val.TryFormat(_buffer.AsSpan(_currentBufferOffset), out int len, default, _invariantCulture);
            _currentBufferOffset += len;
        }
        private void WriteInt64ToBuffer(Int64 val)
        {
            val.TryFormat(_buffer.AsSpan(_currentBufferOffset), out int len, default, _invariantCulture);
            _currentBufferOffset += len;
        }
        private void WriteDoubleToBuffer(double val)
        {
            val.TryFormat(_buffer.AsSpan(_currentBufferOffset), out int len, default, _invariantCulture);
            _currentBufferOffset += len;
        }
        private void WriteFloatToBuffer(float val)
        {
            val.TryFormat(_buffer.AsSpan(_currentBufferOffset), out int len, default, _invariantCulture);
            _currentBufferOffset += len;
        }

        //private void writeStringToBuffer(string val)
        //{
        //    for (int i = 0; i < val.Length; i++)
        //    {
        //        char c = val[i];
        //        buffer[currentBufferOffset++] = c;
        //    }
        //}
        //[MethodImpl(MethodImplOptions.AggressiveOptimization)]
        private void WriteStringToBuffer(ReadOnlySpan<char> val)
        {
            val.CopyTo(_buffer.AsSpan(_currentBufferOffset));
            _currentBufferOffset += val.Length;
        }

        private void WriteXmlEscapedToBuffer(ReadOnlySpan<char> val)
        {
            int specialIdx = val.IndexOfAny(['&', '<', '>', '"', '\'']);
            if (specialIdx < 0)
            {
                val.CopyTo(_buffer.AsSpan(_currentBufferOffset));
                _currentBufferOffset += val.Length;
                return;
            }
            if (_currentBufferOffset + val.Length * 6 >= _buffer.Length)
            {
                return;
            }
            val[..specialIdx].CopyTo(_buffer.AsSpan(_currentBufferOffset));
            _currentBufferOffset += specialIdx;
            for (int i = specialIdx; i < val.Length; i++)
            {
                char c = val[i];
                switch (c)
                {
                    case '&':
                        "&amp;".CopyTo(_buffer.AsSpan(_currentBufferOffset));
                        _currentBufferOffset += 5;
                        break;
                    case '<':
                        "&lt;".CopyTo(_buffer.AsSpan(_currentBufferOffset));
                        _currentBufferOffset += 4;
                        break;
                    case '>':
                        "&gt;".CopyTo(_buffer.AsSpan(_currentBufferOffset));
                        _currentBufferOffset += 4;
                        break;
                    case '"':
                        "&quot;".CopyTo(_buffer.AsSpan(_currentBufferOffset));
                        _currentBufferOffset += 6;
                        break;
                    case '\'':
                        "&apos;".CopyTo(_buffer.AsSpan(_currentBufferOffset));
                        _currentBufferOffset += 6;
                        break;
                    default:
                        _buffer[_currentBufferOffset++] = c;
                        break;
                }
            }
        }

        public bool TryToSpecifyWidthForMemoryMode { get; set; }
        private int WriteSheet(StreamWriter sheetWritter, int startingRow, int startingColumn, bool doAutofilter = false)
        {
            _currentStartingColumn = startingColumn;
            if (doAutofilter)
            {
                _autofilterIsOn = true;
            }
            int rowNum = 0;

            int ColumnCount = _dataColReader.FieldCount;
            _colWidthsArray = new double[ColumnCount];
            Array.Fill<double>(_colWidthsArray, -1.0);

            typesArray = new int[ColumnCount];
            _newTypes = new TypeCode[ColumnCount];

            if (_inMemoryMode)
            {
                sheetWritter.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                sheetWritter.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");

                if (TryToSpecifyWidthForMemoryMode && _dataColReader._dataReader != null)
                {
                    var rdr = _dataColReader._dataReader;
                    for (int l = 1; l <= ColumnCount; l++)
                    {
                        int lenn = rdr.GetName(l - 1).Length + (doAutofilter ?2:0);
                        double tempWidth = 1.25 * lenn + 2;
                        if (tempWidth > _MAX_WIDTH)
                        {
                            tempWidth = _MAX_WIDTH;
                        }
                        if (_colWidthsArray[l - 1] < tempWidth)
                        {
                            _colWidthsArray[l - 1] = tempWidth;
                        }
                    }

                    int nr = 0;
                    _dataColReader.top100 = new List<object[]>();
                    bool areNextRows = true;
                    while (nr < 100)
                    {
                        areNextRows = rdr.Read();
                        if (!areNextRows)
                        {
                            break;
                        }

                        object[] arr = new object[rdr.FieldCount];
                        rdr.GetValues(arr);

                        for (int i = 0; i < arr.Length; i++)
                        {
                            if (arr[i] is Memory<byte> mem)
                            {
                                Memory<byte> m = new byte[mem.Length];
                                mem.CopyTo(m);
                                arr[i] = m;
                            }
                        }

                        _dataColReader.top100.Add(arr);
                        nr++;
                        SetColsLength(ColumnCount, arr);
                    }
                    areNextRows = rdr.Read();
                    _dataColReader.AreNextRows = areNextRows;
                }
                else if (TryToSpecifyWidthForMemoryMode && _dataColReader._dataTable != null)
                {
                    _dataColReader.GetWidthFromDataTable(_colWidthsArray, _MAX_WIDTH, doAutofilter);
                }

                // Write dimension first (required before sheetViews per schema)
                if (TryToSpecifyWidthForMemoryMode && _dataColReader._dataTable != null)
                {
                    sheetWritter.Write($"<dimension ref=\"{_letters[startingColumn]}{startingRow + 1}:{_letters[ColumnCount - 1 + startingColumn]}{_dataColReader._dataTableRowsCount + 1 + startingRow}\"/>");
                }

                if (doAutofilter)
                {
                    if (sheetCnt == 0)
                        sheetWritter.Write("<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/><selection pane=\"bottomLeft\"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"15\"/>");
                    else
                        sheetWritter.Write("<sheetViews><sheetView workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/><selection pane=\"bottomLeft\"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"15\"/>");
                }

                if (TryToSpecifyWidthForMemoryMode && _dataColReader._dataReader != null)
                {
                    sheetWritter.Write("<cols>");
                    for (int l = 1; l <= ColumnCount; l++)
                    {
                        sheetWritter.Write(String.Format(CultureInfo.InvariantCulture.NumberFormat, "<col min=\"{0}\" max=\"{0}\" width=\"{1}\" bestFit = \"1\" customWidth=\"1\" />", l + startingColumn, _colWidthsArray[l - 1]));
                    }
                    sheetWritter.Write("</cols>");
                }
                else if (TryToSpecifyWidthForMemoryMode && _dataColReader._dataTable != null)
                {
                    sheetWritter.Write("<cols>");
                    for (int l = 1; l <= ColumnCount; l++)
                    {
                        sheetWritter.Write(String.Format(CultureInfo.InvariantCulture.NumberFormat, "<col min=\"{0}\" max=\"{0}\" width=\"{1}\" bestFit = \"1\" customWidth=\"1\" />", l + startingColumn, _colWidthsArray[l - 1]));
                    }
                    sheetWritter.Write("</cols>");
                }
                sheetWritter.Write("<sheetData>");
                _colWidthsArray = null;
            }
            else
            {
                for (int i = 0; i < 600 + ColumnCount * 100; i++)
                {
                    sheetWritter.Write(" ");
                }
                sheetWritter.WriteLine();
            }

            while (_dataColReader.Read())
            {
                if (rowNum == 0 || this._areHeaders && rowNum == 1)
                {
                    if (rowNum == 0 && this._areHeaders)
                    {
                        for (int i = 0; i < ColumnCount; i++)
                        {
                            typesArray[i] = 0;
                            _newTypes[i] = TypeCode.String;
                        }
                    }
                    else
                    {
                        ExcelWriter.SetTypes(_dataColReader, typesArray, _newTypes, ColumnCount);
                    }
                }

                if (startingRow != 0)
                {
                    WriteStringToBuffer("<row r=\"");
                    WriteInt32ToBuffer(rowNum + 1 + startingRow);
                    WriteStringToBuffer("\">");
                }
                else
                {
                    WriteStringToBuffer("<row>");
                }

                Dictionary<int, CellStyle>? rowStyles = GetRowStyles(rowNum, startingRow, startingColumn, ColumnCount);
                WriteRow(ColumnCount, rowNum + 1 + startingRow, rowStyles, boldedStyle: _areHeaders && rowNum == 0);

                rowNum++;
                WriteStringToBuffer("</row>");
                if (_currentBufferOffset >= BUFFER_SIZE_HALF)
                {
                    sheetWritter.Write(_buffer, 0, _currentBufferOffset);
                    _currentBufferOffset = 0;
                }
                if (rowNum % 10000 == 0)
                {
                    DoOn10k(rowNum);
                }
            }
            if (_currentBufferOffset > 0)
            {
                sheetWritter.Write(_buffer, 0, _currentBufferOffset);
                _currentBufferOffset = 0;
            }

            sheetWritter.Write("</sheetData>");
            if (doAutofilter)
            {
                (string name, string pathInArchive, string pathOnDisc, bool isHidden, string nameInArchive, int sheetId, string _) = this._sheetList[^1];
                this._sheetList[^1] = (name, pathInArchive, pathOnDisc, isHidden, nameInArchive, sheetId, $"{name}!${_letters[startingColumn]}${startingRow + 1}:${_letters[ColumnCount - 1 + startingColumn]}${rowNum}");

                sheetWritter.Write($"<autoFilter ref=\"{_letters[startingColumn]}{startingRow + 1}:{_letters[ColumnCount - 1 + startingColumn]}{rowNum + startingRow}\"/>");
            }
            sheetWritter.Write("<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/></worksheet>");

            //System.NotSupportedException: 'This stream from ZipArchiveEntry does not support seeking.'
            if (!_inMemoryMode)
            {
                sheetWritter.Flush();
                sheetWritter.BaseStream.Position = 0;
                sheetWritter.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                sheetWritter.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
                sheetWritter.Write(String.Format("<dimension ref=\"A1:{0}{1}\"/>", _letters[ColumnCount - 1], rowNum + 1));

                if (doAutofilter)
                {
                    if (sheetCnt == 0)
                    {
                        sheetWritter.Write("<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\" /> <selection pane=\"bottomLeft\" /></sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"15\"/>");
                    }
                    else
                    {
                        sheetWritter.Write("<sheetViews><sheetView workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\" /> <selection pane=\"bottomLeft\" /></sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"15\"/>");
                    }
                }
                else
                {
                    if (sheetCnt == 0)
                    {
                        sheetWritter.Write("<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews><sheetFormatPr defaultRowHeight=\"15\"/>");
                    }
                    else
                    {
                        sheetWritter.Write("<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><sheetFormatPr defaultRowHeight=\"15\"/>");
                    }
                }

                List<double> colWidth2 = _colWidthsArray.ToArray().ToList();//??
                List<double> colWidth3 = colWidth2.FindAll(x => x != 1.0);

                if (colWidth3.Count > 0)
                {
                    sheetWritter.Write("<cols>");
                    int l = 1;
                    foreach (var item in colWidth2)
                    {
                        if (item == 1.0)
                        {
                            l++;
                            continue;
                        }
                        sheetWritter.Write(String.Format(CultureInfo.InvariantCulture.NumberFormat, "<col min=\"{0}\" max=\"{0}\" width=\"{1}\" bestFit = \"1\" customWidth=\"1\" />", l + startingColumn, Math.Min(item, _MAX_WIDTH)));
                        l++;
                    }
                    sheetWritter.Write("</cols>");
                }
                sheetWritter.Write("<sheetData>");
            }

            return rowNum;


        }
        public override void WriteSheet(string[] oneColumn)
        {
            sheetCnt++;
            SetStyles(headerStyle: null, styles: null);
            string[,] dane = new string[oneColumn.Length, 1];

            for (int i = 0; i < oneColumn.Length; i++)
            {
                dane[i, 0] = oneColumn[i];
            }

            _dataColReader = new DataColReader(dane);
            if (_inMemoryMode)
            {
                var e1 = _excelArchiveFile.CreateEntry(_sheetList[sheetCnt].pathInArchive, _clvl);
                using StreamWriter writter = new FormattingStreamWriter(e1.Open(), Encoding.UTF8, bufferSize: _bufferSize, CultureInfo.InvariantCulture.NumberFormat);
                WriteSheet(writter, 0, 0);
            }
            else
            {
                using StreamWriter writter = new FormattingStreamWriter(_sheetList[sheetCnt].pathOnDisc, false, Encoding.UTF8, bufferSize: _bufferSize, CultureInfo.InvariantCulture.NumberFormat);
                WriteSheet(writter, 0, 0);
            }
        }
        /// <summary>Writes the opening of a cell tag, with the optional style attribute.</summary>
        private void WriteCellStart(bool hasGap, int column, int rowNumber, int styleXf)
        {
            if (hasGap || _currentStartingColumn != 0)
            {
                WriteStringToBuffer("<c r=\"");
                WriteStringToBuffer(_letters[column + _currentStartingColumn]);
                WriteInt32ToBuffer(rowNumber);
                WriteStringToBuffer("\"");
            }
            else
            {
                WriteStringToBuffer("<c");
            }
            if (styleXf >= 0)
            {
                WriteStringToBuffer(" s=\"");
                WriteInt32ToBuffer(styleXf);
                WriteStringToBuffer("\"");
            }
        }

        private void WriteBlankCell(int column, int rowNumber, int styleXf)
        {
            WriteCellStart(hasGap: true, column, rowNumber, styleXf);
            WriteStringToBuffer("/>");
        }

        private void FlushPendingBlankCells(ref int pendingNullStart, ref int lastWrittenColumn, int currentColumn, int rowNumber)
        {
            if (pendingNullStart == -1)
            {
                return;
            }
            for (int blankColumn = pendingNullStart; blankColumn < currentColumn; blankColumn++)
            {
                WriteBlankCell(blankColumn, rowNumber, -1);
            }
            lastWrittenColumn = currentColumn - 1;
            pendingNullStart = -1;
        }

        private void WriteRow(int columnCount, int rowNumber, Dictionary<int, CellStyle>? rowStyles, bool boldedStyle = false)
        {
            int lastWrittenColumn = -1;
            int pendingNullStart = -1;
            for (int column = 0; column < columnCount; column++)
            {
                CellStyle? cellStyle = null;
                rowStyles?.TryGetValue(column, out cellStyle);

                if (_dataColReader.IsDBNull(column))
                {
                    if (cellStyle is not null)
                    {
                        FlushPendingBlankCells(ref pendingNullStart, ref lastWrittenColumn, column, rowNumber);
                        WriteBlankCell(column, rowNumber, RegisterStyle(cellStyle));
                        lastWrittenColumn = column;
                    }
                    else if (pendingNullStart == -1)
                    {
                        pendingNullStart = column;
                    }
                    continue;
                }

                if (pendingNullStart != -1)
                {
                    FlushPendingBlankCells(ref pendingNullStart, ref lastWrittenColumn, column, rowNumber);
                }

                bool hasGap = column > lastWrittenColumn + 1;

                object rawVal = _dataColReader.GetValue(column);
                if (rawVal is FormattedCell fc)
                {
                    object innerVal = fc.Value;
                    int fmtXfIndex = cellStyle is not null
                        ? RegisterStyle(CellStyle.Merge(cellStyle, new CellStyle { NumberFormat = fc.Format }))
                        : RegisterFormat(fc.Format);
                    WriteFormattedCell(innerVal, fmtXfIndex, column, rowNumber, hasGap);
                    lastWrittenColumn = column;
                    continue;
                }

                if (_newTypes[column] == TypeCode.String || _newTypes[column] == TypeCode.Object || typesArray[column] == 5) // string
                {
                    string? stringValue = null;
                    if (_newTypes[column] == TypeCode.String)
                    {
                        stringValue = _dataColReader.GetString(column);
                    }
                    else if (typesArray[column] == 5)
                    {
                        stringValue = Encoding.UTF8.GetString(((Memory<byte>)_dataColReader.GetValue(column)).Span);
                    }
                    else
                    {
                        stringValue = _dataColReader.GetValue(column).ToString()!;
                    }

                    if (!_inMemoryMode)
                    {
                        if (_colWidthsArray[column] < stringValue.Length * 1.25 + 2.0)
                        {
                            _colWidthsArray[column] = stringValue.Length * 1.25 + 2.0;
                        }
                    }

                    int styleXf = cellStyle is not null ? RegisterStyle(cellStyle) : (boldedStyle ? XfStyleBoldHeader : -1);

                    if (!_useScharedStrings)
                    {
                        WriteCellStart(hasGap, column, rowNumber, styleXf);
                        if (ShouldPreserveWhitespace(stringValue))
                        {
                            WriteStringToBuffer(" t=\"inlineStr\"><is><t xml:space=\"preserve\">");
                        }
                        else
                        {
                            WriteStringToBuffer(" t=\"inlineStr\"><is><t>");
                        }
                        WriteXmlEscapedToBuffer(stringValue);
                        WriteStringToBuffer("</t></is></c>");
                    }
                    else
                    {
                        ref var dicRefValue = ref CollectionsMarshal.GetValueRefOrAddDefault(_sstDic, stringValue, out bool exists);
                        if (!exists)
                        {
                            dicRefValue = _sstCntUnique;
                            _sstCntUnique++;
                        }
                        WriteCellStart(hasGap, column, rowNumber, styleXf);
                        WriteStringToBuffer(" t=\"s\"><v>");
                        WriteInt32ToBuffer(dicRefValue);
                        WriteStringToBuffer("</v></c>");
                        _sstCntAll++;
                    }
                    lastWrittenColumn = column;
                }
                else if (typesArray[column] == 1)//number
                {
                    WriteCellStart(hasGap, column, rowNumber, cellStyle is not null ? RegisterStyle(cellStyle) : -1);
                    WriteStringToBuffer("><v>");

                    switch (_newTypes[column])
                    {
                        case TypeCode.Byte:
                            byte byteValue = _dataColReader.GetByte(column);
                            WriteByteToBuffer(byteValue);
                            break;
                        case TypeCode.SByte:
                            sbyte sbyteValue = _dataColReader.GetSByte(column);
                            WritesByteToBuffer(sbyteValue);
                            break;
                        case TypeCode.Int16:
                            Int16 int16Value = _dataColReader.GetInt16(column);
                            WriteInt16ToBuffer(int16Value);
                            break;
                        case TypeCode.Int32:
                            Int32 int32Value = _dataColReader.GetInt32(column);
                            WriteInt32ToBuffer(int32Value);
                            break;
                        case TypeCode.Int64:
                            Int64 int64Value = _dataColReader.GetInt64(column);
                            WriteInt64ToBuffer(int64Value);
                            break;
                        case TypeCode.Single:
                            float doubleValue = _dataColReader.GetFloat(column);
                            WriteFloatToBuffer(doubleValue);
                            break;
                        case TypeCode.Double:
                            double doubleVal = _dataColReader.GetDouble(column);
                            WriteDoubleToBuffer(doubleVal);
                            break;
                        case TypeCode.Decimal:
                            decimal decimalVal = _dataColReader.GetDecimal(column);
                            WriteDoubleToBuffer(decimal.ToDouble(decimalVal));
                            break;
                        default:
                            throw new Exception("number format Excel error");
                    }


                    WriteStringToBuffer("</v></c>");
                    lastWrittenColumn = column;
                }
                else if (typesArray[column] == 2) //date
                {
                    DateTime dtVal = _dataColReader.GetDateTime(column);
                    int styleXf = cellStyle is not null ? RegisterStyle(cellStyle, NumFmtIdDate) : XfStyleDate;
                    WriteCellStart(hasGap, column, rowNumber, styleXf);
                    WriteStringToBuffer("><v>");

                    WriteDoubleToBuffer((double)(dtVal as DateTime?)?.ToOADate()!);
                    WriteStringToBuffer("</v></c>");
                    if (!_inMemoryMode)
                    {
                        _colWidthsArray[column] = _dateWidth;
                    }
                    lastWrittenColumn = column;
                }
                else if (typesArray[column] == 3) //datetime
                {
                    DateTime dtVal = _dataColReader.GetDateTime(column);
                    if (SuppressYear1000Dates && (dtVal as DateTime?).Value.Year == 1000)//1000-xx-xx
                    {
                        continue;
                    }
                    int styleXf = cellStyle is not null ? RegisterStyle(cellStyle, NumFmtIdDateTime) : XfStyleDateTime;
                    WriteCellStart(hasGap, column, rowNumber, styleXf);
                    WriteStringToBuffer("><v>");
                    WriteDoubleToBuffer((double)((dtVal) as DateTime?)?.ToOADate()!);
                    WriteStringToBuffer("</v></c>");
                    if (!_inMemoryMode)
                    {
                        _colWidthsArray[column] = _dateTimeWidth;
                    }
                    lastWrittenColumn = column;
                }
                else if (_newTypes[column] == TypeCode.Boolean)
                {
                    WriteCellStart(hasGap, column, rowNumber, cellStyle is not null ? RegisterStyle(cellStyle) : -1);
                    WriteStringToBuffer(" t=\"b\"><v>");
                    if (_dataColReader.GetBoolean(column))
                    {
                        _buffer[_currentBufferOffset++] = '1';
                    }
                    else
                    {
                        _buffer[_currentBufferOffset++] = '0';
                    }
                    WriteStringToBuffer("</v></c>");
                    lastWrittenColumn = column;
                }
            }
        }
        private void WriteFormattedCell(object val, int xfIndex, int column, int rowNumber, bool hasGap)
        {
            bool writeReference = hasGap || _currentStartingColumn != 0;
            WriteStringToBuffer(writeReference ? "<c r=\"" : "<c ");
            if (writeReference)
            {
                WriteStringToBuffer(_letters[column + _currentStartingColumn]);
                WriteInt32ToBuffer(rowNumber);
                WriteStringToBuffer("\"");
            }
            if (val is string strVal)
            {
                WriteStringToBuffer(" t=\"inlineStr\" s=\"");
                WriteInt32ToBuffer(xfIndex);
                WriteStringToBuffer("\"><is><t>");
                WriteXmlEscapedToBuffer(strVal);
                WriteStringToBuffer("</t></is></c>");
            }
            else if (val is bool bVal)
            {
                WriteStringToBuffer(" t=\"b\" s=\"");
                WriteInt32ToBuffer(xfIndex);
                WriteStringToBuffer("\"><v>");
                _buffer[_currentBufferOffset++] = bVal ? '1' : '0';
                WriteStringToBuffer("</v></c>");
            }
            else if (val is DateTime dtVal)
            {
                WriteStringToBuffer(" s=\"");
                WriteInt32ToBuffer(xfIndex);
                WriteStringToBuffer("\"><v>");
                WriteDoubleToBuffer(dtVal.ToOADate());
                WriteStringToBuffer("</v></c>");
            }
            else
            {
                WriteStringToBuffer(" s=\"");
                WriteInt32ToBuffer(xfIndex);
                WriteStringToBuffer("\"><v>");
                WriteDoubleToBuffer(Convert.ToDouble(val));
                WriteStringToBuffer("</v></c>");
            }
        }

        private const string _filterDefinedName = "_xlnm._FilterDatabase";
        internal override void FinalizeFile()
        {
            var e1 = _excelArchiveFile.CreateEntry("[Content_Types].xml", _clvl);

            using (var writer = new FormattingStreamWriter(e1.Open(), CultureInfo.InvariantCulture.NumberFormat))
            {
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types ");
                writer.Write("xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
                writer.Write("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
                writer.Write("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
                if (!String.IsNullOrWhiteSpace(DocPropertyProgramName))
                {
                    writer.Write("<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
                    writer.Write("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
                }
                writer.Write("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
                //if (UseDocPropsAndTheme)
                //{
                //    writer.Write("<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>");
                //}
                writer.Write("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");

                foreach (var (_, _, _, _, nameInArchive, _, _) in _sheetList)
                {
                    writer.Write($"<Override PartName=\"/xl/worksheets/{nameInArchive}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
                }
                if (_useScharedStrings)
                {
                    writer.Write("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");
                }
                writer.Write("</Types>");
            }

            if (!String.IsNullOrWhiteSpace(DocPropertyProgramName))
            {
                var e2 = _excelArchiveFile.CreateEntry("docProps/app.xml", _clvl);
                using var writer = new FormattingStreamWriter(e2.Open(), CultureInfo.InvariantCulture.NumberFormat);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Properties ");
                writer.Write("xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" ");
                writer.Write("xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><Application");
                writer.Write($">{DocPropertyProgramName}</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs");
                writer.Write("><vt:vector size=\"2\" baseType=\"variant\"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt");
                writer.Write($":variant><vt:i4>{_sheetList.Count}</vt:i4></vt:variant></vt:vector></HeadingPairs>");
                writer.Write("<TitlesOfParts>");
                writer.Write($"<vt:vector size=\"{_sheetList.Count}\" baseType=\"lpstr\">");

                foreach (var (name, _, _, _, _, _,_) in _sheetList)
                {
                    writer.Write($"<vt:lpstr>{name}</vt:lpstr>");
                }
                writer.Write("</vt:vector>");
                writer.Write("</TitlesOfParts>");
                writer.Write("<Company></Company><LinksUpToDate>false</LinksUpToDate");
                writer.Write("><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>12.0000");
                writer.Write("</AppVersion></Properties>");
            }

            var e3 = _excelArchiveFile.CreateEntry("xl/workbook.xml", _clvl);
            using (var writer = new FormattingStreamWriter(e3.Open(), CultureInfo.InvariantCulture.NumberFormat))
            {

                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook ");
                writer.Write("xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ");
                writer.Write("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><fileVersion ");
                writer.Write("appName=\"xl\" lastEdited=\"4\" lowestEdited=\"4\" rupBuild=\"4505\"/><workbookPr ");
                writer.Write("defaultThemeVersion=\"124226\"/><bookViews><workbookView xWindow=\"240\" yWindow=\"15\" ");
                writer.Write("windowWidth=\"16095\" windowHeight=\"9660\"/></bookViews>");
                writer.Write("<sheets>");
                for (int i = 0; i < _sheetList.Count; i++)
                {
                    var a = _sheetList[i].isHidden ? " state =\"hidden\"" : "";
                    writer.Write($"<sheet name=\"{_sheetList[i].name}\" sheetId=\"{i + 1}\"{a} r:id=\"rId{i + 1}\"/>");
                }
                writer.Write("</sheets>");

                if (_autofilterIsOn)
                {
                    writer.Write("<definedNames>");
                    foreach (var item in this._sheetList)
                    {
                        if (!string.IsNullOrWhiteSpace(item.filterHeaderRange))
                        {
                            int localSheetId = item.sheetId - 1;
                            string filterHeaderRange = item.filterHeaderRange;
                            writer.Write($"<definedName name=\"{_filterDefinedName}\" localSheetId=\"{localSheetId}\" hidden=\"1\">{filterHeaderRange}</definedName>");
                        }
                    }
                    writer.Write("</definedNames>");
                }

                writer.Write("<calcPr calcId=\"124519\" fullCalcOnLoad=\"1\"/></workbook>");
            }

            var e4 = _excelArchiveFile.CreateEntry("xl/_rels/workbook.xml.rels", _clvl);
            using (var writer = new FormattingStreamWriter(e4.Open(), CultureInfo.InvariantCulture.NumberFormat))
            {
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                for (int i = 0; i < _sheetList.Count; i++)
                {
                    writer.Write($"<Relationship Id=\"rId{i + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/{_sheetList[i].nameInArchive}.xml\"/>");
                }
                writer.Write($"<Relationship Id=\"rId{_sheetList.Count + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
                if (_useScharedStrings)
                {
                    writer.Write($"<Relationship Id=\"rId{_sheetList.Count + 2}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>");
                }
                writer.Write($"</Relationships>");
            }

            var e5 = _excelArchiveFile.CreateEntry("_rels/.rels", _clvl);
            using (var writer = new FormattingStreamWriter(e5.Open(), CultureInfo.InvariantCulture.NumberFormat))
            {
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                writer.Write("<Relationships ");
                writer.Write("xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                writer.Write("<Relationship Id=\"rId1\" ");
                writer.Write("Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" ");
                writer.Write("Target=\"xl/workbook.xml\"/>");

                if (!String.IsNullOrWhiteSpace(DocPropertyProgramName))
                {
                    writer.Write("<Relationship Id=\"rId2\" ");
                    writer.Write("Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" ");
                    writer.Write("Target=\"docProps/core.xml\"/>");
                    writer.Write("<Relationship Id=\"rId3\" ");
                    writer.Write("Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" ");
                    writer.Write("Target=\"docProps/app.xml\"/>");
                }
                writer.Write("</Relationships>");
            }

            if (!String.IsNullOrWhiteSpace(DocPropertyProgramName))
            {
                var e6 = _excelArchiveFile.CreateEntry("docProps/core.xml", _clvl);
                string stringNow = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssZ");

                using var writer = new FormattingStreamWriter(e6.Open(), CultureInfo.InvariantCulture.NumberFormat);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><cp:coreProperties ");
                writer.Write("xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" ");
                writer.Write("xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" ");
                writer.Write("xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" ");
                writer.Write($"xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dc:creator>{DocPropertyProgramName}</dc:creator>");
                writer.Write($"<cp:lastModifiedBy>{DocPropertyProgramName}</cp:lastModifiedBy>");
                writer.Write($"<dcterms:created xsi:type=\"dcterms:W3CDTF\">{stringNow}</dcterms:created><dcterms:modified ");
                writer.Write($"xsi:type=\"dcterms:W3CDTF\">{stringNow}</dcterms:modified></cp:coreProperties>");
            }

            var e7 = _excelArchiveFile.CreateEntry("xl/styles.xml", _clvl);
            using (var writer = new FormattingStreamWriter(e7.Open(), CultureInfo.InvariantCulture.NumberFormat))
            {
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet ");
                writer.Write("xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");

                if (_formatRegistry.Count > 0)
                {
                    writer.Write($"<numFmts count=\"{_formatRegistry.Count}\">");
                    foreach (var kvp in _formatRegistry)
                    {
                        string escaped = kvp.Key
                            .Replace("&", "&amp;")
                            .Replace("\"", "&quot;")
                            .Replace("<", "&lt;")
                            .Replace(">", "&gt;");
                        writer.Write($"<numFmt numFmtId=\"{kvp.Value}\" formatCode=\"{escaped}\"/>");
                    }
                    writer.Write("</numFmts>");
                }

                // ── Fonts: the two defaults plus deduplicated custom fonts ──
                var fonts = new List<string>
                {
                    "<font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font>",
                    "<font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/><b/></font>",
                };
                foreach (var kvp in _fontRegistry.OrderBy(kvp => kvp.Value))
                {
                    if (kvp.Value < 2)
                    {
                        continue;
                    }
                    FontKey key = kvp.Key;
                    var font = new StringBuilder("<font><sz val=\"11\"/>");
                    font.Append(key.Color is null ? "<color theme=\"1\"/>" : $"<color rgb=\"FF{key.Color}\"/>");
                    font.Append("<name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/>");
                    if (key.Bold == true) font.Append("<b/>");
                    if (key.Italic == true) font.Append("<i/>");
                    if (key.Underline == true) font.Append("<u/>");
                    font.Append("</font>");
                    fonts.Add(font.ToString());
                }

                // ── Fills: none, gray125 plus deduplicated solid fills ──
                var fills = new List<string>
                {
                    "<fill><patternFill patternType=\"none\"/></fill>",
                    "<fill><patternFill patternType=\"gray125\"/></fill>",
                };
                foreach (string fillColor in _fillRegistry.OrderBy(kvp => kvp.Value).Select(kvp => kvp.Key))
                {
                    fills.Add($"<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF{fillColor}\"/><bgColor indexed=\"64\"/></patternFill></fill>");
                }
                writer.Write($"<fonts count=\"{fonts.Count}\">");
                foreach (string font in fonts)
                {
                    writer.Write(font);
                }
                writer.Write($"</fonts><fills count=\"{fills.Count}\">");
                foreach (string fill in fills)
                {
                    writer.Write(fill);
                }
                writer.Write("</fills><borders ");
                writer.Write("count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs ");
                writer.Write("count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");

                // Format XFs and style XFs share one index space, so they must be emitted
                // in index order: the position in cellXfs is the value a cell references
                // through its s= attribute.
                var customXfs = new SortedDictionary<int, string>();
                foreach (var kvp in _formatXfMap)
                {
                    int numFmtId = _formatRegistry[kvp.Key];
                    customXfs[kvp.Value] = $"<xf numFmtId=\"{numFmtId}\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>";
                }
                foreach (var kvp in _styleComponents)
                {
                    StyleComponent component = kvp.Value;
                    CellStyle style = component.Style;
                    var xf = new StringBuilder("<xf ");
                    xf.Append($"numFmtId=\"{component.NumFmtId}\" fontId=\"{component.FontId}\" fillId=\"{component.FillId}\" borderId=\"0\" xfId=\"0\"");
                    if (component.NumFmtId != 0) xf.Append(" applyNumberFormat=\"1\"");
                    if (component.FontId != 0) xf.Append(" applyFont=\"1\"");
                    if (component.FillId != 0) xf.Append(" applyFill=\"1\"");
                    List<string>? alignment = null;
                    if (style.HasAlignment)
                    {
                        alignment = new List<string>(3);
                        if (style.Horizontal is not null) alignment.Add($"horizontal=\"{style.Horizontal}\"");
                        if (style.Vertical is not null) alignment.Add($"vertical=\"{style.Vertical}\"");
                        if (style.Wrap is not null) alignment.Add($"wrapText=\"{(style.Wrap.Value ? 1 : 0)}\"");
                        xf.Append(" applyAlignment=\"1\"");
                    }
                    xf.Append('>');
                    if (alignment is not null)
                    {
                        xf.Append($"<alignment {string.Join(" ", alignment)}/>");
                    }
                    xf.Append("</xf>");
                    customXfs[kvp.Key] = xf.ToString();
                }

                int totalXf = 4 + customXfs.Count;
                writer.Write($"<cellXfs count=\"{totalXf}\">");
                writer.Write("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>");
                writer.Write("<xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>");  // date
                writer.Write("<xf numFmtId=\"22\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>");  // datetime
                writer.Write("<xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>");  // bold header

                foreach (var kvp in customXfs)
                {
                    writer.Write(kvp.Value);
                }

                writer.Write("</cellXfs>");
                writer.Write("<cellStyles ");
                writer.Write("count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>");
                writer.Write("</cellStyles>");
                writer.Write("<dxfs ");
                writer.Write("count=\"0\"/><tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium9\" ");
                writer.Write("defaultPivotStyle=\"PivotStyleLight16\"/></styleSheet>");
            }

            if (_useScharedStrings)
            {
                var entry = _excelArchiveFile.CreateEntry("xl/sharedStrings.xml", _clvl);

                using var o = entry.Open();
                using var _sharedStringWritter = new FormattingStreamWriter(o, Encoding.UTF8, bufferSize: _bufferSize, CultureInfo.InvariantCulture.NumberFormat);
                _sharedStringWritter.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                _sharedStringWritter.Write("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
                _sharedStringWritter.Write($" count=\"{_sstCntAll}\" uniqueCount=\"{_sstDic.Count}\">");

                foreach (string dana in _sstDic.Keys)
                {
                    if (ShouldPreserveWhitespace(dana))
                    {
                        _sharedStringWritter.Write("<si><t xml:space=\"preserve\">");
                        _sharedStringWritter.Write(dana);
                        _sharedStringWritter.Write("</t></si>");
                    }
                    else
                    {
                        _sharedStringWritter.Write("<si><t>");
                        _sharedStringWritter.Write(dana);
                        _sharedStringWritter.Write("</t></si>");
                    }
                }
                _sharedStringWritter.Write("</sst>");
            }
        }

        private static bool ShouldPreserveWhitespace(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return false;
            }

            return char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1]);
        }

        public override void Dispose()
        {
            this.Save();
        }

        public static int WriteToExisting(StreamWriter sheetWritter, IDataReader reader, int dateStyleNum = 1, int dateTimeStyleNum = 2, int startingRow = 0, int startingColumn = 0)
        {
            sheetWritter.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sheetWritter.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            sheetWritter.Write("<sheetData>");

            int rowNum = 0;
            int colCnt = reader.FieldCount;
            int[] typesArray = new int[colCnt];

            for (int j = 0; j < colCnt; j++)
            {
                if (ExcelWriter._stringTypes.Contains(reader.GetFieldType(j)))
                {
                    typesArray[j] = 0;
                }
                else if (ExcelWriter._numberTypes.Contains(reader.GetFieldType(j)))
                {
                    typesArray[j] = 1;
                }
                else if (reader.GetFieldType(j) == typeof(System.DateTime) && reader.GetDataTypeName(j).Equals("Date", StringComparison.OrdinalIgnoreCase))
                {
                    typesArray[j] = 2;
                }
                else if (reader.GetFieldType(j) == typeof(System.DateTime)
                    && (reader.GetDataTypeName(j).Equals("timestamp", StringComparison.OrdinalIgnoreCase) || reader.GetDataTypeName(j).Equals("DateTime", StringComparison.OrdinalIgnoreCase)))
                {
                    typesArray[j] = 3;
                }
                else // Boolean, String, other -> as String
                {
                    typesArray[j] = -1;
                }
            }

            //sheetWritter.Write("<row r=\"");
            //sheetWritter.Write(rowNum + 1 + startingRow);
            //sheetWritter.Write("\">");
            sheetWritter.Write("<row>");

            for (int j = 0; j < colCnt; j++)
            {
                string stringValue = reader.GetName(j);
                sheetWritter.Write("<c r=\"");
                sheetWritter.Write(_letters[j + startingColumn]);
                sheetWritter.Write((rowNum + 1 + startingRow));
                sheetWritter.Write("\"");
                if (ShouldPreserveWhitespace(stringValue))
                {
                    sheetWritter.Write(" t=\"inlineStr\"><is><t xml:space=\"preserve\">");
                }
                else
                {
                    sheetWritter.Write(" t=\"inlineStr\"><is><t>");
                }
                sheetWritter.Write(stringValue);
                sheetWritter.Write("</t></is></c>");
            }
            sheetWritter.Write("</row>");
            rowNum++;

            while (reader.Read())
            {
                sheetWritter.Write("<row r=\"");
                sheetWritter.Write(rowNum + 1 + startingRow);
                sheetWritter.Write("\">");

                for (int j = 0; j < colCnt; j++)
                {
                    var rawValue = reader.GetValue(j);
                    if (rawValue == null || rawValue == DBNull.Value)
                        continue;

                    if (typesArray[j] == 0 || typesArray[j] == -1)
                    {
 
                        string stringValue = rawValue.ToString()!;

                        if (stringValue.Contains('&'))
                        {
                            stringValue = stringValue.Replace("&", "&amp;");
                        }
                        if (stringValue.Contains('<'))
                        {
                            stringValue = stringValue.Replace("<", "&lt;");
                        }
                        if (stringValue.Contains('>'))
                        {
                            stringValue = stringValue.Replace(">", "&gt;");
                        }
                        if (stringValue.Contains('\"'))
                        {
                            stringValue = stringValue.Replace("\"", "&quot;");
                        }
                        if (stringValue.Contains('\''))
                        {
                            stringValue = stringValue.Replace("'", "&apos;");
                        }

                        sheetWritter.Write("<c r=\"");
                        sheetWritter.Write(_letters[j + startingColumn]);
                        sheetWritter.Write((rowNum + 1 + startingRow));
                        if (ShouldPreserveWhitespace(stringValue))
                        {
                            sheetWritter.Write("\" t=\"inlineStr\"><is><t xml:space=\"preserve\">");
                        }
                        else
                        {
                            sheetWritter.Write("\" t=\"inlineStr\"><is><t>");
                        }
                        sheetWritter.Write(stringValue);
                        sheetWritter.Write("</t></is></c>");
                    }
                    else if (typesArray[j] == 1)//number
                    {
                        sheetWritter.Write("<c r=\"");
                        sheetWritter.Write(_letters[j + startingColumn]);
                        sheetWritter.Write((rowNum + 1 + startingRow));
                        sheetWritter.Write("\"><v>");
                        if (rawValue is decimal decimalValue)
                        {
                            sheetWritter.Write(decimal.ToDouble(decimalValue));
                        }
                        else
                        {
                            sheetWritter.Write(rawValue);
                        }

                        sheetWritter.Write("</v></c>");
                    }
                    else if (typesArray[j] == 2) //date
                    {
                        sheetWritter.Write("<c r=\"");
                        sheetWritter.Write(_letters[j + startingColumn]);
                        sheetWritter.Write((rowNum + 1 + startingRow));
                        sheetWritter.Write($"\" s=\"{dateStyleNum}\"><v>");
                        sheetWritter.Write(((rawValue) as DateTime?)?.ToOADate());
                        sheetWritter.Write("</v></c>");

                    }
                    else if (typesArray[j] == 3) //datetime
                    {
                        if (rawValue is DateTime valDateTime && valDateTime.Year == 1000)//1000-xx-xx
                        {
                            continue;
                        }

                        sheetWritter.Write("<c r=\"");
                        sheetWritter.Write(_letters[j + startingColumn]);
                        sheetWritter.Write((rowNum + 1 + startingRow));
                        sheetWritter.Write($"\" s=\"{dateTimeStyleNum}\"><v>");
                        sheetWritter.Write(((rawValue) as DateTime?)?.ToOADate());
                        sheetWritter.Write("</v></c>");
                    }
                }

                rowNum++;
                sheetWritter.Write("</row>");
            }

            sheetWritter.Write("</sheetData>");
            sheetWritter.Write("<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/></worksheet>");
            return rowNum - 1;
        }

    }

    internal class FormattingStreamWriter : StreamWriter
    {
        private readonly IFormatProvider _formatProvider;

        public FormattingStreamWriter(string path, bool append, Encoding encoding, int bufferSize, IFormatProvider formatProvider)
            : base(path, append, encoding, bufferSize)
        {
            this._formatProvider = formatProvider;
        }

        public FormattingStreamWriter(string path, bool append, Encoding encoding, IFormatProvider formatProvider)
        : base(path, append, encoding)
        {
            this._formatProvider = formatProvider;
        }
        public FormattingStreamWriter(string path, IFormatProvider formatProvider)
            : base(path)
        {
            this._formatProvider = formatProvider;
        }
        public FormattingStreamWriter(Stream stream, IFormatProvider formatProvider)
        : base(stream)
        {
            this._formatProvider = formatProvider;
        }
        public FormattingStreamWriter(Stream stream, Encoding encoding, int bufferSize, IFormatProvider formatProvider)
        : base(stream, encoding, bufferSize)
        {
            this._formatProvider = formatProvider;
        }

        public FormattingStreamWriter(Stream stream, Encoding encoding, IFormatProvider formatProvider)
        : base(stream, encoding)
        {
            this._formatProvider = formatProvider;
        }


        public override IFormatProvider FormatProvider
        {
            get
            {
                return this._formatProvider;
            }
        }
    }

    /// <summary>
    /// cellXfs
    /// </summary>
    internal class StyleInfo
    {
        public int NumFmtId;
        public int XfId;
        //public int ApplyNumberFormat;
    }

    //internal class NumberFormatInfo
    //{
    //    public string Name;
    //    public Type proposedType;
    //}

}
