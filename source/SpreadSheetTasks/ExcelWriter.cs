using SpreadSheetTasks.Helpers;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace SpreadSheetTasks
{
    public abstract class ExcelWriter : IDisposable
    {
        internal static readonly HashSet<string> _stringDb = ["nvarchar", "varchar", "char"];
        internal static readonly HashSet<Type> _stringTypes = [typeof(String), typeof(Char), typeof(Boolean)];

        internal Dictionary<string, int> _formatRegistry = new();
        internal Dictionary<string, int> _formatXfMap = new();
        internal int _nextNumFmtId = 165;
        internal int _nextXfIndex = 4;

        /// <summary>Custom font definitions, keyed by the properties that change the font. Id 1 is the built-in bold font.</summary>
        internal Dictionary<FontKey, int> _fontRegistry = new() { [new FontKey(null, true, null, null)] = 1 };
        /// <summary>Custom solid fill colors; ids start at 2 because 0 (none) and 1 (gray125) are built in.</summary>
        internal Dictionary<string, int> _fillRegistry = new() { };
        private int _nextFontId = 2;
        private int _nextFillId = 2;

        /// <summary>Cell style definitions, keyed by style and an optional built-in number-format id.</summary>
        internal Dictionary<(CellStyle Style, int NumFmtIdOverride), int> _styleXfMap = new();
        /// <summary>Components of every style XF, keyed by the XF index.</summary>
        internal Dictionary<int, StyleComponent> _styleComponents = new();

        internal List<StyleRange>? _styleRanges;
        internal CellStyle? _headerStyle;
        /// <summary>Reused by <see cref="GetRowStyles"/>; only valid until the next call.</summary>
        private readonly Dictionary<int, CellStyle> _rowStylesScratch = new();

        internal bool _hasCustomFormats => _formatRegistry.Count > 0 || _styleComponents.Count > 0;

        internal int RegisterFormat(string formatString)
        {
            if (_formatXfMap.TryGetValue(formatString, out int xfIndex))
                return xfIndex;
            // A style may already have registered this format; reuse its id so both
            // XFs keep referencing a number format that is actually emitted.
            if (!_formatRegistry.TryGetValue(formatString, out int numFmtId))
            {
                numFmtId = _nextNumFmtId++;
                _formatRegistry[formatString] = numFmtId;
            }
            xfIndex = _nextXfIndex++;
            _formatXfMap[formatString] = xfIndex;
            return xfIndex;
        }

        /// <summary>
        /// Registers a cell style and returns its cellXfs index. Fonts, fills and number
        /// formats are deduplicated, so cells sharing a style share one Excel style record.
        /// </summary>
        /// <param name="style">The style to register.</param>
        /// <param name="numFmtIdOverride">
        /// Built-in number-format id (for example 14 for a date) used when the style does not
        /// set <see cref="CellStyle.NumberFormat"/> itself, so a styled date keeps looking like a date.
        /// </param>
        internal int RegisterStyle(CellStyle style, int numFmtIdOverride = 0)
        {
            ArgumentNullException.ThrowIfNull(style);
            int effectiveOverride = style.NumberFormat is null ? numFmtIdOverride : 0;
            var key = (style, effectiveOverride);
            if (_styleXfMap.TryGetValue(key, out int xfIndex))
                return xfIndex;

            int fontId = 0;
            FontKey fontKey = style.ToFontKey();
            if (fontKey != new FontKey(null, null, null, null))
            {
                if (!_fontRegistry.TryGetValue(fontKey, out fontId))
                {
                    fontId = _nextFontId++;
                    _fontRegistry[fontKey] = fontId;
                }
            }

            int fillId = 0;
            if (style.Fill is not null && !_fillRegistry.TryGetValue(style.Fill, out fillId))
            {
                fillId = _nextFillId++;
                _fillRegistry[style.Fill] = fillId;
            }

            int numFmtId = 0;
            if (style.NumberFormat is not null)
            {
                if (!_formatRegistry.TryGetValue(style.NumberFormat, out numFmtId))
                {
                    numFmtId = _nextNumFmtId++;
                    _formatRegistry[style.NumberFormat] = numFmtId;
                }
            }
            else if (numFmtIdOverride != 0)
            {
                numFmtId = numFmtIdOverride;
            }

            xfIndex = _nextXfIndex++;
            _styleXfMap[key] = xfIndex;
            _styleComponents[xfIndex] = new StyleComponent(fontId, fillId, numFmtId, style);
            return xfIndex;
        }

        /// <summary>Stores the per-sheet header style and range styles selected by the caller.</summary>
        internal void SetStyles(CellStyle? headerStyle, IReadOnlyDictionary<string, CellStyle>? styles)
        {
            List<StyleRange>? parsedRanges = A1Reference.ParseStyleRanges(styles);
            _headerStyle = headerStyle;
            _styleRanges = parsedRanges;
        }

        /// <summary>
        /// Throws when any configured style needs alignment/wrapping, which is not yet
        /// supported by the XLSB writer.
        /// </summary>
        internal void EnsureNoAlignment()
        {
            if (_headerStyle is not null && _headerStyle.HasAlignment)
            {
                ThrowAlignmentNotSupported();
            }
            if (_styleRanges is not null)
            {
                foreach (StyleRange range in _styleRanges)
                {
                    if (range.Style.HasAlignment)
                    {
                        ThrowAlignmentNotSupported();
                    }
                }
            }

            static void ThrowAlignmentNotSupported() => throw new NotSupportedException(
                "horizontal/vertical alignment and wrapping are not yet supported for XLSB output.");
        }

        /// <summary>
        /// Returns the sparse column-to-style map for one written row, or null when the row
        /// has no styles. Range keys are absolute worksheet coordinates; the returned map is
        /// keyed by column index relative to <paramref name="startingColumn"/>. The returned
        /// dictionary is only valid until the next call.
        /// </summary>
        internal Dictionary<int, CellStyle>? GetRowStyles(int rowNum, int startingRow, int startingColumn, int columnCount)
        {
            if (_headerStyle is null && _styleRanges is null)
            {
                return null;
            }

            _rowStylesScratch.Clear();
            bool anyStyle = false;
            if (rowNum == 0 && _headerStyle is not null)
            {
                CellStyle effectiveHeader = CellStyle.Merge(new CellStyle { Bold = true }, _headerStyle);
                for (int column = 0; column < columnCount; column++)
                {
                    _rowStylesScratch[column] = effectiveHeader;
                }
                anyStyle = true;
            }

            if (_styleRanges is not null)
            {
                // In a headed write a range style on the first row is merged on top of the
                // same bold default that header strings receive, so styling a header cell
                // does not silently drop the bold.
                bool defaultBoldHeaderRow = rowNum == 0 && _areHeaders && _headerStyle is null;
                int absoluteRow = rowNum + startingRow;
                foreach (StyleRange range in _styleRanges)
                {
                    if (range.FirstRow > absoluteRow || absoluteRow > range.LastRow)
                    {
                        continue;
                    }
                    int firstAbsoluteColumn = Math.Max(range.FirstColumn, startingColumn);
                    int lastAbsoluteColumn = Math.Min(range.LastColumn, startingColumn + columnCount - 1);
                    if (firstAbsoluteColumn > lastAbsoluteColumn)
                    {
                        continue;
                    }
                    for (int absoluteColumn = firstAbsoluteColumn; absoluteColumn <= lastAbsoluteColumn; absoluteColumn++)
                    {
                        int column = absoluteColumn - startingColumn;
                        _rowStylesScratch.TryGetValue(column, out CellStyle? existing);
                        if (existing is null && defaultBoldHeaderRow)
                        {
                            existing = new CellStyle { Bold = true };
                        }
                        _rowStylesScratch[column] = CellStyle.Merge(existing, range.Style);
                        anyStyle = true;
                    }
                }
            }

            return anyStyle ? _rowStylesScratch : null;
        }

        internal static readonly HashSet<Type> _numberTypes =
        [
            typeof(sbyte), typeof(byte)
            , typeof(Int16), typeof(UInt16)
            , typeof(Int32), typeof(UInt32)
            , typeof(Int64), typeof(UInt64)
            , typeof(Single), typeof(Double)
            , typeof(Decimal)
        ];

        private static readonly HashSet<string> _DbNumbers =
        [
            "integer", "bigint"
            , "numeric", "Decimal"
            , "Double", "Single"
            , "Sbyte", "Byte"
            , "Int16", "Int32"
            , "Int64", "UInt16"
            , "UInt32", "UInt64"
        ];

        /// <summary>
        /// Application name written to the document's extended properties (docProps/app.xml) and
        /// core properties (docProps/core.xml). When set, the output file will include
        /// <c>docProps/app.xml</c> and <c>docProps/core.xml</c> parts referencing this name.
        /// Leave null/empty to omit the docProps parts entirely.
        /// </summary>
        public string? DocPropertyProgramName { get; set; }

        /// <summary>
        /// Obsolete. Use <see cref="DocPropertyProgramName"/> instead.
        /// </summary>
        [Obsolete("Use DocPropertyProgramName instead. This misspelled member will be removed in a future release.")]
        public string? DocPopertyProgramName
        {
            get => DocPropertyProgramName;
            set => DocPropertyProgramName = value;
        }

        internal static void SetTypes(DataColReader _dataColReader, int[] typesArray, TypeCode[] newTypes, int ColumnCount, bool detectBooleanType = false)
        {
            if (_dataColReader._dataReader != null)
            {
                var rdr = _dataColReader._dataReader;
                for (int j = 0; j < ColumnCount; j++)
                {
                    var tempType = rdr.GetFieldType(j);
                    newTypes[j] = Type.GetTypeCode(tempType);
                    if (detectBooleanType && tempType == typeof(Boolean))
                    {
                        typesArray[j] = 4;
                    }
                    else if (_stringTypes.Contains(tempType))
                    {
                        typesArray[j] = 0;
                    }
                    else if (_numberTypes.Contains(tempType))
                    {
                        typesArray[j] = 1;
                    }
                    else if (tempType == typeof(System.DateTime) && _dataColReader._databaseTypes[j].EndsWith("Date", StringComparison.OrdinalIgnoreCase))
                    {
                        typesArray[j] = 2;
                    }
                    else if (tempType == typeof(System.DateTime)
                        && (_dataColReader._databaseTypes[j].Equals("timestamp", StringComparison.OrdinalIgnoreCase) ||
                        _dataColReader._databaseTypes[j].EndsWith("DateTime", StringComparison.OrdinalIgnoreCase) ||
                         _dataColReader._databaseTypes[j].EndsWith("abstime", StringComparison.OrdinalIgnoreCase)
                        ))
                    {
                        typesArray[j] = 3;
                    }
                    else if (tempType == typeof(System.TimeSpan))
{
                        typesArray[j] = 3;
                    }
                    else if (tempType == typeof(Memory<byte>))
                    {
                        typesArray[j] = 5;
                    }
                    else // String, other -> as String
                    {
                        typesArray[j] = 0;
                        //throw new Exception("Excel type problem !");
                        //typesArray[j] = -1;
                    }
                }
            }
            else if (_dataColReader._dataTable != null)
            {
                var dt = _dataColReader._dataTable;
                for (int j = 0; j < ColumnCount; j++)
                {
                    newTypes[j] = Type.GetTypeCode(dt.Columns[j].DataType);
                    if (detectBooleanType && dt.Columns[j].DataType == typeof(Boolean))
                    {
                        typesArray[j] = 4;
                    }
                    else if (_stringTypes.Contains(dt.Columns[j].DataType))
                    {
                        typesArray[j] = 0;
                    }
                    else if (_numberTypes.Contains(dt.Columns[j].DataType))
                    {
                        typesArray[j] = 1;
                    }
                    else if (dt.Columns[j].DataType == typeof(System.DateTime))
                    {
                        typesArray[j] = 3;
                    }
                    else if (dt.Columns[j].DataType == typeof(System.TimeSpan))
{
                        typesArray[j] = 3;
                    }
                    else if (dt.Columns[j].DataType == typeof(Memory<byte>))
                    {
                        typesArray[j] = 5;
                    }
                    else // Boolean, String, other -> as String
                    {
                        typesArray[j] = 0;
                        //throw new Exception("Excel type problem !");
                        //typesArray[j] = -1;
                    }
                }
            }
            else
            {
                for (int j = 0; j < ColumnCount; j++)
                {
                    newTypes[j] = Type.GetTypeCode(_dataColReader.GetValue(j).GetType());
                    if (detectBooleanType && _dataColReader.GetValue(j).GetType() == typeof(Boolean))
                    {
                        typesArray[j] = 4;
                    }
                    else if (_stringTypes.Contains(_dataColReader.GetValue(j).GetType()) || _stringDb.Contains(_dataColReader._databaseTypes[j]))
                    {
                        typesArray[j] = 0;
                    }
                    else if (_numberTypes.Contains(_dataColReader.GetValue(j).GetType()) || _DbNumbers.Contains(_dataColReader._databaseTypes[j]))
                    {
                        typesArray[j] = 1;
                    }
                    else if (_dataColReader._databaseTypes[j].Equals("Date", StringComparison.OrdinalIgnoreCase))
                    {
                        typesArray[j] = 2;
                    }
                    else if (_dataColReader.GetValue(j).GetType() == typeof(System.DateTime) || _dataColReader._databaseTypes[j] == "timestamp" || _dataColReader._databaseTypes[j] == "DateTime" /*|| kolekcjaDanych.typyBazy[j] == "TimeSpan"*/)
                    {
                        typesArray[j] = 3;
                    }
                    else if (_dataColReader.GetValue(j).GetType() == typeof(System.TimeSpan))
                    {
                        typesArray[j] = 3;
                    }
                    else if (_dataColReader.GetValue(j).GetType() == typeof(Memory<byte>))
                    {
                        typesArray[j] = 5;
                    }
                    else // other
                    {
                        throw new Exception("Excel type problem !");
                        //typesArray[j] = -1;
                    }
                }
            }
        }

        internal Stream _newExcelFileStream;
        internal ExcelZipWriter _excelArchiveFile;
        internal List<(string name, string pathInArchive, string pathOnDisc, bool isHidden, string nameInArchive, int sheetId, string filterHeaderRange)> _sheetList = new();

        internal const int _MAX_WIDTH = 80;
        internal int _sstCntUnique = 0;
        internal int _sstCntAll = 0;
        internal int sheetCnt = -1;
        internal DataColReader _dataColReader;
        internal bool _areHeaders = false;
        internal Dictionary<string, int> _sstDic = [];

        internal double[] _colWidthsArray;
        internal int[]? typesArray;
        internal TypeCode[] _newTypes;

        protected int _rowsCount;
        public int RowsCount { get => _rowsCount; }

        protected bool _autofilterIsOn = false;

        internal abstract void FinalizeFile();
        public abstract void AddSheet(string sheetName, bool hidden = false);

        /// <summary>
        /// Writes the rows of an <see cref="IDataReader"/> to the current sheet.
        /// </summary>
        /// <param name="dataReader">Source rows.</param>
        /// <param name="headers">If true, the first row is treated as a header row (bold).</param>
        /// <param name="maxRows">Maximum number of columns to write; -1 writes all.</param>
        /// <param name="startingRow">Zero-based row index where writing begins.</param>
        /// <param name="startingColumn">Zero-based column index where writing begins.</param>
        /// <param name="doAutofilter">If true, add an autofilter on the written range.</param>
        public abstract void WriteSheet(IDataReader dataReader, Boolean headers = true, int maxRows = -1, int startingRow = 0, int startingColumn = 0, bool doAutofilter = false);

        /// <summary>
        /// Writes rows with a style for the first written row and/or A1 cell ranges.
        /// The header style is merged with the default bold header style; set
        /// <see cref="CellStyle.Bold"/> to false to disable bold.
        /// </summary>
        /// <param name="dataReader">Source rows.</param>
        /// <param name="headerStyle">Optional style applied to the first written row.</param>
        /// <param name="styles">Optional mapping of absolute worksheet A1 cell/range references to styles.</param>
        /// <param name="headers">If true, the first row is treated as a header row.</param>
        /// <param name="maxRows">Maximum number of columns to write; -1 writes all.</param>
        /// <param name="startingRow">Zero-based row index where writing begins.</param>
        /// <param name="startingColumn">Zero-based column index where writing begins.</param>
        /// <param name="doAutofilter">If true, add an autofilter on the written range.</param>
        public void WriteSheet(IDataReader dataReader, CellStyle? headerStyle, IReadOnlyDictionary<string, CellStyle>? styles = null,
            Boolean headers = true, int maxRows = -1, int startingRow = 0, int startingColumn = 0, bool doAutofilter = false)
        {
            SetStyles(headerStyle, styles);
            try
            {
                WriteSheet(dataReader, headers, maxRows, startingRow, startingColumn, doAutofilter);
            }
            finally
            {
                SetStyles(null, null);
            }
        }

        /// <summary>
        /// Obsolete. Use <see cref="WriteSheet(IDataReader, bool, int, int, int, bool)"/> with the
        /// <c>maxRows</c> parameter instead.
        /// </summary>
        [Obsolete("Use the maxRows parameter instead. This wrapper will be removed in a future release.")]
        public void WriteSheet(IDataReader dataReader, bool headers, int overLimit, int startingRow, int startingColumn, bool doAutofilter, bool _overLimitRenameMarker)
            => WriteSheet(dataReader, headers, overLimit, startingRow, startingColumn, doAutofilter);

        /// <summary>
        /// Writes the rows of a <see cref="DataTable"/> to the current sheet.
        /// </summary>
        public virtual void WriteSheet(DataTable dataTable, Boolean headers = true, int maxRows = -1, int startingRow = 0, int startingColumn = 0, bool doAutofilter = false)
        {
            WriteSheet(dataTable.CreateDataReader(), headers, maxRows, startingRow, startingColumn, doAutofilter);
        }

        /// <summary>Writes a table with a style for the first written row and/or A1 cell ranges.</summary>
        public void WriteSheet(DataTable dataTable, CellStyle? headerStyle, IReadOnlyDictionary<string, CellStyle>? styles = null,
            Boolean headers = true, int maxRows = -1, int startingRow = 0, int startingColumn = 0, bool doAutofilter = false)
        {
            SetStyles(headerStyle, styles);
            try
            {
                WriteSheet(dataTable, headers, maxRows, startingRow, startingColumn, doAutofilter);
            }
            finally
            {
                SetStyles(null, null);
            }
        }

        /// <summary>
        /// Obsolete. Use <see cref="WriteSheet(DataTable, bool, int, int, int, bool)"/> with the
        /// <c>maxRows</c> parameter instead.
        /// </summary>
        [Obsolete("Use the maxRows parameter instead. This wrapper will be removed in a future release.")]
        public void WriteSheet(DataTable dataTable, bool headers, int overLimit, int startingRow, int startingColumn, bool doAutofilter, bool _overLimitRenameMarker)
            => WriteSheet(dataTable, headers, overLimit, startingRow, startingColumn, doAutofilter);

        /// <summary>
        /// Writes typed rows to the current sheet.
        /// </summary>
        public void WriteSheet(List<string> headersList, List<TypeCode> typeCodes, List<object?[]> rows, Boolean headers = true, int maxRows = -1, int startingRow = 0, int startingColumn = 0, bool doAutofilter = false)
        {
            ArgumentNullException.ThrowIfNull(headersList);
            ArgumentNullException.ThrowIfNull(typeCodes);
            ArgumentNullException.ThrowIfNull(rows);

            using var reader = new ReaderFromList(rows, headersList, typeCodes);
            WriteSheet(reader, headers, maxRows, startingRow, startingColumn, doAutofilter);
        }

        /// <summary>Writes typed rows with a style for the first written row and/or A1 cell ranges.</summary>
        public void WriteSheet(List<string> headersList, List<TypeCode> typeCodes, List<object?[]> rows, CellStyle? headerStyle,
            IReadOnlyDictionary<string, CellStyle>? styles = null, Boolean headers = true, int maxRows = -1,
            int startingRow = 0, int startingColumn = 0, bool doAutofilter = false)
        {
            ArgumentNullException.ThrowIfNull(headersList);
            ArgumentNullException.ThrowIfNull(typeCodes);
            ArgumentNullException.ThrowIfNull(rows);
            using var reader = new ReaderFromList(rows, headersList, typeCodes);
            WriteSheet(reader, headerStyle, styles, headers, maxRows, startingRow, startingColumn, doAutofilter);
        }

        /// <summary>
        /// Obsolete. Use <see cref="WriteSheet(List{string}, List{TypeCode}, List{object?[]}, bool, int, int, int, bool)"/>
        /// with the <c>maxRows</c> parameter instead.
        /// </summary>
        [Obsolete("Use the maxRows parameter instead. This wrapper will be removed in a future release.")]
        public void WriteSheet(List<string> headersList, List<TypeCode> typeCodes, List<object?[]> rows, bool headers, int overLimit, int startingRow, int startingColumn, bool doAutofilter, bool _overLimitRenameMarker)
            => WriteSheet(headersList, typeCodes, rows, headers, overLimit, startingRow, startingColumn, doAutofilter);

        /// <summary>
        /// Writes a 2D jagged array of raw values to a new sheet. Each inner array is one row.
        /// Use <see cref="FormattedCell"/> items to apply a per-cell Excel number format.
        /// When <paramref name="headers"/> is true, the first row receives bold formatting and
        /// column names are auto-generated (<c>C0, C1, …</c>). For custom header names use
        /// the <see cref="WriteSheet(object?[][], string[], int, int, bool)"/> overload.
        /// </summary>
        /// <param name="rows">Rows of values. <c>null</c> cells are written as empty cells.</param>
        /// <param name="startingRow">Zero-based row index where writing begins.</param>
        /// <param name="startingColumn">Zero-based column index where writing begins.</param>
        /// <param name="doAutofilter">If true, add an autofilter on the written range.</param>
        /// <param name="headers">If true, the first row is treated as a header row (bold).</param>
        public virtual void WriteSheet(object?[][] rows, int startingRow = 0, int startingColumn = 0, bool doAutofilter = false, bool headers = true)
        {
            ArgumentNullException.ThrowIfNull(rows);
            using var dt = new DataTable();
            int columnCount = rows.Length == 0 ? 0 : (rows[0]?.Length ?? 0);
            for (int c = 0; c < columnCount; c++)
            {
                dt.Columns.Add($"C{c}", typeof(object));
            }
            foreach (var row in rows)
            {
                if (row == null) { dt.Rows.Add(dt.NewRow()); continue; }
                var dataRow = dt.NewRow();
                for (int c = 0; c < columnCount; c++)
                {
                    dataRow[c] = c < row.Length && row[c] != null ? row[c] : DBNull.Value;
                }
                dt.Rows.Add(dataRow);
            }
            WriteSheet(dt, headers: headers, startingRow: startingRow, startingColumn: startingColumn, doAutofilter: doAutofilter);
        }

        /// <summary>Writes raw rows with a style for the first written row and/or A1 cell ranges.</summary>
        public void WriteSheet(object?[][] rows, CellStyle? headerStyle, IReadOnlyDictionary<string, CellStyle>? styles = null,
            int startingRow = 0, int startingColumn = 0, bool doAutofilter = false, bool headers = true)
        {
            SetStyles(headerStyle, styles);
            try
            {
                WriteSheet(rows, startingRow, startingColumn, doAutofilter, headers);
            }
            finally
            {
                SetStyles(null, null);
            }
        }

        /// <summary>
        /// Writes a 2D jagged array of raw values with custom column header names.
        /// The first row of <paramref name="rows"/> is data (not headers); column names are
        /// taken from <paramref name="headers"/>. The header row receives bold formatting.
        /// </summary>
        /// <param name="rows">Rows of data values.</param>
        /// <param name="headers">Column header names. Must match the column count of the first data row.</param>
        /// <param name="startingRow">Zero-based row index where writing begins.</param>
        /// <param name="startingColumn">Zero-based column index where writing begins.</param>
        /// <param name="doAutofilter">If true, add an autofilter on the written range.</param>
        public virtual void WriteSheet(object?[][] rows, string[] headers, int startingRow = 0, int startingColumn = 0, bool doAutofilter = false)
        {
            ArgumentNullException.ThrowIfNull(rows);
            ArgumentNullException.ThrowIfNull(headers);
            using var dt = new DataTable();
            int columnCount = rows.Length == 0 ? 0 : (rows[0]?.Length ?? 0);
            for (int c = 0; c < columnCount; c++)
            {
                dt.Columns.Add(c < headers.Length ? headers[c] : $"C{c}", typeof(object));
            }
            foreach (var row in rows)
            {
                if (row == null) { dt.Rows.Add(dt.NewRow()); continue; }
                var dataRow = dt.NewRow();
                for (int c = 0; c < columnCount; c++)
                {
                    dataRow[c] = c < row.Length && row[c] != null ? row[c] : DBNull.Value;
                }
                dt.Rows.Add(dataRow);
            }
            WriteSheet(dt, headers: true, startingRow: startingRow, startingColumn: startingColumn, doAutofilter: doAutofilter);
        }

        /// <summary>Writes raw rows with named columns and a style for the first written row and/or A1 ranges.</summary>
        public void WriteSheet(object?[][] rows, string[] headers, CellStyle? headerStyle,
            IReadOnlyDictionary<string, CellStyle>? styles = null, int startingRow = 0, int startingColumn = 0, bool doAutofilter = false)
        {
            SetStyles(headerStyle, styles);
            try
            {
                WriteSheet(rows, headers, startingRow, startingColumn, doAutofilter);
            }
            finally
            {
                SetStyles(null, null);
            }
        }

        /// <summary>
        /// Writes a 2D jagged array of <see cref="FormattedCell"/> values to a new sheet.
        /// The <see cref="FormattedCell.Format"/> of each cell controls the Excel number format used.
        /// When <paramref name="headers"/> is true, the first row receives bold formatting and
        /// column names are auto-generated (<c>C0, C1, …</c>). For custom header names use
        /// the <see cref="WriteSheet(FormattedCell?[][], string[], int, int, bool)"/> overload.
        /// </summary>
        /// <param name="rows">Rows of formatted cells.</param>
        /// <param name="startingRow">Zero-based row index where writing begins.</param>
        /// <param name="startingColumn">Zero-based column index where writing begins.</param>
        /// <param name="doAutofilter">If true, add an autofilter on the written range.</param>
        /// <param name="headers">If true, the first row is treated as a header row (bold).</param>
        public virtual void WriteSheet(FormattedCell?[][] rows, int startingRow = 0, int startingColumn = 0, bool doAutofilter = false, bool headers = true)
        {
            ArgumentNullException.ThrowIfNull(rows);
            using var dt = new DataTable();
            int columnCount = rows.Length == 0 ? 0 : (rows[0]?.Length ?? 0);
            for (int c = 0; c < columnCount; c++)
            {
                dt.Columns.Add($"C{c}", typeof(object));
            }
            foreach (var row in rows)
            {
                if (row == null) { dt.Rows.Add(dt.NewRow()); continue; }
                var dataRow = dt.NewRow();
                for (int c = 0; c < columnCount; c++)
                {
                    dataRow[c] = (row[c].HasValue && row[c]!.Value.Value != null) ? row[c]!.Value : (object)DBNull.Value;
                }
                dt.Rows.Add(dataRow);
            }
            WriteSheet(dt, headers: headers, startingRow: startingRow, startingColumn: startingColumn, doAutofilter: doAutofilter);
        }

        /// <summary>Writes formatted rows with a style for the first written row and/or A1 cell ranges.</summary>
        public void WriteSheet(FormattedCell?[][] rows, CellStyle? headerStyle, IReadOnlyDictionary<string, CellStyle>? styles = null,
            int startingRow = 0, int startingColumn = 0, bool doAutofilter = false, bool headers = true)
        {
            SetStyles(headerStyle, styles);
            try
            {
                WriteSheet(rows, startingRow, startingColumn, doAutofilter, headers);
            }
            finally
            {
                SetStyles(null, null);
            }
        }

        /// <summary>
        /// Writes a 2D jagged array of <see cref="FormattedCell"/> values with custom column header names.
        /// The first row of <paramref name="rows"/> is data (not headers); column names are taken from
        /// <paramref name="headers"/>. The header row receives bold formatting.
        /// Each <see cref="FormattedCell.Format"/> controls the Excel number format used for that cell.
        /// </summary>
        /// <param name="rows">Rows of formatted cell data.</param>
        /// <param name="headers">Column header names. Must match the column count of the first data row.</param>
        /// <param name="startingRow">Zero-based row index where writing begins.</param>
        /// <param name="startingColumn">Zero-based column index where writing begins.</param>
        /// <param name="doAutofilter">If true, add an autofilter on the written range.</param>
        public virtual void WriteSheet(FormattedCell?[][] rows, string[] headers, int startingRow = 0, int startingColumn = 0, bool doAutofilter = false)
        {
            ArgumentNullException.ThrowIfNull(rows);
            ArgumentNullException.ThrowIfNull(headers);
            using var dt = new DataTable();
            int columnCount = rows.Length == 0 ? 0 : (rows[0]?.Length ?? 0);
            for (int c = 0; c < columnCount; c++)
            {
                dt.Columns.Add(c < headers.Length ? headers[c] : $"C{c}", typeof(object));
            }
            foreach (var row in rows)
            {
                if (row == null) { dt.Rows.Add(dt.NewRow()); continue; }
                var dataRow = dt.NewRow();
                for (int c = 0; c < columnCount; c++)
                {
                    dataRow[c] = (row[c].HasValue && row[c]!.Value.Value != null) ? row[c]!.Value : (object)DBNull.Value;
                }
                dt.Rows.Add(dataRow);
            }
            WriteSheet(dt, headers: true, startingRow: startingRow, startingColumn: startingColumn, doAutofilter: doAutofilter);
        }

        /// <summary>Writes formatted rows with named columns and a style for the first row and/or A1 ranges.</summary>
        public void WriteSheet(FormattedCell?[][] rows, string[] headers, CellStyle? headerStyle,
            IReadOnlyDictionary<string, CellStyle>? styles = null, int startingRow = 0, int startingColumn = 0, bool doAutofilter = false)
        {
            SetStyles(headerStyle, styles);
            try
            {
                WriteSheet(rows, headers, startingRow, startingColumn, doAutofilter);
            }
            finally
            {
                SetStyles(null, null);
            }
        }

        public abstract void WriteSheet(string[] oneColumn);

        internal bool _excelStreamWasProvided = false;
        internal bool _disposed = false;

        /// <summary>
        /// Finalizes the workbook and writes it to the underlying stream/path.
        /// Throws <see cref="ObjectDisposedException"/> if called after a previous successful <see cref="Save"/>.
        /// </summary>
        /// <exception cref="ObjectDisposedException">Thrown when called a second time after a successful first call.</exception>
        public virtual void Save()
        {
            if (_disposed) throw new ObjectDisposedException(nameof(ExcelWriter));
            FinalizeFile();
            _excelArchiveFile.Dispose();
            if (!_excelStreamWasProvided)
            {
                _newExcelFileStream.Dispose();
            }
            _disposed = true;
        }

        public event Action? OnCompress;
        internal void DoOnCompress()
        {
            OnCompress?.Invoke();
        }
        public event Action<int>? On10k;
        internal void DoOn10k(int arg)
        {
            On10k?.Invoke(arg);
        }
        public abstract void Dispose();

        /// <summary>
        /// If true, <c>DateTime</c> values with year 1000 (commonly used as a "no date" placeholder)
        /// are suppressed (skipped) when writing DateTime cells.
        /// </summary>
        public bool SuppressYear1000Dates { get; set; }

        /// <summary>
        /// Obsolete. Use <see cref="SuppressYear1000Dates"/> instead.
        /// </summary>
        [Obsolete("Use SuppressYear1000Dates instead. This misspelled member will be removed in a future release.")]
        public bool SuppressSomeDate
        {
            get => SuppressYear1000Dates;
            set => SuppressYear1000Dates = value;
        }

        internal void SetColsLength(int ColumnCount, object[] arr)
        {
            for (int l = 1; l <= ColumnCount; l++)
            {
                if (arr[l - 1] != null)
                {
                    var itm = arr[l - 1];
                    int lenn = 0;
                    if (itm is Memory<byte> mem)
                    {
                        lenn = mem.Length;
                    }
                    else
                    {
                        lenn = arr[l - 1].ToString().Length;
                    }
                    if (_colWidthsArray[l - 1] < 1.25 * lenn + 2)
                    {
                        _colWidthsArray[l - 1] = 1.25 * lenn + 2;
                    }
                }
            }
        }

        public static ExcelWriter CreateWriter(string path)
        {
            ArgumentNullException.ThrowIfNull(path);
            ArgumentException.ThrowIfNullOrWhiteSpace(path);

            if (path.EndsWith(".xlsb", StringComparison.OrdinalIgnoreCase))
            {
                return new XlsbWriter(path);
            }
            else if (path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return new XlsxWriter(path);
            }
            else
            {
                throw new Exception("Unknown file type !");
            }
        }
    }

    /// <summary>Font, fill and number-format ids of one cellXfs entry together with its style.</summary>
    internal readonly record struct StyleComponent(int FontId, int FillId, int NumFmtId, CellStyle Style);

}
