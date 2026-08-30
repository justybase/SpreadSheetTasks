using System.Buffers.Binary;
using System.Data;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace SpreadSheetTasks;

/// <summary>
/// Replaces the cell data of an existing XLSB worksheet while preserving the
/// workbook package, worksheet metadata, styles and other sheets.
/// </summary>
public sealed class XlsbUpdater : IDisposable
{
    private static readonly HashSet<uint> CellRecordIds =
    [0x0001, 0x0002, 0x0003, 0x0004, 0x0005, 0x0006, 0x0007, 0x0008, 0x0009, 0x000A, 0x000B];

    private const int RkIntegerLowerLimit = -(1 << 29);
    private const int RkIntegerUpperLimit = (1 << 29) - 1;

    private readonly UpdaterPackage _package;
    private readonly string _sourcePath;
    private readonly Dictionary<string, string> _sheetNameToPath = new(StringComparer.Ordinal);
    private readonly List<string> _sheetNames = [];
    private readonly List<string> _pivotCachePaths = [];
    private readonly Dictionary<string, string> _relationshipTargets = new(StringComparer.Ordinal);
    private string? _sharedStringsPath;
    private string? _stylesPath;
    private Biff12UpdaterUtils.SharedStringsState? _sharedStrings;
    private bool _uses1904DateSystem;
    private bool _disposed;

    /// <summary>Opens an existing XLSB file for in-memory updates.</summary>
    public XlsbUpdater(string inputPath)
    {
        ArgumentNullException.ThrowIfNull(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);

        _sourcePath = Path.GetFullPath(inputPath);
        UpdaterPackage.EnsureExtension(_sourcePath, ".xlsb", nameof(XlsbUpdater));
        _package = UpdaterPackage.Load(_sourcePath);
        if (_package.Contains("xl/workbook.xml"))
        {
            _package.Dispose();
            throw new InvalidDataException("The supplied file is XLSX. Use XlsxUpdater for .xlsx files.");
        }
        if (!_package.Contains("xl/workbook.bin"))
        {
            _package.Dispose();
            throw new InvalidDataException("The supplied file is not a valid XLSB workbook (xl/workbook.bin is missing).");
        }

        try
        {
            LoadWorkbookStructure();
        }
        catch
        {
            _package.Dispose();
            throw;
        }
    }

    /// <summary>Returns worksheet names in workbook order.</summary>
    public IReadOnlyList<string> GetSheetNames()
    {
        ThrowIfDisposed();
        return _sheetNames.ToArray();
    }

    /// <summary>Replaces worksheet data read from an IDataReader.</summary>
    public void ReplaceSheetData(string sheetName, IDataReader reader, ReplaceSheetDataOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(reader);
        ReplaceSheetDataCore(sheetName, UpdaterRows.FromReader(reader, options), options);
    }

    /// <summary>Replaces worksheet data read from a DataTable.</summary>
    public void ReplaceSheetData(string sheetName, DataTable table, ReplaceSheetDataOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(table);
        ReplaceSheetDataCore(sheetName, UpdaterRows.FromDataTable(table, options), options);
    }

    /// <summary>Replaces worksheet data supplied as a jagged array.</summary>
    public void ReplaceSheetData(string sheetName, object?[][] rows, ReplaceSheetDataOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(rows);
        ReplaceSheetDataCore(sheetName, UpdaterRows.FromArray(rows, options), options);
    }

    /// <summary>Replaces worksheet data supplied as a list of rows.</summary>
    public void ReplaceSheetData(string sheetName, IReadOnlyList<object?[]> rows, ReplaceSheetDataOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(rows);
        ReplaceSheetDataCore(sheetName, UpdaterRows.FromList(rows, options), options);
    }

    /// <summary>Saves the updated workbook, overwriting the source by default.</summary>
    public void Save(string? outputPath = null)
    {
        ThrowIfDisposed();
        if (outputPath is not null)
            UpdaterPackage.EnsureExtension(Path.GetFullPath(outputPath), ".xlsb", nameof(XlsbUpdater));
        _package.Save(outputPath);
    }

    /// <summary>Returns the updated workbook as a complete ZIP package.</summary>
    public byte[] ToArray()
    {
        ThrowIfDisposed();
        return _package.ToArray();
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;
        _package.Dispose();
    }

    private void ReplaceSheetDataCore(string sheetName, UpdaterRows data, ReplaceSheetDataOptions? options)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);

        if (!_sheetNameToPath.TryGetValue(sheetName, out string? sheetPath))
            throw new KeyNotFoundException($"XlsbUpdater: worksheet '{sheetName}' was not found.");

        byte[] originalSheet = _package.GetPart(sheetPath);
        EnsureSharedStringsLoaded();
        if (_sharedStrings is not null)
            _sharedStrings.TotalCount = (uint)Math.Max(0, (long)_sharedStrings.TotalCount - CountSharedStringReferences(originalSheet));

        Dictionary<int, int> dataStyles = [];
        Dictionary<int, int> headerStyles = [];
        int? dateStyle = null;
        if (options?.StyleFallback != ReplaceSheetDataStyleFallback.General)
        {
            (dataStyles, headerStyles) = CollectExistingStyles(originalSheet);
            dateStyle = FindDateStyleIndex();
        }

        byte[] replacementRows = BuildRows(data, dataStyles, headerStyles, dateStyle);
        RowRegion region = FindRowsRegion(originalSheet);
        int delta = replacementRows.Length - (region.RowsEnd - region.RowsStart);
        var updatedSheet = new byte[originalSheet.Length + delta];

        Buffer.BlockCopy(originalSheet, 0, updatedSheet, 0, region.RowsStart);
        Buffer.BlockCopy(replacementRows, 0, updatedSheet, region.RowsStart, replacementRows.Length);
        Buffer.BlockCopy(originalSheet, region.RowsEnd, updatedSheet, region.RowsStart + replacementRows.Length,
            originalSheet.Length - region.RowsEnd);

        int totalRows = data.TotalRowCount;
        int lastRow = totalRows == 0 ? 0 : totalRows - 1;
        int lastColumn = Math.Max(0, data.LastColumnIndex);
        foreach (int originalOffset in region.AutoFilterPayloadOffsets)
        {
            int offset = originalOffset >= region.RowsEnd ? originalOffset + delta : originalOffset;
            if (offset < 0 || offset > updatedSheet.Length - 16)
                continue;
            Biff12UpdaterUtils.WriteInt32(updatedSheet, offset + 4, lastRow);
            Biff12UpdaterUtils.WriteInt32(updatedSheet, offset + 12, lastColumn);
        }

        PatchDimension(updatedSheet, lastRow, lastColumn);
        _package.SetPart(sheetPath, updatedSheet);

        CommitSharedStrings();
        PatchPivotCaches(sheetName, lastRow, lastColumn);
    }

    private void LoadWorkbookStructure()
    {
        byte[]? relationships = _package.TryGetPart("xl/_rels/workbook.bin.rels");
        if (relationships is not null)
        {
            var targets = UpdaterXmlUtils.ReadRelationships(relationships, "xl/workbook.bin");
            foreach (var pair in targets)
            {
                _relationshipTargets[pair.Key] = pair.Value;
                if (pair.Value.EndsWith("sharedStrings.bin", StringComparison.OrdinalIgnoreCase))
                    _sharedStringsPath = pair.Value;
                else if (pair.Value.EndsWith("styles.bin", StringComparison.OrdinalIgnoreCase))
                    _stylesPath = pair.Value;
            }
        }

        byte[] workbook = _package.GetPart("xl/workbook.bin");
        int position = 0;
        while (TryReadRecordOrThrow(workbook, position, out var record))
        {
            position = record.DataEnd;
            if (record.Id == 0x0099 && record.Length >= sizeof(uint))
            {
                _uses1904DateSystem = (Biff12UpdaterUtils.ReadUInt32(workbook, record.DataStart) & 1) != 0;
                continue;
            }
            if (record.Id != 0x009C || record.Length < 12)
                continue;

            int cursor = record.DataStart + 8; // hidden + sheet id
            uint relationshipLength = Biff12UpdaterUtils.ReadUInt32(workbook, cursor);
            cursor += 4;
            if (relationshipLength > int.MaxValue || relationshipLength > (uint)((record.DataEnd - cursor) / 2))
                throw new InvalidDataException("Invalid XLSB worksheet relationship string.");
            string relationshipId = Biff12UpdaterUtils.ReadUtf16(workbook, cursor, (int)relationshipLength);
            cursor += checked((int)relationshipLength * 2);
            uint nameLength = Biff12UpdaterUtils.ReadUInt32(workbook, cursor);
            cursor += 4;
            if (nameLength > int.MaxValue || nameLength > (uint)((record.DataEnd - cursor) / 2))
                throw new InvalidDataException("Invalid XLSB worksheet name string.");
            string name = Biff12UpdaterUtils.ReadUtf16(workbook, cursor, (int)nameLength);

            if (!_relationshipTargets.TryGetValue(relationshipId, out string? location) || string.IsNullOrEmpty(location))
                continue;
            if (_sheetNameToPath.ContainsKey(name))
                throw new InvalidDataException($"The XLSB workbook contains duplicate worksheet name '{name}'.");
            _sheetNameToPath[name] = location;
            _sheetNames.Add(name);
        }

        foreach (string partName in _package.PartNames)
        {
            if (Regex.IsMatch(partName, @"^xl/pivotCache/pivotCacheDefinition\d*\.bin$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                _pivotCachePaths.Add(partName);
        }

        if (_sheetNames.Count == 0)
            throw new InvalidDataException("The XLSB workbook contains no resolvable worksheets.");
    }

    private void EnsureSharedStringsLoaded()
    {
        if (_sharedStrings is not null || _sharedStringsPath is null)
            return;

        byte[]? bytes = _package.TryGetPart(_sharedStringsPath);
        if (bytes is not null)
            _sharedStrings = Biff12UpdaterUtils.ParseSharedStrings(bytes);
    }

    private int CountSharedStringReferences(byte[] sheet)
    {
        int count = 0;
        int position = 0;
        while (TryReadRecordOrThrow(sheet, position, out var record))
        {
            position = record.DataEnd;
            if (record.Id == 0x0007)
                count++;
        }
        return count;
    }

    private void CommitSharedStrings()
    {
        if (_sharedStrings is null || !_sharedStrings.Dirty || _sharedStringsPath is null)
            return;

        int newValueCount = _sharedStrings.Values.Count - _sharedStrings.CommittedValueCount;
        using var appended = new MemoryStream();
        for (int i = _sharedStrings.CommittedValueCount; i < _sharedStrings.Values.Count; i++)
        {
            string value = _sharedStrings.Values[i];
            var payload = new byte[5 + value.Length * 2];
            payload[0] = 0;
            Biff12UpdaterUtils.WriteUInt32(payload, 1, (uint)value.Length);
            Encoding.Unicode.GetBytes(value, payload.AsSpan(5));
            byte[] record = Biff12UpdaterUtils.BuildRecord(0x0013, payload);
            appended.Write(record, 0, record.Length);
        }

        byte[] old = _sharedStrings.Buffer;
        byte[] added = appended.ToArray();
        byte[] updated = new byte[old.Length + added.Length];
        Buffer.BlockCopy(old, 0, updated, 0, _sharedStrings.EndSstOffset);
        Buffer.BlockCopy(added, 0, updated, _sharedStrings.EndSstOffset, added.Length);
        Buffer.BlockCopy(old, _sharedStrings.EndSstOffset, updated, _sharedStrings.EndSstOffset + added.Length,
            old.Length - _sharedStrings.EndSstOffset);

        int position = 0;
        bool beginFound = false;
        while (TryReadRecordOrThrow(updated, position, out var record))
        {
            position = record.DataEnd;
            if (record.Id != 0x009F || record.Length < 8)
                continue;
            Biff12UpdaterUtils.WriteUInt32(updated, record.DataStart, _sharedStrings.TotalCount);
            Biff12UpdaterUtils.WriteUInt32(updated, record.DataStart + 4, (uint)_sharedStrings.Values.Count);
            beginFound = true;
            break;
        }
        if (!beginFound)
            throw new InvalidDataException("The XLSB Shared Strings part has no BrtBeginSst record.");

        _sharedStrings.Buffer = updated;
        _sharedStrings.EndSstOffset += added.Length;
        _sharedStrings.CommittedValueCount += newValueCount;
        _sharedStrings.UniqueCount = (uint)_sharedStrings.Values.Count;
        _sharedStrings.Dirty = false;
        _package.SetPart(_sharedStringsPath, updated);
    }

    private (Dictionary<int, int> dataStyles, Dictionary<int, int> headerStyles) CollectExistingStyles(byte[] sheet)
    {
        var dataCounts = new Dictionary<int, Dictionary<int, int>>();
        var headerStyles = new Dictionary<int, int>();
        int currentRow = -1;
        int position = 0;
        while (TryReadRecordOrThrow(sheet, position, out var record))
        {
            position = record.DataEnd;
            if (record.Id == 0x0000)
            {
                if (record.Length >= 4)
                    currentRow = Biff12UpdaterUtils.ReadInt32(sheet, record.DataStart);
                continue;
            }
            if (!CellRecordIds.Contains(record.Id) || record.Length < 8)
                continue;

            int column = checked((int)Biff12UpdaterUtils.ReadUInt32(sheet, record.DataStart));
            int style = (int)(Biff12UpdaterUtils.ReadUInt32(sheet, record.DataStart + 4) & 0x00FF_FFFF);
            if (currentRow == 0)
            {
                headerStyles.TryAdd(column, style);
                continue;
            }
            if (style <= 0)
                continue;
            if (!dataCounts.TryGetValue(column, out var counts))
            {
                counts = new Dictionary<int, int>();
                dataCounts[column] = counts;
            }
            counts[style] = counts.GetValueOrDefault(style) + 1;
        }

        var dataStyles = new Dictionary<int, int>();
        foreach (var pair in dataCounts)
        {
            int bestStyle = 0;
            int bestCount = -1;
            foreach (var style in pair.Value)
            {
                if (style.Value > bestCount)
                {
                    bestStyle = style.Key;
                    bestCount = style.Value;
                }
            }
            if (bestCount > 0)
                dataStyles[pair.Key] = bestStyle;
        }
        return (dataStyles, headerStyles);
    }

    private int? FindDateStyleIndex()
    {
        if (_stylesPath is null)
            return null;
        byte[]? styles = _package.TryGetPart(_stylesPath);
        if (styles is null)
            return null;

        bool inCellXfs = false;
        int styleIndex = 0;
        int position = 0;
        while (TryReadRecordOrThrow(styles, position, out var record))
        {
            position = record.DataEnd;
            if (record.Id == 0x0269)
            {
                inCellXfs = true;
                styleIndex = 0;
                continue;
            }
            if (record.Id == 0x026A)
            {
                inCellXfs = false;
                continue;
            }
            if (inCellXfs && record.Id == 0x002F && record.Length >= 4)
            {
                int numberFormat = BinaryPrimitives.ReadUInt16LittleEndian(styles.AsSpan(record.DataStart + 2));
                if ((numberFormat >= 14 && numberFormat <= 22) || (numberFormat >= 45 && numberFormat <= 47))
                    return styleIndex;
                styleIndex++;
            }
        }
        return null;
    }

    private byte[] BuildRows(
        UpdaterRows data,
        IReadOnlyDictionary<int, int> dataStyles,
        IReadOnlyDictionary<int, int> headerStyles,
        int? dateStyle)
    {
        int maxColumn = Math.Max(0, data.LastColumnIndex + 1);
        int lastColumn = Math.Max(0, maxColumn - 1);
        using var output = new MemoryStream();

        int rowNumber = 0;
        if (data.Headers is not null)
        {
            object?[] headers = data.Headers.Select(static value => (object?)value).ToArray();
            WriteRow(output, rowNumber++, headers, lastColumn, headerStyles, dateStyle: null, isHeader: true);
        }

        foreach (var row in data.Rows)
        {
            WriteRow(output, rowNumber++, row, lastColumn, dataStyles, dateStyle, isHeader: false);
        }
        return output.ToArray();
    }

    private void WriteRow(
        Stream output,
        int rowNumber,
        IReadOnlyList<object?> values,
        int lastColumn,
        IReadOnlyDictionary<int, int> styles,
        int? dateStyle,
        bool isHeader)
    {
        var rowPayload = new byte[25];
        Biff12UpdaterUtils.WriteInt32(rowPayload, 0, rowNumber);
        rowPayload[8] = 0x2C;
        rowPayload[9] = 0x01;
        rowPayload[13] = 0x01;
        Biff12UpdaterUtils.WriteInt32(rowPayload, 17, 0);
        Biff12UpdaterUtils.WriteInt32(rowPayload, 21, lastColumn);
        byte[] rowRecord = Biff12UpdaterUtils.BuildRecord(0x0000, rowPayload);
        output.Write(rowRecord, 0, rowRecord.Length);

        for (int column = 0; column < values.Count; column++)
        {
            object? value = UpdaterRows.Unwrap(values[column]);
            if (UpdaterRows.IsEmpty(value))
                continue;
            byte[] bytes = isHeader
                ? BuildStringCell(column, styles.GetValueOrDefault(column), Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty)
                : BuildValueCell(value!, column, styles, dateStyle);
            output.Write(bytes, 0, bytes.Length);
        }
    }

    private byte[] BuildValueCell(object value, int column, IReadOnlyDictionary<int, int> styles, int? dateStyle)
    {
        value = UpdaterRows.Unwrap(value)!;
        int style = styles.GetValueOrDefault(column);

        if (value is DateTime dateTime)
        {
            try
            {
                double serial = ToExcelDateSerial(dateTime);
                if (double.IsFinite(serial))
                    return BuildDoubleCell(column, style > 0 ? style : dateStyle ?? 0, serial);
            }
            catch (ArgumentException)
            {
                // Store an unrepresentable date as text.
            }
            return BuildStringCell(column, style, dateTime.ToString(CultureInfo.InvariantCulture));
        }
        if (value is DateTimeOffset dateTimeOffset)
            return BuildValueCell(dateTimeOffset.DateTime, column, styles, dateStyle);
        if (value is bool boolean)
        {
            var payload = new byte[9];
            Biff12UpdaterUtils.WriteUInt32(payload, 0, (uint)column);
            Biff12UpdaterUtils.WriteUInt32(payload, 4, (uint)style);
            payload[8] = boolean ? (byte)1 : (byte)0;
            return Biff12UpdaterUtils.BuildRecord(0x0004, payload);
        }

        if (TryConvertNumber(value, out double number))
        {
            if (double.IsFinite(number))
            {
                if (number >= RkIntegerLowerLimit && number <= RkIntegerUpperLimit && number == Math.Truncate(number))
                {
                    var payload = new byte[12];
                    Biff12UpdaterUtils.WriteUInt32(payload, 0, (uint)column);
                    Biff12UpdaterUtils.WriteUInt32(payload, 4, (uint)style);
                    Biff12UpdaterUtils.WriteInt32(payload, 8, ((int)number << 2) | 2);
                    return Biff12UpdaterUtils.BuildRecord(0x0002, payload);
                }
                return BuildDoubleCell(column, style, number);
            }
            return BuildStringCell(column, style, number.ToString(CultureInfo.InvariantCulture));
        }

        string text = value switch
        {
            string stringValue => stringValue,
            char character => character.ToString(),
            byte[] bytes => Encoding.UTF8.GetString(bytes),
            Memory<byte> memory => Encoding.UTF8.GetString(memory.Span),
            _ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty
        };
        return BuildStringCell(column, style, text);
    }

    private double ToExcelDateSerial(DateTime value)
    {
        const double DateSystemOffset = 1462d;
        return value.ToOADate() - (_uses1904DateSystem ? DateSystemOffset : 0d);
    }

    private static bool TryConvertNumber(object value, out double number)
    {
        number = 0;
        if (value is bool || value is DateTime || value is DateTimeOffset || value is string || value is char)
            return false;
        TypeCode typeCode = Type.GetTypeCode(value.GetType());
        if (typeCode is not (TypeCode.Byte or TypeCode.SByte or TypeCode.UInt16 or TypeCode.UInt32 or TypeCode.UInt64 or
            TypeCode.Int16 or TypeCode.Int32 or TypeCode.Int64 or TypeCode.Single or TypeCode.Double or TypeCode.Decimal))
        {
            return false;
        }
        try
        {
            number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
            return true;
        }
        catch (Exception) when (value is IConvertible)
        {
            return false;
        }
    }

    private byte[] BuildDoubleCell(int column, int style, double value)
    {
        var payload = new byte[16];
        Biff12UpdaterUtils.WriteUInt32(payload, 0, (uint)column);
        Biff12UpdaterUtils.WriteUInt32(payload, 4, (uint)style);
        Biff12UpdaterUtils.WriteDouble(payload, 8, value);
        return Biff12UpdaterUtils.BuildRecord(0x0005, payload);
    }

    private byte[] BuildStringCell(int column, int style, string value)
    {
        if (_sharedStrings is null)
        {
            var payload = new byte[12 + value.Length * 2];
            Biff12UpdaterUtils.WriteUInt32(payload, 0, (uint)column);
            Biff12UpdaterUtils.WriteUInt32(payload, 4, (uint)style);
            Biff12UpdaterUtils.WriteUInt32(payload, 8, (uint)value.Length);
            Encoding.Unicode.GetBytes(value, payload.AsSpan(12));
            return Biff12UpdaterUtils.BuildRecord(0x0006, payload);
        }

        if (!_sharedStrings.Index.TryGetValue(value, out int index))
        {
            index = _sharedStrings.Values.Count;
            _sharedStrings.Values.Add(value);
            _sharedStrings.Index[value] = index;
        }
        _sharedStrings.TotalCount++;
        _sharedStrings.Dirty = true;

        var sharedPayload = new byte[12];
        Biff12UpdaterUtils.WriteUInt32(sharedPayload, 0, (uint)column);
        Biff12UpdaterUtils.WriteUInt32(sharedPayload, 4, (uint)style);
        Biff12UpdaterUtils.WriteUInt32(sharedPayload, 8, (uint)index);
        return Biff12UpdaterUtils.BuildRecord(0x0007, sharedPayload);
    }

    private readonly struct RowRegion
    {
        internal RowRegion(int rowsStart, int rowsEnd, List<int> autoFilterPayloadOffsets)
        {
            RowsStart = rowsStart;
            RowsEnd = rowsEnd;
            AutoFilterPayloadOffsets = autoFilterPayloadOffsets;
        }

        internal int RowsStart { get; }
        internal int RowsEnd { get; }
        internal List<int> AutoFilterPayloadOffsets { get; }
    }

    private RowRegion FindRowsRegion(byte[] sheet)
    {
        int rowsStart = -1;
        int lastRowEnd = -1;
        int lastCellEnd = -1;
        int lastWrapper = -1;
        var autoFilters = new List<int>();
        int position = 0;
        while (TryReadRecordOrThrow(sheet, position, out var record))
        {
            position = record.DataEnd;
            if (record.Id == 0x0000)
            {
                if (rowsStart < 0)
                    rowsStart = record.HeaderStart;
                lastRowEnd = record.DataEnd;
            }
            else if (CellRecordIds.Contains(record.Id))
            {
                lastCellEnd = record.DataEnd;
            }
            else if (record.Id == 0x0025 && record.Length == 6)
            {
                lastWrapper = record.HeaderStart;
            }
            else if (record.Id == 0x00A1 && record.Length >= 16)
            {
                autoFilters.Add(record.DataStart);
            }
        }

        if (rowsStart < 0)
        {
            int insertAt = lastWrapper >= 0 ? lastWrapper : sheet.Length;
            return new RowRegion(insertAt, insertAt, autoFilters);
        }

        int rowsEnd = Math.Max(lastRowEnd, lastCellEnd);
        if (rowsEnd < 0)
            rowsEnd = lastWrapper >= 0 ? lastWrapper : sheet.Length;
        return new RowRegion(rowsStart, rowsEnd, autoFilters);
    }

    private static void PatchDimension(byte[] sheet, int lastRow, int lastColumn)
    {
        int position = 0;
        while (TryReadRecordOrThrow(sheet, position, out var record))
        {
            position = record.DataEnd;
            if (record.Id == 0x0098 && record.Length >= 36)
            {
                Biff12UpdaterUtils.WriteInt32(sheet, record.DataStart + 24, lastRow);
                Biff12UpdaterUtils.WriteInt32(sheet, record.DataStart + 32, lastColumn);
                return;
            }
        }
    }

    private void PatchPivotCaches(string sheetName, int lastRow, int lastColumn)
    {
        foreach (string path in _pivotCachePaths)
        {
            byte[] cache = _package.GetPart(path);
            bool affected = false;
            int? refreshFlagsOffset = null;
            int position = 0;
            while (TryReadRecordOrThrow(cache, position, out var record))
            {
                position = record.DataEnd;
                if (record.Id == 0x00B3 && record.Length >= 4)
                {
                    // BrtBeginPivotCacheDefinition flags: refreshOnLoad is bit 0x04.
                    refreshFlagsOffset = record.DataStart + 3;
                }
                else if (record.Id == 0x00BB && record.Length >= 23)
                {
                    uint characterCount = Biff12UpdaterUtils.ReadUInt32(cache, record.DataStart + 3);
                    if (characterCount > int.MaxValue || characterCount > (uint)((record.DataEnd - (record.DataStart + 7)) / 2))
                        throw new InvalidDataException("Invalid XLSB pivot cache worksheet source.");
                    int nameStart = record.DataStart + 7;
                    int rangeStart = checked(nameStart + (int)characterCount * 2);
                    string sourceSheet = Biff12UpdaterUtils.ReadUtf16(cache, nameStart, (int)characterCount);
                    if (string.Equals(sourceSheet, sheetName, StringComparison.Ordinal))
                    {
                        if (rangeStart > record.DataEnd - 16)
                            throw new InvalidDataException("XLSB pivot cache worksheet source is truncated.");
                        Biff12UpdaterUtils.WriteInt32(cache, rangeStart + 4, lastRow);
                        Biff12UpdaterUtils.WriteInt32(cache, rangeStart + 12, lastColumn);
                        if (refreshFlagsOffset.HasValue)
                            cache[refreshFlagsOffset.Value] |= 0x04;
                        affected = true;
                    }
                }
            }

            if (affected)
                _package.SetPart(path, cache);
        }
    }

    private static bool TryReadRecordOrThrow(byte[] buffer, int position, out Biff12Record record)
    {
        if (position == buffer.Length)
        {
            record = default;
            return false;
        }
        if (!Biff12UpdaterUtils.TryReadRecord(buffer, position, out record))
            throw new InvalidDataException($"Malformed BIFF12 record at byte offset {position}.");
        return true;
    }

    private void ThrowIfDisposed()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
    }
}
