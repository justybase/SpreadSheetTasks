using System.Data;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace SpreadSheetTasks;

/// <summary>
/// Replaces the cell data of an existing XLSX worksheet while preserving the
/// workbook package, worksheet metadata, styles and other sheets.
/// </summary>
public sealed class XlsxUpdater : IDisposable
{
    // Keep small imports in memory; staging them to disk only pays off above this measured break-even range.
    private const int ReaderInMemoryRowLimit = 4_096;

    private readonly record struct StreamedSheetData(int DataRowCount, int TotalRowCount, int LastColumnIndex);

    private readonly record struct SharedStringsCheckpoint(
        SharedStringsState? State,
        int ValueCount,
        int TotalCount,
        int CommittedValueCount,
        bool Dirty);

    private readonly record struct XmlReplacement(
        int Start,
        int Length,
        string? PrefixText,
        string? FilePath,
        string? SuffixText);

    private sealed class SharedStringsState
    {
        internal string Xml { get; set; } = string.Empty;
        internal List<string> Values { get; } = [];
        internal Dictionary<string, int> Index { get; } = new(StringComparer.Ordinal);
        internal int TotalCount { get; set; }
        internal int CommittedValueCount { get; set; }
        internal bool Dirty { get; set; }
    }

    private readonly UpdaterPackage _package;
    private readonly string _sourcePath;
    private readonly Dictionary<string, string> _sheetNameToPath = new(StringComparer.Ordinal);
    private readonly List<string> _sheetNames = [];
    private readonly List<string> _pivotCachePaths = [];
    private readonly List<string> _pivotTablePaths = [];
    private readonly Dictionary<string, string> _relationshipTargets = new(StringComparer.Ordinal);
    private string? _sharedStringsPath;
    private string? _stylesPath;
    private SharedStringsState? _sharedStrings;
    private bool _uses1904DateSystem;
    private bool _disposed;

    /// <summary>Opens an existing XLSX file for in-memory updates.</summary>
    public XlsxUpdater(string inputPath)
    {
        ArgumentNullException.ThrowIfNull(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);

        _sourcePath = Path.GetFullPath(inputPath);
        UpdaterPackage.EnsureExtension(_sourcePath, ".xlsx", nameof(XlsxUpdater));
        _package = UpdaterPackage.Load(_sourcePath);

        if (_package.Contains("xl/workbook.bin"))
        {
            _package.Dispose();
            throw new InvalidDataException("The supplied file is XLSB. Use XlsbUpdater for .xlsb files.");
        }
        if (!_package.Contains("xl/workbook.xml"))
        {
            _package.Dispose();
            throw new InvalidDataException("The supplied file is not a valid XLSX workbook (xl/workbook.xml is missing).");
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

    /// <summary>
    /// Replaces worksheet data read from an IDataReader. Larger inputs may be staged in the system
    /// temporary directory to reduce peak managed-memory use.
    /// </summary>
    public void ReplaceSheetData(string sheetName, IDataReader reader, ReplaceSheetDataOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(reader);
        ReplaceSheetDataFromReaderCore(sheetName, reader, options);
    }

    /// <summary>
    /// Replaces worksheet data read from a DataTable. Larger inputs may be staged in the system
    /// temporary directory to reduce peak managed-memory use.
    /// </summary>
    public void ReplaceSheetData(string sheetName, DataTable table, ReplaceSheetDataOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(table);
        if (table.Rows.Count <= ReaderInMemoryRowLimit)
        {
            ReplaceSheetDataCore(sheetName, UpdaterRows.FromDataTable(table, options), options);
            return;
        }
        using var reader = table.CreateDataReader();
        ReplaceSheetData(sheetName, reader, options);
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

    /// <summary>
    /// Saves the updated workbook. When no path is supplied, the source file is
    /// replaced safely through a temporary file.
    /// </summary>
    public void Save(string? outputPath = null)
    {
        ThrowIfDisposed();
        if (outputPath is not null)
            UpdaterPackage.EnsureExtension(Path.GetFullPath(outputPath), ".xlsx", nameof(XlsxUpdater));
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
            throw new KeyNotFoundException($"XlsxUpdater: worksheet '{sheetName}' was not found.");

        byte[] originalBytes = _package.GetPart(sheetPath);
        string originalXml = Encoding.UTF8.GetString(originalBytes);

        EnsureSharedStringsLoaded();
        if (_sharedStrings is not null)
            _sharedStrings.TotalCount = Math.Max(0, _sharedStrings.TotalCount - CountSharedStringReferences(originalXml));

        Dictionary<int, int> dataStyles = [];
        Dictionary<int, int> headerStyles = [];
        int? dateStyle = null;
        if (options?.StyleFallback != ReplaceSheetDataStyleFallback.General)
        {
            (dataStyles, headerStyles) = CollectExistingStyles(originalXml);
            dateStyle = FindDateStyleIndex();
        }

        string sheetData = BuildSheetData(data, dataStyles, headerStyles, dateStyle, out string elementPrefix);
        string updatedXml = PatchWorksheetXml(originalXml, sheetData, data.TotalRowCount, data.LastColumnIndex, elementPrefix);
        _package.SetPart(sheetPath, Encoding.UTF8.GetBytes(updatedXml));

        CommitSharedStrings();

        string dimension = data.TotalRowCount > 0 && data.LastColumnIndex >= 0
            ? $"A1:{UpdaterXmlUtils.ColumnToLetters(data.LastColumnIndex)}{data.TotalRowCount}"
            : "A1";
        HashSet<string> affectedCacheIds = UpdatePivotCaches(sheetName, dimension, data.DataRowCount);
        UpdatePivotTables(affectedCacheIds);
    }

    private void ReplaceSheetDataFromReaderCore(string sheetName, IDataReader reader, ReplaceSheetDataOptions? options)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        UpdaterRows.ValidateHeaders(options?.Headers);

        if (!_sheetNameToPath.TryGetValue(sheetName, out string? sheetPath))
            throw new KeyNotFoundException($"XlsxUpdater: worksheet '{sheetName}' was not found.");

        List<object?[]> initialRows = UpdaterRows.ReadReaderPrefix(reader, ReaderInMemoryRowLimit, out bool hasMore);
        if (!hasMore)
        {
            ReplaceSheetDataCore(sheetName, UpdaterRows.FromRows(initialRows, options), options);
            return;
        }

        ReplaceSheetDataFromLargeReaderCore(sheetName, sheetPath, reader, initialRows, options);
    }

    private void ReplaceSheetDataFromLargeReaderCore(
        string sheetName,
        string sheetPath,
        IDataReader reader,
        IReadOnlyList<object?[]> initialRows,
        ReplaceSheetDataOptions? options)
    {
        string originalXml = Encoding.UTF8.GetString(_package.GetPart(sheetPath));
        EnsureSharedStringsLoaded();
        SharedStringsCheckpoint checkpoint = CaptureSharedStringsCheckpoint();

        string rowsPath = CreateTemporaryPath(".sheet-data");
        string? updatedSheetPath = CreateTemporaryPath(".worksheet");
        bool worksheetStaged = false;
        StreamedSheetData data = default;
        try
        {
            if (_sharedStrings is not null)
                _sharedStrings.TotalCount = Math.Max(0, _sharedStrings.TotalCount - CountSharedStringReferences(originalXml));

            Dictionary<int, int> dataStyles = [];
            Dictionary<int, int> headerStyles = [];
            int? dateStyle = null;
            if (options?.StyleFallback != ReplaceSheetDataStyleFallback.General)
            {
                (dataStyles, headerStyles) = CollectExistingStyles(originalXml);
                dateStyle = FindDateStyleIndex();
            }

            Match sheetDataOpen = UpdaterXmlUtils.FindStartTag(originalXml, "sheetData");
            string prefix = UpdaterXmlUtils.PrefixFromTag(sheetDataOpen);
            data = WriteSheetDataFile(
                reader, initialRows, options?.Headers, dataStyles, headerStyles, dateStyle, prefix, rowsPath);
            PatchWorksheetXmlToFile(originalXml, rowsPath, updatedSheetPath, data);

            _package.SetPartFromFile(sheetPath, updatedSheetPath);
            worksheetStaged = true;
            updatedSheetPath = null;
        }
        catch
        {
            if (!worksheetStaged)
                RestoreSharedStringsCheckpoint(checkpoint);
            throw;
        }
        finally
        {
            DeleteTemporaryFile(rowsPath);
            if (updatedSheetPath is not null)
                DeleteTemporaryFile(updatedSheetPath);
        }

        CommitSharedStrings();

        string dimension = data.TotalRowCount > 0 && data.LastColumnIndex >= 0
            ? $"A1:{UpdaterXmlUtils.ColumnToLetters(data.LastColumnIndex)}{data.TotalRowCount}"
            : "A1";
        HashSet<string> affectedCacheIds = UpdatePivotCaches(sheetName, dimension, data.DataRowCount);
        UpdatePivotTables(affectedCacheIds);
    }

    private StreamedSheetData WriteSheetDataFile(
        IDataReader reader,
        IReadOnlyList<object?[]> initialRows,
        IReadOnlyList<string>? headers,
        IReadOnlyDictionary<int, int> dataStyles,
        IReadOnlyDictionary<int, int> headerStyles,
        int? dateStyle,
        string prefix,
        string path)
    {
        using var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None,
            bufferSize: 1024 * 1024, FileOptions.SequentialScan);
        using var writer = new StreamWriter(output, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), 65_536);

        int rowNumber = 1;
        int dataRowCount = 0;
        int totalRowCount = 0;
        int lastColumnIndex = -1;
        int pendingEmptyRows = 0;

        if (headers is not null)
        {
            WriteRowStart(writer, prefix, rowNumber);
            for (int column = 0; column < headers.Count; column++)
                writer.Write(BuildStringCell(headers[column], column, rowNumber, headerStyles.GetValueOrDefault(column), prefix));
            WriteRowEnd(writer, prefix);
            lastColumnIndex = headers.Count - 1;
            rowNumber++;
            totalRowCount++;
        }

        int fieldCount = reader.FieldCount;
        void WriteDataRow(object?[] rowValues)
        {
            if (UpdaterRows.IsEmptyRow(rowValues))
            {
                pendingEmptyRows++;
                return;
            }

            while (pendingEmptyRows > 0)
            {
                WriteEmptyRow(writer, prefix, rowNumber++);
                pendingEmptyRows--;
                dataRowCount++;
                totalRowCount++;
                lastColumnIndex = Math.Max(lastColumnIndex, fieldCount - 1);
            }

            WriteRowStart(writer, prefix, rowNumber);
            for (int column = 0; column < fieldCount; column++)
            {
                object? value = UpdaterRows.Unwrap(rowValues[column]);
                if (UpdaterRows.IsEmpty(value))
                    continue;
                writer.Write(BuildValueCell(value!, column, rowNumber, dataStyles.GetValueOrDefault(column), dateStyle, prefix));
            }
            WriteRowEnd(writer, prefix);
            rowNumber++;
            dataRowCount++;
            totalRowCount++;
            lastColumnIndex = Math.Max(lastColumnIndex, fieldCount - 1);
        }

        foreach (object?[] initialRow in initialRows)
            WriteDataRow(initialRow);

        var row = new object?[fieldCount];
        while (reader.Read())
        {
            for (int column = 0; column < fieldCount; column++)
            {
                object value = reader.GetValue(column);
                row[column] = value == DBNull.Value ? null : value;
            }
            WriteDataRow(row);
        }

        writer.Flush();
        return new StreamedSheetData(dataRowCount, totalRowCount, lastColumnIndex);
    }

    private static void WriteRowStart(TextWriter writer, string prefix, int rowNumber)
    {
        writer.Write('<');
        writer.Write(prefix);
        writer.Write("row r=\"");
        writer.Write(rowNumber.ToString(CultureInfo.InvariantCulture));
        writer.Write("\">");
    }

    private static void WriteRowEnd(TextWriter writer, string prefix)
    {
        writer.Write("</");
        writer.Write(prefix);
        writer.Write("row>");
    }

    private static void WriteEmptyRow(TextWriter writer, string prefix, int rowNumber)
    {
        WriteRowStart(writer, prefix, rowNumber);
        WriteRowEnd(writer, prefix);
    }

    private void PatchWorksheetXmlToFile(string xml, string sheetDataPath, string outputPath, StreamedSheetData data)
    {
        string dimension = data.TotalRowCount > 0 && data.LastColumnIndex >= 0
            ? $"A1:{UpdaterXmlUtils.ColumnToLetters(data.LastColumnIndex)}{data.TotalRowCount}"
            : "A1";
        Match root = UpdaterXmlUtils.FindStartTag(xml, "worksheet");
        Match? dimensionTag = FindOptionalStartTag(xml, "dimension");
        Match sheetDataOpen = UpdaterXmlUtils.FindStartTag(xml, "sheetData");
        string prefix = UpdaterXmlUtils.PrefixFromTag(sheetDataOpen);
        string closeTag = $"</{prefix}sheetData>";
        var replacements = new List<XmlReplacement>(3);

        if (dimensionTag is not null)
        {
            string replacement = UpdaterXmlUtils.ReplaceTagAttribute(dimensionTag.Value, "ref", dimension);
            replacements.Add(new XmlReplacement(dimensionTag.Index, dimensionTag.Length, replacement, null, null));
        }
        else
        {
            string rootPrefix = UpdaterXmlUtils.PrefixFromTag(root);
            string replacement = $"<{rootPrefix}dimension ref=\"{dimension}\"/>";
            replacements.Add(new XmlReplacement(root.Index + root.Length, 0, replacement, null, null));
        }

        string openTag = sheetDataOpen.Value;
        if (openTag.TrimEnd().EndsWith("/>", StringComparison.Ordinal))
        {
            string expandedOpenTag = openTag.TrimEnd()[..^2] + ">";
            replacements.Add(new XmlReplacement(
                sheetDataOpen.Index,
                sheetDataOpen.Length,
                expandedOpenTag,
                sheetDataPath,
                closeTag));
        }
        else
        {
            int contentStart = sheetDataOpen.Index + sheetDataOpen.Length;
            int contentEnd = xml.IndexOf(closeTag, contentStart, StringComparison.Ordinal);
            if (contentEnd < 0)
                throw new InvalidDataException("The XLSX worksheet sheetData element is not closed.");
            replacements.Add(new XmlReplacement(contentStart, contentEnd - contentStart, string.Empty, sheetDataPath, string.Empty));
        }

        Match? autoFilter = FindOptionalStartTag(xml, "autoFilter");
        if (autoFilter is not null)
        {
            string replacement = UpdaterXmlUtils.ReplaceTagAttribute(autoFilter.Value, "ref", dimension);
            replacements.Add(new XmlReplacement(autoFilter.Index, autoFilter.Length, replacement, null, null));
        }

        replacements.Sort(static (left, right) => left.Start.CompareTo(right.Start));
        using var destination = new FileStream(outputPath, FileMode.CreateNew, FileAccess.Write, FileShare.None,
            bufferSize: 1024 * 1024, FileOptions.SequentialScan);
        using var writer = new StreamWriter(destination, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), 65_536);
        int position = 0;
        foreach (XmlReplacement replacement in replacements)
        {
            if (replacement.Start < position)
                throw new InvalidDataException("Overlapping worksheet XML replacements were detected.");

            writer.Write(xml.AsSpan(position, replacement.Start - position));
            if (replacement.PrefixText is not null)
                writer.Write(replacement.PrefixText);
            if (replacement.FilePath is not null)
            {
                writer.Flush();
                using var input = new FileStream(replacement.FilePath, FileMode.Open, FileAccess.Read, FileShare.Read,
                    bufferSize: 1024 * 1024, FileOptions.SequentialScan);
                input.CopyTo(destination, 1024 * 1024);
            }
            if (replacement.SuffixText is not null)
                writer.Write(replacement.SuffixText);
            position = replacement.Start + replacement.Length;
        }
        writer.Write(xml.AsSpan(position));
    }

    private SharedStringsCheckpoint CaptureSharedStringsCheckpoint()
    {
        if (_sharedStrings is null)
            return default;

        return new SharedStringsCheckpoint(
            _sharedStrings,
            _sharedStrings.Values.Count,
            _sharedStrings.TotalCount,
            _sharedStrings.CommittedValueCount,
            _sharedStrings.Dirty);
    }

    private void RestoreSharedStringsCheckpoint(SharedStringsCheckpoint checkpoint)
    {
        if (checkpoint.State is null)
        {
            _sharedStrings = null;
            return;
        }

        SharedStringsState state = checkpoint.State;
        for (int i = state.Values.Count - 1; i >= checkpoint.ValueCount; i--)
            state.Index.Remove(state.Values[i]);
        if (state.Values.Count > checkpoint.ValueCount)
            state.Values.RemoveRange(checkpoint.ValueCount, state.Values.Count - checkpoint.ValueCount);
        state.TotalCount = checkpoint.TotalCount;
        state.CommittedValueCount = checkpoint.CommittedValueCount;
        state.Dirty = checkpoint.Dirty;
        _sharedStrings = state;
    }

    private static string CreateTemporaryPath(string suffix)
    {
        return Path.Combine(Path.GetTempPath(), $"SpreadSheetTasks-{Guid.NewGuid():N}{suffix}");
    }

    private static void DeleteTemporaryFile(string path)
    {
        try
        {
            File.Delete(path);
        }
        catch (FileNotFoundException)
        {
            // A failed staging operation may already have removed the file.
        }
    }

    private void LoadWorkbookStructure()
    {
        byte[]? relationships = _package.TryGetPart("xl/_rels/workbook.xml.rels");
        if (relationships is not null)
        {
            var targets = UpdaterXmlUtils.ReadRelationships(relationships, "xl/workbook.xml");
            foreach (var pair in targets)
            {
                _relationshipTargets[pair.Key] = pair.Value;
                if (pair.Value.EndsWith("sharedStrings.xml", StringComparison.OrdinalIgnoreCase))
                    _sharedStringsPath = pair.Value;
                else if (pair.Value.EndsWith("styles.xml", StringComparison.OrdinalIgnoreCase))
                    _stylesPath = pair.Value;
            }
        }

        var workbook = UpdaterXmlUtils.LoadDocument(_package.GetPart("xl/workbook.xml"));
        var workbookProperties = UpdaterXmlUtils.FirstDescendant(workbook, "workbookPr");
        string date1904 = workbookProperties is null ? string.Empty : UpdaterXmlUtils.GetAttribute(workbookProperties, "date1904");
        _uses1904DateSystem = string.Equals(date1904, "1", StringComparison.Ordinal) ||
            string.Equals(date1904, "true", StringComparison.OrdinalIgnoreCase);
        foreach (var sheet in UpdaterXmlUtils.Descendants(workbook, "sheet"))
        {
            string name = UpdaterXmlUtils.GetAttribute(sheet, "name");
            string relationshipId = UpdaterXmlUtils.GetAttribute(sheet, "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            if (string.IsNullOrEmpty(relationshipId))
                relationshipId = UpdaterXmlUtils.GetAttribute(sheet, "id");

            if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(relationshipId) ||
                !_relationshipTargets.TryGetValue(relationshipId, out string? location))
            {
                continue;
            }

            if (_sheetNameToPath.ContainsKey(name))
                throw new InvalidDataException($"The XLSX workbook contains duplicate worksheet name '{name}'.");
            _sheetNameToPath[name] = location;
            _sheetNames.Add(name);
        }

        foreach (string partName in _package.PartNames)
        {
            if (Regex.IsMatch(partName, @"^xl/pivotCache/pivotCacheDefinition\d*\.xml$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                _pivotCachePaths.Add(partName);
            else if (Regex.IsMatch(partName, @"^xl/pivotTables/pivotTable\d*\.xml$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                _pivotTablePaths.Add(partName);
        }

        if (_sheetNames.Count == 0)
            throw new InvalidDataException("The XLSX workbook contains no resolvable worksheets.");
    }

    private void EnsureSharedStringsLoaded()
    {
        if (_sharedStrings is not null || _sharedStringsPath is null)
            return;

        byte[]? bytes = _package.TryGetPart(_sharedStringsPath);
        if (bytes is null)
            return;

        var state = new SharedStringsState
        {
            Xml = Encoding.UTF8.GetString(bytes)
        };
        state.Values.AddRange(UpdaterXmlUtils.ParseSharedStringsXml(bytes));
        for (int i = 0; i < state.Values.Count; i++)
            state.Index.TryAdd(state.Values[i], i);

        var document = UpdaterXmlUtils.LoadDocument(bytes);
        var root = UpdaterXmlUtils.FirstDescendant(document, "sst");
        if (root is not null && int.TryParse(UpdaterXmlUtils.GetAttribute(root, "count"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int count))
            state.TotalCount = count;
        else
            state.TotalCount = 0;

        state.CommittedValueCount = state.Values.Count;
        _sharedStrings = state;
    }

    private int CountSharedStringReferences(string sheetXml)
    {
        if (_sharedStrings is null)
            return 0;

        var document = UpdaterXmlUtils.LoadDocument(Encoding.UTF8.GetBytes(sheetXml));
        var sheetData = UpdaterXmlUtils.FirstDescendant(document, "sheetData");
        if (sheetData is null)
            return 0;

        int count = 0;
        foreach (var cell in UpdaterXmlUtils.Descendants(sheetData, "c"))
        {
            if (UpdaterXmlUtils.GetAttribute(cell, "t") == "s")
                count++;
        }
        return count;
    }

    private void CommitSharedStrings()
    {
        if (_sharedStrings is null || !_sharedStrings.Dirty || _sharedStringsPath is null)
            return;

        string xml = _sharedStrings.Xml;
        Match rootMatch = UpdaterXmlUtils.FindStartTag(xml, "sst");
        string rootTag = rootMatch.Value;
        rootTag = UpdaterXmlUtils.ReplaceTagAttribute(rootTag, "count", _sharedStrings.TotalCount.ToString(CultureInfo.InvariantCulture));
        rootTag = UpdaterXmlUtils.ReplaceTagAttribute(rootTag, "uniqueCount", _sharedStrings.Values.Count.ToString(CultureInfo.InvariantCulture));
        xml = xml[..rootMatch.Index] + rootTag + xml[(rootMatch.Index + rootMatch.Length)..];

        Match updatedRoot = UpdaterXmlUtils.FindStartTag(xml, "sst");
        string prefix = UpdaterXmlUtils.PrefixFromTag(updatedRoot);
        string closing = $"</{prefix}sst>";
        int end = xml.LastIndexOf(closing, StringComparison.Ordinal);
        if (end < 0)
            throw new InvalidDataException("The XLSX Shared Strings part has no closing sst element.");

        var appended = new StringBuilder();
        for (int i = _sharedStrings.CommittedValueCount; i < _sharedStrings.Values.Count; i++)
        {
            string item = UpdaterXmlUtils.BuildSharedStringItem(_sharedStrings.Values[i]);
            if (!string.IsNullOrEmpty(prefix))
                item = item.Replace("<si>", $"<{prefix}si>", StringComparison.Ordinal)
                    .Replace("</si>", $"</{prefix}si>", StringComparison.Ordinal)
                    .Replace("<t", $"<{prefix}t", StringComparison.Ordinal)
                    .Replace("</t>", $"</{prefix}t>", StringComparison.Ordinal);
            appended.Append(item);
        }

        xml = xml[..end] + appended + xml[end..];
        _sharedStrings.Xml = xml;
        _sharedStrings.CommittedValueCount = _sharedStrings.Values.Count;
        _sharedStrings.Dirty = false;
        _package.SetPart(_sharedStringsPath, Encoding.UTF8.GetBytes(xml));
    }

    private (Dictionary<int, int> dataStyles, Dictionary<int, int> headerStyles) CollectExistingStyles(string sheetXml)
    {
        var dataCounts = new Dictionary<int, Dictionary<int, int>>();
        var headerStyles = new Dictionary<int, int>();
        var document = UpdaterXmlUtils.LoadDocument(Encoding.UTF8.GetBytes(sheetXml));
        var sheetData = UpdaterXmlUtils.FirstDescendant(document, "sheetData");
        if (sheetData is null)
            return ([], headerStyles);

        foreach (var cell in UpdaterXmlUtils.Descendants(sheetData, "c"))
        {
            string reference = UpdaterXmlUtils.GetAttribute(cell, "r");
            if (!UpdaterXmlUtils.TryParseCellReference(reference, out int column, out int row))
                continue;
            string styleText = UpdaterXmlUtils.GetAttribute(cell, "s");
            if (!int.TryParse(styleText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int style) || style < 0)
                style = 0;

            if (row == 1)
            {
                if (!headerStyles.ContainsKey(column))
                    headerStyles[column] = style;
                continue;
            }

            if (style <= 0)
                continue;
            if (!dataCounts.TryGetValue(column, out var styles))
            {
                styles = new Dictionary<int, int>();
                dataCounts[column] = styles;
            }
            styles[style] = styles.GetValueOrDefault(style) + 1;
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
        byte[]? bytes = _package.TryGetPart(_stylesPath);
        if (bytes is null)
            return null;

        var document = UpdaterXmlUtils.LoadDocument(bytes);
        var customDates = new HashSet<int>();
        var numberFormats = UpdaterXmlUtils.FirstDescendant(document, "numFmts");
        if (numberFormats is not null)
        {
            foreach (XmlNode child in numberFormats.ChildNodes)
            {
                if (child is not XmlElement format || format.LocalName != "numFmt")
                    continue;
                if (!int.TryParse(UpdaterXmlUtils.GetAttribute(format, "numFmtId"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int id))
                    continue;
                string code = UpdaterXmlUtils.GetAttribute(format, "formatCode").ToLowerInvariant();
                if (code.Contains("yy", StringComparison.Ordinal) || code.Contains("mm", StringComparison.Ordinal) ||
                    code.Contains("dd", StringComparison.Ordinal) || code.Contains("h:mm", StringComparison.Ordinal))
                {
                    customDates.Add(id);
                }
            }
        }

        var cellXfs = UpdaterXmlUtils.FirstDescendant(document, "cellXfs");
        if (cellXfs is null)
            return null;

        int index = 0;
        foreach (XmlNode child in cellXfs.ChildNodes)
        {
            if (child is not XmlElement xf || xf.LocalName != "xf")
                continue;
            _ = int.TryParse(UpdaterXmlUtils.GetAttribute(xf, "numFmtId"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int numFmtId);
            bool isDate = (numFmtId >= 14 && numFmtId <= 22) ||
                          (numFmtId >= 45 && numFmtId <= 47) || customDates.Contains(numFmtId);
            if (isDate)
                return index;
            index++;
        }
        return null;
    }

    private string BuildSheetData(
        UpdaterRows data,
        IReadOnlyDictionary<int, int> dataStyles,
        IReadOnlyDictionary<int, int> headerStyles,
        int? dateStyle,
        out string elementPrefix)
    {
        // The caller supplies the prefix from the worksheet's sheetData tag. The
        // normal Open XML form uses the default namespace and therefore an empty prefix.
        elementPrefix = string.Empty;
        var builder = new StringBuilder();
        int rowNumber = 1;

        if (data.Headers is not null)
        {
            builder.Append("<row r=\"1\">");
            for (int column = 0; column < data.Headers.Count; column++)
                builder.Append(BuildStringCell(data.Headers[column], column, 1, headerStyles.GetValueOrDefault(column), string.Empty));
            builder.Append("</row>");
            rowNumber++;
        }

        foreach (var row in data.Rows)
        {
            builder.Append("<row r=\"").Append(rowNumber).Append("\">");
            for (int column = 0; column < row.Length; column++)
            {
                object? value = UpdaterRows.Unwrap(row[column]);
                if (UpdaterRows.IsEmpty(value))
                    continue;
                builder.Append(BuildValueCell(value!, column, rowNumber, dataStyles.GetValueOrDefault(column), dateStyle, string.Empty));
            }
            builder.Append("</row>");
            rowNumber++;
        }

        return builder.ToString();
    }

    private string BuildValueCell(object value, int column, int row, int style, int? dateStyle, string prefix)
    {
        string reference = $"{UpdaterXmlUtils.ColumnToLetters(column)}{row}";
        string styleAttribute = style > 0 ? $" s=\"{style}\"" : string.Empty;
        value = UpdaterRows.Unwrap(value)!;

        if (value is DateTime dateTime)
        {
            try
            {
                double serial = ToExcelDateSerial(dateTime);
                if (double.IsFinite(serial))
                {
                    int dateCellStyle = style > 0 ? style : dateStyle ?? 0;
                    string dateStyleAttribute = dateCellStyle > 0 ? $" s=\"{dateCellStyle}\"" : string.Empty;
                    return $"<{prefix}c r=\"{reference}\"{dateStyleAttribute}><{prefix}v>{serial.ToString("R", CultureInfo.InvariantCulture)}</{prefix}v></{prefix}c>";
                }
            }
            catch (ArgumentException)
            {
                // Fall through and store an unrepresentable date as text.
            }
            return BuildStringCell(dateTime.ToString(CultureInfo.InvariantCulture), column, row, style, prefix);
        }

        if (value is DateTimeOffset dateTimeOffset)
            return BuildValueCell(dateTimeOffset.DateTime, column, row, style, dateStyle, prefix);

        if (value is bool boolean)
            return $"<{prefix}c r=\"{reference}\" t=\"b\"{styleAttribute}><{prefix}v>{(boolean ? 1 : 0)}</{prefix}v></{prefix}c>";

        if (TryConvertNumber(value, out double number))
        {
            if (double.IsFinite(number))
                return $"<{prefix}c r=\"{reference}\"{styleAttribute}><{prefix}v>{number.ToString("R", CultureInfo.InvariantCulture)}</{prefix}v></{prefix}c>";
            return BuildStringCell(number.ToString(CultureInfo.InvariantCulture), column, row, style, prefix);
        }

        string text = value switch
        {
            string stringValue => stringValue,
            char character => character.ToString(),
            byte[] bytes => Encoding.UTF8.GetString(bytes),
            Memory<byte> memory => Encoding.UTF8.GetString(memory.Span),
            _ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty
        };
        return BuildStringCell(text, column, row, style, prefix);
    }

    private double ToExcelDateSerial(DateTime value)
    {
        const double DateSystemOffset = 1462d;
        return value.ToOADate() - (_uses1904DateSystem ? DateSystemOffset : 0d);
    }

    private string BuildStringCell(string value, int column, int row, int style, string prefix)
    {
        string reference = $"{UpdaterXmlUtils.ColumnToLetters(column)}{row}";
        string styleAttribute = style > 0 ? $" s=\"{style}\"" : string.Empty;
        string cellPrefix = prefix;
        if (_sharedStrings is null)
        {
            string preserve = UpdaterXmlUtils.ShouldPreserveWhitespace(value) ? " xml:space=\"preserve\"" : string.Empty;
            string escaped = UpdaterXmlUtils.EscapeXmlText(value);
            return $"<{cellPrefix}c r=\"{reference}\" t=\"inlineStr\"{styleAttribute}><{cellPrefix}is><{cellPrefix}t{preserve}>{escaped}</{cellPrefix}t></{cellPrefix}is></{cellPrefix}c>";
        }

        if (!_sharedStrings.Index.TryGetValue(value, out int index))
        {
            index = _sharedStrings.Values.Count;
            _sharedStrings.Values.Add(value);
            _sharedStrings.Index[value] = index;
        }
        _sharedStrings.TotalCount++;
        _sharedStrings.Dirty = true;
        return $"<{cellPrefix}c r=\"{reference}\" t=\"s\"{styleAttribute}><{cellPrefix}v>{index}</{cellPrefix}v></{cellPrefix}c>";
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

    private string PatchWorksheetXml(string xml, string newSheetData, int totalRows, int lastColumn, string unusedPrefix)
    {
        string dimension = totalRows > 0 && lastColumn >= 0
            ? $"A1:{UpdaterXmlUtils.ColumnToLetters(lastColumn)}{totalRows}"
            : "A1";

        Match root = UpdaterXmlUtils.FindStartTag(xml, "worksheet");
        Match? dimensionTag = FindOptionalStartTag(xml, "dimension");
        if (dimensionTag is not null)
        {
            string replacement = UpdaterXmlUtils.ReplaceTagAttribute(dimensionTag.Value, "ref", dimension);
            xml = xml[..dimensionTag.Index] + replacement + xml[(dimensionTag.Index + dimensionTag.Length)..];
        }
        else
        {
            root = UpdaterXmlUtils.FindStartTag(xml, "worksheet");
            string prefix = UpdaterXmlUtils.PrefixFromTag(root);
            string dimensionElement = $"<{prefix}dimension ref=\"{dimension}\"/>";
            int insertion = root.Index + root.Length;
            xml = xml[..insertion] + dimensionElement + xml[insertion..];
        }

        Match sheetDataOpen = UpdaterXmlUtils.FindStartTag(xml, "sheetData");
        string prefixFromSheetData = UpdaterXmlUtils.PrefixFromTag(sheetDataOpen);
        string prefixedData = PrefixSheetData(newSheetData, prefixFromSheetData);
        if (sheetDataOpen.Value.TrimEnd().EndsWith("/>", StringComparison.Ordinal))
        {
            string openTag = sheetDataOpen.Value[..^2] + ">";
            string closeTag = $"</{prefixFromSheetData}sheetData>";
            xml = xml[..sheetDataOpen.Index] + openTag + prefixedData + closeTag + xml[(sheetDataOpen.Index + sheetDataOpen.Length)..];
        }
        else
        {
            string closeTag = $"</{prefixFromSheetData}sheetData>";
            int contentStart = sheetDataOpen.Index + sheetDataOpen.Length;
            int contentEnd = xml.IndexOf(closeTag, contentStart, StringComparison.Ordinal);
            if (contentEnd < 0)
                throw new InvalidDataException("The XLSX worksheet sheetData element is not closed.");
            xml = xml[..contentStart] + prefixedData + xml[contentEnd..];
        }

        Match? autoFilter = FindOptionalStartTag(xml, "autoFilter");
        if (autoFilter is not null)
        {
            string replacement = UpdaterXmlUtils.ReplaceTagAttribute(autoFilter.Value, "ref", dimension);
            xml = xml[..autoFilter.Index] + replacement + xml[(autoFilter.Index + autoFilter.Length)..];
        }
        return xml;
    }

    private static string PrefixSheetData(string sheetData, string prefix)
    {
        if (string.IsNullOrEmpty(prefix) || sheetData.Length == 0)
            return sheetData;
        return sheetData.Replace("<row ", $"<{prefix}row ", StringComparison.Ordinal)
            .Replace("</row>", $"</{prefix}row>", StringComparison.Ordinal)
            .Replace("<c ", $"<{prefix}c ", StringComparison.Ordinal)
            .Replace("</c>", $"</{prefix}c>", StringComparison.Ordinal)
            .Replace("><v>", $"><{prefix}v>", StringComparison.Ordinal)
            .Replace("</v>", $"</{prefix}v>", StringComparison.Ordinal)
            .Replace("><is>", $"><{prefix}is>", StringComparison.Ordinal)
            .Replace("</is>", $"</{prefix}is>", StringComparison.Ordinal)
            .Replace("><t", $"><{prefix}t", StringComparison.Ordinal)
            .Replace("</t>", $"</{prefix}t>", StringComparison.Ordinal);
    }

    private static Match? FindOptionalStartTag(string xml, string localName)
    {
        var regex = new Regex($"<(?<prefix>[A-Za-z_][A-Za-z0-9_.-]*:)?{Regex.Escape(localName)}\\b[^>]*>",
            RegexOptions.CultureInvariant | RegexOptions.Singleline);
        Match match = regex.Match(xml);
        return match.Success ? match : null;
    }

    private HashSet<string> UpdatePivotCaches(string sheetName, string dimension, int recordCount)
    {
        var affectedCacheIds = new HashSet<string>(StringComparer.Ordinal);
        bool affectedWithoutId = false;

        foreach (string path in _pivotCachePaths)
        {
            byte[] bytes = _package.GetPart(path);
            var document = UpdaterXmlUtils.LoadDocument(bytes);
            var root = UpdaterXmlUtils.FirstDescendant(document, "pivotCacheDefinition");
            if (root is null)
                continue;

            bool affected = false;
            foreach (var source in UpdaterXmlUtils.Descendants(document, "worksheetSource"))
            {
                if (!string.Equals(UpdaterXmlUtils.GetAttribute(source, "sheet"), sheetName, StringComparison.Ordinal))
                    continue;
                UpdaterXmlUtils.SetAttribute(source, "ref", dimension);
                affected = true;
            }
            if (!affected)
                continue;

            UpdaterXmlUtils.SetAttribute(root, "recordCount", recordCount.ToString(CultureInfo.InvariantCulture));
            UpdaterXmlUtils.SetAttribute(root, "refreshOnLoad", "1");
            string cacheId = UpdaterXmlUtils.GetAttribute(root, "cacheId");
            if (!string.IsNullOrEmpty(cacheId))
                affectedCacheIds.Add(cacheId);
            else
                affectedWithoutId = true;
            _package.SetPart(path, UpdaterXmlUtils.SaveDocument(document));
        }

        if (affectedWithoutId)
            affectedCacheIds.Add(string.Empty);
        return affectedCacheIds;
    }

    private void UpdatePivotTables(HashSet<string> affectedCacheIds)
    {
        if (affectedCacheIds.Count == 0)
            return;

        bool refreshAll = affectedCacheIds.Contains(string.Empty);
        foreach (string path in _pivotTablePaths)
        {
            byte[] bytes = _package.GetPart(path);
            var document = UpdaterXmlUtils.LoadDocument(bytes);
            var root = UpdaterXmlUtils.FirstDescendant(document, "pivotTableDefinition");
            if (root is null)
                continue;
            string cacheId = UpdaterXmlUtils.GetAttribute(root, "cacheId");
            if (!refreshAll && !affectedCacheIds.Contains(cacheId))
                continue;
            UpdaterXmlUtils.SetAttribute(root, "refreshOnLoad", "1");
            _package.SetPart(path, UpdaterXmlUtils.SaveDocument(document));
        }
    }

    private void ThrowIfDisposed()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
    }
}
