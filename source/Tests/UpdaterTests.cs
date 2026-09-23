using SpreadSheetTasks;
using System.Buffers.Binary;
using System.Data;
using System.IO.Compression;
using System.Reflection;
using System.Xml.Linq;

namespace Tests;

[Collection("Sequential")]
public class UpdaterTests
{
    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void ReplaceSheetData_RoundTripsThroughExistingReader(string extension)
    {
        string source = SylvanInteropTestHelpers.CreateTempExcelPath(extension);
        string output = SylvanInteropTestHelpers.CreateTempExcelPath(extension);
        try
        {
            using (var writer = ExcelWriter.CreateWriter(source))
            {
                writer.AddSheet("Data");
                writer.WriteSheet(new object?[][]
                {
                    ["Old", 1],
                    ["Old2", 2]
                }, ["Name", "Amount"]);
                writer.AddSheet("Other");
                writer.WriteSheet(new object?[][] { ["Keep"] }, headers: false);
            }

            var rows = new object?[][]
            {
                ["Alice", 10, "2025-02-03", true],
                [],
                ["Bob & <x>", null, "2025-03-04", false],
                [null, null, null, null]
            };

            {
                if (extension == ".xlsx")
                {
                    using var xlsx = new XlsxUpdater(source);
                    Assert.Equal(["Data", "Other"], xlsx.GetSheetNames());
                    xlsx.ReplaceSheetData("Data", rows, new ReplaceSheetDataOptions
                    {
                        Headers = ["Name", "Amount", "When", "Active"]
                    });
                    xlsx.Save(output);
                }
                else
                {
                    using var xlsb = new XlsbUpdater(source);
                    Assert.Equal(["Data", "Other"], xlsb.GetSheetNames());
                    xlsb.ReplaceSheetData("Data", rows, new ReplaceSheetDataOptions
                    {
                        Headers = ["Name", "Amount", "When", "Active"]
                    });
                    xlsb.Save(output);
                }
            }

            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.Open(output);
                reader.ActualSheetName = "Data";
                Assert.True(reader.Read());
                Assert.Equal("Name", reader.GetValue(0));
                Assert.Equal("Active", reader.GetValue(3));
                Assert.True(reader.Read());
                Assert.Equal("Alice", reader.GetValue(0));
                Assert.Equal(10L, reader.GetValue(1));
                Assert.Equal("2025-02-03", reader.GetValue(2));
                Assert.Equal(true, reader.GetValue(3));
                Assert.True(reader.Read());
                Assert.All(Enumerable.Range(0, 4), column => Assert.Equal(DBNull.Value, reader.GetValue(column)));
                Assert.True(reader.Read());
                Assert.Equal("Bob & <x>", reader.GetValue(0));
                Assert.Equal(DBNull.Value, reader.GetValue(1));
                Assert.Equal(false, reader.GetValue(3));
                Assert.False(reader.Read());
            }

            using var otherReader = new XlsxOrXlsbReadOrEdit();
            otherReader.Open(output);
            otherReader.ActualSheetName = "Other";
            Assert.True(otherReader.Read());
            Assert.Equal("Keep", otherReader.GetValue(0));
        }
        finally
        {
            DeleteIfExists(source);
            DeleteIfExists(output);
        }
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void ReplaceSheetData_ToArrayAndGeneralFallback_DoNotChangeSource(string extension)
    {
        string source = SylvanInteropTestHelpers.CreateTempExcelPath(extension);
        string output = SylvanInteropTestHelpers.CreateTempExcelPath(extension);
        try
        {
            using (var writer = ExcelWriter.CreateWriter(source))
            {
                writer.AddSheet("Data");
                writer.WriteSheet(new object?[][] { ["Original", 1] }, ["Text", "Number"]);
                writer.AddSheet("Other");
                writer.WriteSheet(new object?[][] { ["Keep"] }, headers: false);
            }

            byte[] originalSheet = ReadPart(source, extension == ".xlsx" ? "xl/worksheets/sheet1.xml" : "xl/worksheets/sheet1.bin");
            byte[] originalOtherSheet = ReadPart(source, extension == ".xlsx" ? "xl/worksheets/sheet2.xml" : "xl/worksheets/sheet2.bin");
            byte[] originalWorkbook = ReadPart(source, extension == ".xlsx" ? "xl/workbook.xml" : "xl/workbook.bin");
            byte[] packageBytes;
            if (extension == ".xlsx")
            {
                using var updater = new XlsxUpdater(source);
                updater.ReplaceSheetData("Data", new object?[][] { ["Changed", 2.5] },
                    new ReplaceSheetDataOptions { Headers = ["Text", "Number"], StyleFallback = ReplaceSheetDataStyleFallback.General });
                packageBytes = updater.ToArray();
            }
            else
            {
                using var updater = new XlsbUpdater(source);
                updater.ReplaceSheetData("Data", new object?[][] { ["Changed", 2.5] },
                    new ReplaceSheetDataOptions { Headers = ["Text", "Number"], StyleFallback = ReplaceSheetDataStyleFallback.General });
                packageBytes = updater.ToArray();
            }

            File.WriteAllBytes(output, packageBytes);
            Assert.Equal(originalSheet, ReadPart(source, extension == ".xlsx" ? "xl/worksheets/sheet1.xml" : "xl/worksheets/sheet1.bin"));
            Assert.NotEqual(originalSheet, ReadPart(output, extension == ".xlsx" ? "xl/worksheets/sheet1.xml" : "xl/worksheets/sheet1.bin"));
            Assert.Equal(originalOtherSheet, ReadPart(output, extension == ".xlsx" ? "xl/worksheets/sheet2.xml" : "xl/worksheets/sheet2.bin"));
            Assert.Equal(originalWorkbook, ReadPart(output, extension == ".xlsx" ? "xl/workbook.xml" : "xl/workbook.bin"));
        }
        finally
        {
            DeleteIfExists(source);
            DeleteIfExists(output);
        }
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void ReplaceSheetData_AcceptsDataTable(string extension)
    {
        string source = SylvanInteropTestHelpers.CreateTempExcelPath(extension);
        string output = SylvanInteropTestHelpers.CreateTempExcelPath(extension);
        try
        {
            using (var writer = ExcelWriter.CreateWriter(source))
            {
                writer.AddSheet("Data");
                writer.WriteSheet(new object?[][] { ["old", 0] }, headers: ["Text", "Value"]);
            }

            var table = new DataTable();
            table.Columns.Add("Text", typeof(string));
            table.Columns.Add("Value", typeof(int));
            table.Rows.Add("new", 42);

            if (extension == ".xlsx")
            {
                using var updater = new XlsxUpdater(source);
                updater.ReplaceSheetData("Data", table, new ReplaceSheetDataOptions { Headers = ["Text", "Value"] });
                updater.Save(output);
            }
            else
            {
                using var updater = new XlsbUpdater(source);
                updater.ReplaceSheetData("Data", table, new ReplaceSheetDataOptions { Headers = ["Text", "Value"] });
                updater.Save(output);
            }

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(output);
            reader.ActualSheetName = "Data";
            Assert.True(reader.Read());
            Assert.Equal("Text", reader.GetValue(0));
            Assert.True(reader.Read());
            Assert.Equal("new", reader.GetValue(0));
            Assert.Equal(42L, reader.GetValue(1));
            Assert.False(reader.Read());
        }
        finally
        {
            DeleteIfExists(source);
            DeleteIfExists(output);
        }
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void ReplaceSheetData_ReaderFailureCanBeRetried(string extension)
    {
        string source = SylvanInteropTestHelpers.CreateTempExcelPath(extension);
        string output = SylvanInteropTestHelpers.CreateTempExcelPath(extension);
        string arrayOutput = SylvanInteropTestHelpers.CreateTempExcelPath(extension);
        try
        {
            using (var writer = ExcelWriter.CreateWriter(source))
            {
                writer.AddSheet("Data");
                writer.WriteSheet(new object?[][] { ["old"] }, ["Value"]);
            }

            var failedTable = new DataTable();
            failedTable.Columns.Add("Value", typeof(string));
            int rowsBeforeFailure = extension == ".xlsb" ? 16_385 : 4_097;
            for (int i = 0; i <= rowsBeforeFailure; i++)
                failedTable.Rows.Add("discarded before failure");

            var retryTable = new DataTable();
            retryTable.Columns.Add("Value", typeof(string));
            retryTable.Rows.Add("retry succeeded");

            var secondReplacementTable = new DataTable();
            secondReplacementTable.Columns.Add("Value", typeof(string));
            secondReplacementTable.Rows.Add("second replacement wins");

            if (extension == ".xlsx")
            {
                using var updater = new XlsxUpdater(source);
                using (IDataReader inner = failedTable.CreateDataReader())
                using (IDataReader failing = ThrowAfterReads(inner, readsBeforeThrow: rowsBeforeFailure))
                {
                    Assert.Throws<InvalidOperationException>(() => updater.ReplaceSheetData("Data", failing,
                        new ReplaceSheetDataOptions { Headers = ["Value"] }));
                }
                using (IDataReader reader = failedTable.CreateDataReader())
                    updater.ReplaceSheetData("Data", reader, new ReplaceSheetDataOptions { Headers = ["Value"] });
                updater.Save(output);
                File.WriteAllBytes(arrayOutput, updater.ToArray());
                AssertStreamedRows(arrayOutput, failedTable.Rows.Count);

                using (IDataReader reader = retryTable.CreateDataReader())
                    updater.ReplaceSheetData("Data", reader, new ReplaceSheetDataOptions { Headers = ["Value"] });
                using (IDataReader reader = secondReplacementTable.CreateDataReader())
                    updater.ReplaceSheetData("Data", reader, new ReplaceSheetDataOptions { Headers = ["Value"] });
                updater.Save(output);
            }
            else
            {
                using var updater = new XlsbUpdater(source);
                using (IDataReader inner = failedTable.CreateDataReader())
                using (IDataReader failing = ThrowAfterReads(inner, readsBeforeThrow: rowsBeforeFailure))
                {
                    Assert.Throws<InvalidOperationException>(() => updater.ReplaceSheetData("Data", failing,
                        new ReplaceSheetDataOptions { Headers = ["Value"] }));
                }
                using (IDataReader reader = failedTable.CreateDataReader())
                    updater.ReplaceSheetData("Data", reader, new ReplaceSheetDataOptions { Headers = ["Value"] });
                updater.Save(output);
                File.WriteAllBytes(arrayOutput, updater.ToArray());
                AssertStreamedRows(arrayOutput, failedTable.Rows.Count);

                using (IDataReader reader = retryTable.CreateDataReader())
                    updater.ReplaceSheetData("Data", reader, new ReplaceSheetDataOptions { Headers = ["Value"] });
                using (IDataReader reader = secondReplacementTable.CreateDataReader())
                    updater.ReplaceSheetData("Data", reader, new ReplaceSheetDataOptions { Headers = ["Value"] });
                updater.Save(output);
            }

            using var result = new XlsxOrXlsbReadOrEdit();
            result.Open(output);
            result.ActualSheetName = "Data";
            Assert.True(result.Read());
            Assert.Equal("Value", result.GetValue(0));
            Assert.True(result.Read());
            Assert.Equal("second replacement wins", result.GetValue(0));
            Assert.False(result.Read());
        }
        finally
        {
            DeleteIfExists(source);
            DeleteIfExists(output);
            DeleteIfExists(arrayOutput);
        }
    }

    [Fact]
    public void ReplaceSheetData_HandlesPrefixedSelfClosingXlsxSheetData()
    {
        string source = SylvanInteropTestHelpers.CreateTempExcelPath(".xlsx");
        string output = SylvanInteropTestHelpers.CreateTempExcelPath(".xlsx");
        try
        {
            using (var writer = new XlsxWriter(source))
            {
                writer.AddSheet("Data");
                writer.WriteSheet(new object?[][] { ["old", 1] }, ["Text", "Value"]);
            }

            string worksheetPath = "xl/worksheets/sheet1.xml";
            string xml = System.Text.Encoding.UTF8.GetString(ReadPart(source, worksheetPath));
            int start = xml.IndexOf("<sheetData", StringComparison.Ordinal);
            int openEnd = start < 0 ? -1 : xml.IndexOf('>', start);
            int closeStart = openEnd < 0 ? -1 : xml.IndexOf("</sheetData>", openEnd + 1, StringComparison.Ordinal);
            Assert.True(start >= 0 && openEnd >= 0 && closeStart >= 0, "The test workbook should contain sheetData.");
            xml = string.Concat(
                xml.AsSpan(0, start),
                "<x:sheetData xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>",
                xml.AsSpan(xml.IndexOf("</sheetData>", closeStart, StringComparison.Ordinal) + "</sheetData>".Length));
            WritePart(source, worksheetPath, System.Text.Encoding.UTF8.GetBytes(xml));

            var table = new DataTable();
            table.Columns.Add("Text", typeof(string));
            table.Columns.Add("Value", typeof(int));
            table.Rows.Add("updated", 42);

            using (var updater = new XlsxUpdater(source))
            using (IDataReader reader = table.CreateDataReader())
            {
                updater.ReplaceSheetData("Data", reader, new ReplaceSheetDataOptions { Headers = ["Text", "Value"] });
                updater.Save(output);
            }

            using var worksheetXml = new MemoryStream(ReadPart(output, worksheetPath));
            var document = XDocument.Load(worksheetXml);
            XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var sheetData = document.Descendants(spreadsheet + "sheetData").Single();
            var rows = sheetData.Elements(spreadsheet + "row").ToArray();
            Assert.Equal(2, rows.Length);
            Assert.Equal("A1", (string?)rows[0].Elements(spreadsheet + "c").First().Attribute("r"));
            Assert.Equal("A2", (string?)rows[1].Elements(spreadsheet + "c").First().Attribute("r"));
            var numericCell = rows[1].Elements(spreadsheet + "c")
                .Single(cell => (string?)cell.Attribute("r") == "B2");
            Assert.Equal("42", numericCell.Element(spreadsheet + "v")?.Value);
        }
        finally
        {
            DeleteIfExists(source);
            DeleteIfExists(output);
        }
    }

    [Fact]
    public void UpdatersRejectUnsupportedFileExtensions()
    {
        string source = SylvanInteropTestHelpers.CreateTempExcelPath(".xlsx");
        string renamed = Path.Combine(Path.GetDirectoryName(source)!, Guid.NewGuid().ToString("N") + ".csv");
        try
        {
            using (var writer = new XlsxWriter(source))
            {
                writer.AddSheet("Data");
                writer.WriteSheet(new object?[][] { ["value"] }, headers: false);
            }
            File.Copy(source, renamed);

            Assert.Throws<NotSupportedException>(() => new XlsxUpdater(renamed));
            Assert.Throws<NotSupportedException>(() => new XlsbUpdater(renamed));
        }
        finally
        {
            DeleteIfExists(source);
            DeleteIfExists(renamed);
        }
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void ReplaceSheetData_UsesWorkbook1904DateSystem(string extension)
    {
        string source = SylvanInteropTestHelpers.CreateTempExcelPath(extension);
        string output = SylvanInteropTestHelpers.CreateTempExcelPath(extension);
        DateTime date = new(2025, 2, 3, 12, 30, 0);
        try
        {
            using (var writer = ExcelWriter.CreateWriter(source))
            {
                writer.AddSheet("Data");
                writer.WriteSheet(new object?[][] { [new DateTime(2024, 1, 1)] }, ["Date"]);
            }
            SetWorkbook1904DateSystem(source, extension);

            if (extension == ".xlsx")
            {
                using var updater = new XlsxUpdater(source);
                updater.ReplaceSheetData("Data", [[date]]);
                updater.Save(output);
            }
            else
            {
                using var updater = new XlsbUpdater(source);
                updater.ReplaceSheetData("Data", [[date]]);
                updater.Save(output);
            }

            Assert.Equal(date.ToOADate() - 1462d, ReadFirstDateSerial(output, extension), 10);
        }
        finally
        {
            DeleteIfExists(source);
            DeleteIfExists(output);
        }
    }

    private static byte[] ReadPart(string path, string partName)
    {
        using var archive = ZipFile.OpenRead(path);
        var entry = archive.GetEntry(partName);
        Assert.NotNull(entry);
        using var stream = entry!.Open();
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }

    private static void WritePart(string path, string partName, byte[] data)
    {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        archive.GetEntry(partName)!.Delete();
        var entry = archive.CreateEntry(partName);
        using var output = entry.Open();
        output.Write(data, 0, data.Length);
    }

    private static void AssertStreamedRows(string path, int expectedDataRows)
    {
        using var result = new XlsxOrXlsbReadOrEdit();
        result.Open(path);
        result.ActualSheetName = "Data";
        Assert.True(result.Read());
        Assert.Equal("Value", result.GetValue(0));
        Assert.True(result.Read());
        Assert.Equal("discarded before failure", result.GetValue(0));

        int rows = 1;
        while (result.Read())
            rows++;
        Assert.Equal(expectedDataRows, rows);
    }

    private static void SetWorkbook1904DateSystem(string path, string extension)
    {
        string partName = extension == ".xlsx" ? "xl/workbook.xml" : "xl/workbook.bin";
        byte[] part = ReadPart(path, partName);
        if (extension == ".xlsx")
        {
            string xml = System.Text.Encoding.UTF8.GetString(part);
            part = System.Text.Encoding.UTF8.GetBytes(xml.Replace("<workbookPr ", "<workbookPr date1904=\"1\" ", StringComparison.Ordinal));
        }
        else
        {
            int recordDataOffset = FindBiffRecordDataOffset(part, 0x0099);
            part[recordDataOffset] |= 1;
        }

        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        archive.GetEntry(partName)!.Delete();
        var entry = archive.CreateEntry(partName);
        using var stream = entry.Open();
        stream.Write(part, 0, part.Length);
    }

    private static double ReadFirstDateSerial(string path, string extension)
    {
        byte[] part = ReadPart(path, extension == ".xlsx" ? "xl/worksheets/sheet1.xml" : "xl/worksheets/sheet1.bin");
        if (extension == ".xlsx")
        {
            string xml = System.Text.Encoding.UTF8.GetString(part);
            var match = System.Text.RegularExpressions.Regex.Match(xml, "<c r=\"A1\"[^>]*><v>([^<]+)</v>");
            Assert.True(match.Success, "Updated XLSX date cell was not found.");
            return double.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
        }

        int recordDataOffset = FindBiffRecordDataOffset(part, 0x0005);
        return BinaryPrimitives.ReadDoubleLittleEndian(part.AsSpan(recordDataOffset + 8));
    }

    private static int FindBiffRecordDataOffset(ReadOnlySpan<byte> data, uint soughtId)
    {
        int position = 0;
        while (position < data.Length)
        {
            uint id = ReadBiffVlq(data, ref position);
            uint length = ReadBiffVlq(data, ref position);
            Assert.InRange(length, 0u, (uint)(data.Length - position));
            if (id == soughtId)
                return position;
            position += checked((int)length);
        }
        throw new InvalidDataException($"BIFF record 0x{soughtId:X} was not found.");
    }

    private static uint ReadBiffVlq(ReadOnlySpan<byte> data, ref int position)
    {
        uint value = 0;
        for (int shift = 0; shift < 35; shift += 7)
        {
            Assert.True(position < data.Length, "Unexpected end of BIFF record.");
            byte current = data[position++];
            value |= (uint)(current & 0x7F) << shift;
            if ((current & 0x80) == 0)
                return value;
        }
        throw new InvalidDataException("Invalid BIFF variable-length integer.");
    }

    private static void DeleteIfExists(string path)
    {
        if (File.Exists(path))
            File.Delete(path);
    }

    private static IDataReader ThrowAfterReads(IDataReader inner, int readsBeforeThrow)
    {
        IDataReader proxy = DispatchProxy.Create<IDataReader, ThrowingDataReaderProxy>();
        ((ThrowingDataReaderProxy)proxy).Initialize(inner, readsBeforeThrow);
        return proxy;
    }

    public class ThrowingDataReaderProxy : DispatchProxy
    {
        private IDataReader _inner = null!;
        private int _readsBeforeThrow;
        private int _successfulReads;
        private bool _didThrow;

        internal void Initialize(IDataReader inner, int readsBeforeThrow)
        {
            _inner = inner;
            _readsBeforeThrow = readsBeforeThrow;
        }

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            if (targetMethod is null)
                throw new InvalidOperationException("The data reader proxy received no method.");
            if (targetMethod.Name == nameof(IDataReader.Read) && !_didThrow && _successfulReads >= _readsBeforeThrow)
            {
                _didThrow = true;
                throw new InvalidOperationException("Synthetic reader failure.");
            }

            object? result = targetMethod.Invoke(_inner, args);
            if (targetMethod.Name == nameof(IDataReader.Read) && result is true)
                _successfulReads++;
            return result;
        }
    }
}
