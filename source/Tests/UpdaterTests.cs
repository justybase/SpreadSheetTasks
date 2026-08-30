using SpreadSheetTasks;
using System.Buffers.Binary;
using System.Data;
using System.IO.Compression;

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
}
