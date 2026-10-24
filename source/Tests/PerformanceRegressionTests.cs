using SpreadSheetTasks;
using System.Data;
using System.IO.Compression;
using System.Text;

namespace Tests;

/// <summary>
/// Regression tests for the performance work: inline XLSB strings, the non-boxing
/// <see cref="ExcelReaderAbstract.CurrentRow"/> API, shared/read-only file access and the
/// non-seekable-stream fallback of the custom ZIP writer.
/// </summary>
[Collection("Sequential")]
public class PerformanceRegressionTests
{
    private static DataTable CreateTypedData()
    {
        var dt = new DataTable();
        dt.Columns.Add("Id", typeof(long));
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Amount", typeof(double));
        dt.Columns.Add("When", typeof(DateTime));
        dt.Columns.Add("Active", typeof(bool));
        dt.Rows.Add(1L, "Alpha", 1.5, new DateTime(2023, 5, 6, 7, 8, 9), true);
        dt.Rows.Add(2L, "Beta", 2.25, new DateTime(2024, 1, 2, 3, 4, 5), false);
        return dt;
    }

    [Fact]
    public void XlsbWriter_InlineStrings_RoundTripsAndOmitsSharedStringsPart()
    {
        string path = $"test_inline_strings_{Guid.NewGuid():N}.xlsb";
        string[] values = ["Alpha", "Zażółć gęślą jaźń", "Alpha", "Line\nBreak"];

        try
        {
            using (var writer = new XlsbWriter(path, useSharedStrings: false))
            {
                writer.AddSheet("Sheet1");
                writer.WriteSheet(
                    ["Header"],
                    [TypeCode.String],
                    values.Select(v => new object?[] { v }).ToList());
            }

            using (var archive = ZipFile.OpenRead(path))
            {
                Assert.Null(archive.GetEntry("xl/sharedStrings.bin"));

                var relationships = archive.GetEntry("xl/_rels/workbook.bin.rels");
                Assert.NotNull(relationships);
                using (var reader = new StreamReader(relationships!.Open(), Encoding.UTF8))
                {
                    Assert.DoesNotContain("sharedStrings", reader.ReadToEnd());
                }

                var contentTypes = archive.GetEntry("[Content_Types].xml");
                Assert.NotNull(contentTypes);
                using (var reader = new StreamReader(contentTypes!.Open(), Encoding.UTF8))
                {
                    Assert.DoesNotContain("sharedStrings", reader.ReadToEnd());
                }

                // The unique text must be stored inline in the worksheet part, not in an SST.
                var sheet = archive.GetEntry("xl/worksheets/sheet1.bin");
                Assert.NotNull(sheet);
                using var sheetStream = new MemoryStream();
                using (var entry = sheet!.Open())
                {
                    entry.CopyTo(sheetStream);
                }
                byte[] sheetBytes = sheetStream.ToArray();
                Assert.True(
                    sheetBytes.AsSpan().IndexOf(Encoding.Unicode.GetBytes("Zażółć gęślą jaźń")) >= 0,
                    "inline string not found in the worksheet part");
            }

            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.Open(path);
                reader.ActualSheetName = "Sheet1";

                Assert.True(reader.Read());
                Assert.Equal("Header", reader.GetValue(0));

                foreach (string expected in values)
                {
                    Assert.True(reader.Read());
                    Assert.Equal(expected, reader.GetValue(0));
                }

                Assert.False(reader.Read());
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void CurrentRow_ExposesTypedValuesWithoutBoxing(string extension)
    {
        string path = $"test_current_row_{Guid.NewGuid():N}{extension}";
        try
        {
            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Sheet1");
                writer.WriteSheet(CreateTypedData().CreateDataReader());
            }

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(path);
            reader.ActualSheetName = "Sheet1";

            Assert.True(reader.Read()); // header
            Assert.True(reader.Read());

            ReadOnlySpan<FieldInfo> row = reader.CurrentRow;
            Assert.Equal(reader.FieldCount, row.Length);
            Assert.True(row.Length >= 5);

            Assert.True(row[0].type is ExcelDataType.Int32 or ExcelDataType.Int64 or ExcelDataType.Double);
            long id = row[0].type switch
            {
                ExcelDataType.Int32 => row[0].int32Value,
                ExcelDataType.Int64 => row[0].int64Value,
                _ => (long)row[0].doubleValue,
            };
            Assert.Equal(1L, id);

            Assert.Equal(ExcelDataType.String, row[1].type);
            Assert.Equal("Alpha", row[1].strValue);

            Assert.Equal(ExcelDataType.Double, row[2].type);
            Assert.Equal(1.5, row[2].doubleValue);

            Assert.Equal(ExcelDataType.DateTime, row[3].type);
            Assert.Equal(new DateTime(2023, 5, 6, 7, 8, 9), row[3].dtValue);

            Assert.Equal(ExcelDataType.Boolean, row[4].type);
            Assert.True(row[4].boolValue);

            // The span reflects the next row after Read(), never stale data.
            Assert.True(reader.Read());
            ReadOnlySpan<FieldInfo> next = reader.CurrentRow;
            Assert.Equal(ExcelDataType.String, next[1].type);
            Assert.Equal("Beta", next[1].strValue);
            Assert.False(next[4].boolValue);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Open_ReadOnlyFile_Succeeds(string extension)
    {
        string path = $"test_read_only_{Guid.NewGuid():N}{extension}";
        try
        {
            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Sheet1");
                writer.WriteSheet(CreateTypedData().CreateDataReader());
            }

            File.SetAttributes(path, FileAttributes.ReadOnly);

            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.Open(path);
                reader.ActualSheetName = "Sheet1";
                Assert.True(reader.Read());
                Assert.Equal("Id", reader.GetValue(0));
                Assert.True(reader.Read());
                Assert.Equal("Alpha", reader.GetValue(1));
            }
        }
        finally
        {
            if (File.Exists(path))
            {
                File.SetAttributes(path, FileAttributes.Normal);
                File.Delete(path);
            }
        }
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Open_TwoReadersOnSameFile_CanReadConcurrently(string extension)
    {
        string path = $"test_concurrent_readers_{Guid.NewGuid():N}{extension}";
        try
        {
            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Sheet1");
                writer.WriteSheet(CreateTypedData().CreateDataReader());
            }

            using var first = new XlsxOrXlsbReadOrEdit();
            using var second = new XlsxOrXlsbReadOrEdit();

            // Both archives are open at the same time; with FileShare.None the second call failed.
            first.Open(path);
            second.Open(path);
            first.ActualSheetName = "Sheet1";
            second.ActualSheetName = "Sheet1";

            Assert.True(first.Read());
            Assert.True(second.Read());

            // Interleave reads to prove the per-instance row buffer is not shared.
            while (first.Read() && second.Read())
            {
                Assert.Equal(first.GetValue(0), second.GetValue(0));
                Assert.Equal(first.GetValue(1), second.GetValue(1));
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Open_MemoryStream_ReadsWorkbook(string extension)
    {
        using var buffer = new MemoryStream();
        ExcelWriter writer = extension == ".xlsx"
            ? new XlsxWriter(buffer)
            : new XlsbWriter(buffer);
        using (writer)
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(CreateTypedData().CreateDataReader());
        }

        using var reader = new XlsxOrXlsbReadOrEdit();
        using var input = new MemoryStream(buffer.ToArray(), writable: false);
        reader.Open(input);
        reader.ActualSheetName = "Sheet1";

        Assert.True(reader.Read());
        Assert.Equal("Id", reader.GetValue(0));
        Assert.True(reader.Read());
        Assert.Equal(1L, Convert.ToInt64(reader.GetValue(0)));
        Assert.Equal("Alpha", reader.GetValue(1));
        Assert.Equal(ExcelDataType.DateTime, reader.GetExcelDataType(3));
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Writer_NonSeekableStream_FallsBackToZipArchive(string extension)
    {
        string path = $"test_non_seekable_{Guid.NewGuid():N}{extension}";
        try
        {
            using (var file = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
            using (var nonSeekable = new NonSeekableWriteStream(file))
            {
                ExcelWriter writer = extension == ".xlsx"
                    ? new XlsxWriter(nonSeekable)
                    : new XlsbWriter(nonSeekable);
                using (writer)
                {
                    writer.AddSheet("Sheet1");
                    writer.WriteSheet(CreateTypedData().CreateDataReader());
                }
            }

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(path);
            reader.ActualSheetName = "Sheet1";
            Assert.True(reader.Read());
            Assert.True(reader.Read());
            Assert.Equal("Alpha", reader.GetValue(1));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(".xlsx", CompressionLevel.Fastest)]
    [InlineData(".xlsx", CompressionLevel.Optimal)]
    [InlineData(".xlsb", CompressionLevel.Fastest)]
    [InlineData(".xlsb", CompressionLevel.Optimal)]
    public void Writer_ExplicitCompressionLevel_StillRoundTrips(string extension, CompressionLevel level)
    {
        string path = $"test_explicit_level_{Guid.NewGuid():N}{extension}";
        try
        {
            ExcelWriter writer = extension == ".xlsx"
                ? new XlsxWriter(path, _clvl: level)
                : new XlsbWriter(path, level);
            using (writer)
            {
                writer.AddSheet("Sheet1");
                writer.WriteSheet(CreateTypedData().CreateDataReader());
            }

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(path);
            reader.ActualSheetName = "Sheet1";
            Assert.True(reader.Read());
            Assert.True(reader.Read());
            Assert.Equal("Alpha", reader.GetValue(1));
        }
        finally
        {
            File.Delete(path);
        }
    }

    /// <summary>Forward-only stream: CanSeek is false, so the writer must use the BCL fallback.</summary>
    private sealed class NonSeekableWriteStream(Stream inner) : Stream
    {
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush() => inner.Flush();
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => inner.Write(buffer, offset, count);
        public override void Write(ReadOnlySpan<byte> buffer) => inner.Write(buffer);
        public override void WriteByte(byte value) => inner.WriteByte(value);

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
