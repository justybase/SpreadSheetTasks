using SpreadSheetTasks;
using System.Data;
using System.IO.Compression;
using System.Xml.Linq;

namespace Tests;

[Collection("Sequential")]
public class StyleTests
{
    private static DataTable CreateTable()
    {
        var dt = new DataTable();
        dt.Columns.Add("Product", typeof(string));
        dt.Columns.Add("Quantity", typeof(int));
        dt.Columns.Add("Note", typeof(string));
        dt.Rows.Add("Apples", 1200, DBNull.Value);
        dt.Rows.Add("Pears", 8, "fresh");
        return dt;
    }

    private static Dictionary<string, CellStyle> CreateStyles() => new()
    {
        ["B2:B3"] = new CellStyle { Fill = "FFF2CC" },
        ["C2"] = new CellStyle { Fill = "FF0000", FontColor = "FFFFFF" },
    };

    // GetRowsOfSheet yields the same array instance for every row, so clone it before storing.
    private static List<object[]> ReadRows(XlsxOrXlsbReadOrEdit reader, string sheetName)
        => reader.GetRowsOfSheet(sheetName).Select(row => (object[])row.Clone()).ToList();

    private static byte[] ReadEntry(ZipArchive archive, string name)
    {
        using var stream = archive.GetEntry(name)!.Open();
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }

    private static List<(int Id, int Offset, int Length)> ReadRecords(byte[] data)
    {
        var result = new List<(int, int, int)>();
        int i = 0;
        while (i < data.Length)
        {
            int id = ReadVarint(data, ref i);
            int length = ReadVarint(data, ref i);
            result.Add((id, i, length));
            i += length;
        }
        return result;
    }

    private static int ReadVarint(byte[] data, ref int i)
    {
        int value = 0;
        int shift = 0;
        while (true)
        {
            byte b = data[i++];
            value |= (b & 0x7F) << shift;
            if ((b & 0x80) == 0)
            {
                return value;
            }
            shift += 7;
        }
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteSheet_WithStyles_RoundTripsValues(string extension)
    {
        var path = $"test_style_roundtrip{extension}";
        try
        {
            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Data");
                writer.WriteSheet(CreateTable().CreateDataReader(),
                    headerStyle: new CellStyle { Fill = "008000", FontColor = "FFFFFF" },
                    styles: CreateStyles());
            }

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(path);
            var rows = ReadRows(reader, "Data");
            Assert.Equal(3, rows.Count);
            Assert.Equal("Product", rows[0][0]);
            Assert.Equal("Quantity", rows[0][1]);
            Assert.Equal(1200, Convert.ToInt32(rows[1][1]));
            Assert.Equal("fresh", rows[2][2]);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void WriteSheet_Xlsx_EmitsDeduplicatedFontsFillsAndAlignment()
    {
        var path = "test_style_structure.xlsx";
        try
        {
            using (var writer = new XlsxWriter(path))
            {
                writer.AddSheet("Data");
                writer.WriteSheet(CreateTable().CreateDataReader(),
                    headerStyle: new CellStyle { Fill = "008000", FontColor = "FFFFFF" },
                    styles: new Dictionary<string, CellStyle>
                    {
                        ["B2:B3"] = new CellStyle { Fill = "FFF2CC" },
                        ["C2"] = new CellStyle { Fill = "FF0000", FontColor = "FFFFFF", Horizontal = "center", Wrap = true },
                    });
            }

            using var archive = ZipFile.OpenRead(path);
            using var stream = archive.GetEntry("xl/styles.xml")!.Open();
            var doc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var root = doc.Root!;

            var fonts = root.Element(ns + "fonts")!;
            var fills = root.Element(ns + "fills")!;
            var cellXfs = root.Element(ns + "cellXfs")!;

            // default + bold + (white,bold) + white
            Assert.Equal(4, fonts.Elements(ns + "font").Count());
            // none + gray125 + green + yellow + red
            Assert.Equal(5, fills.Elements(ns + "fill").Count());
            // general + date + datetime + bold + 3 style XFs
            Assert.Equal(7, cellXfs.Elements(ns + "xf").Count());
            Assert.Equal(fonts.Elements(ns + "font").Count(), int.Parse(fonts.Attribute("count")!.Value));
            Assert.Equal(fills.Elements(ns + "fill").Count(), int.Parse(fills.Attribute("count")!.Value));
            Assert.Equal(cellXfs.Elements(ns + "xf").Count(), int.Parse(cellXfs.Attribute("count")!.Value));

            var solidColors = fills.Elements(ns + "fill")
                .Select(fill => fill.Element(ns + "patternFill")!.Element(ns + "fgColor")?.Attribute("rgb")?.Value)
                .Where(value => value is not null)
                .ToList();
            Assert.Contains("FF008000", solidColors);
            Assert.Contains("FFFFF2CC", solidColors);
            Assert.Contains("FFFF0000", solidColors);

            XElement aligned = cellXfs.Elements(ns + "xf").Single(xf => xf.Attribute("applyAlignment") is not null);
            XElement alignment = aligned.Element(ns + "alignment")!;
            Assert.Equal("center", alignment.Attribute("horizontal")!.Value);
            Assert.Equal("1", alignment.Attribute("wrapText")!.Value);

            // A blank cell inside a styled range keeps its style.
            var sheetXml = System.Text.Encoding.UTF8.GetString(ReadEntry(archive, "xl/worksheets/sheet1.xml"));
            Assert.Contains("<c r=\"C2\" s=\"", sheetXml);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void WriteSheet_Xlsb_EmitsCustomFontsFillsAndCellXfs()
    {
        var path = "test_style_structure.xlsb";
        try
        {
            using (var writer = new XlsbWriter(path))
            {
                writer.AddSheet("Data");
                writer.WriteSheet(CreateTable().CreateDataReader(),
                    headerStyle: new CellStyle { Fill = "008000", FontColor = "FFFFFF" },
                    styles: new Dictionary<string, CellStyle>
                    {
                        ["B2:B3"] = new CellStyle { Fill = "FFF2CC" },
                        ["C2"] = new CellStyle { Fill = "FF0000", FontColor = "FFFFFF" },
                    });
            }

            using var archive = ZipFile.OpenRead(path);
            byte[] styles = ReadEntry(archive, "xl/styles.bin");
            var records = ReadRecords(styles);

            var fonts = records.Where(r => r.Id == 0x2B).ToList();
            var fills = records.Where(r => r.Id == 0x2D).ToList();
            Assert.Equal(4, fonts.Count); // 2 built-in + bold white + white
            Assert.Equal(5, fills.Count); // none + gray125 + green + yellow + red

            var beginFonts = records.Single(r => r.Id == 0x263);
            Assert.Equal(4u, BitConverter.ToUInt32(styles, beginFonts.Offset));
            var beginFills = records.Single(r => r.Id == 0x25B);
            Assert.Equal(5u, BitConverter.ToUInt32(styles, beginFills.Offset));
            var beginCellXfs = records.Single(r => r.Id == 0x269);
            Assert.Equal(7u, BitConverter.ToUInt32(styles, beginCellXfs.Offset));

            // The first custom font is bold white (bold header with font color).
            var customFont = fonts.OrderBy(r => r.Offset).Skip(2).First();
            Assert.Equal(1, BitConverter.ToUInt16(styles, customFont.Offset + 2) & 1);
            Assert.Equal(
                new byte[] { 0x05, 0xFF, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF },
                styles.AsSpan(customFont.Offset + 12, 8).ToArray());

            // A cell XF must reference both a custom font and a custom fill.
            var cellXfs = new List<(int Offset, int Length)>();
            bool inCellXfs = false;
            foreach (var record in records)
            {
                if (record.Id == 0x269) { inCellXfs = true; continue; }
                if (record.Id == 0x26A) { inCellXfs = false; continue; }
                if (inCellXfs && record.Id == 0x2F) { cellXfs.Add((record.Offset, record.Length)); }
            }
            Assert.Equal(7, cellXfs.Count);
            Assert.Contains(cellXfs, xf =>
                BitConverter.ToUInt16(styles, xf.Offset + 4) >= 2
                && BitConverter.ToUInt16(styles, xf.Offset + 6) >= 2);

            // The styled blank cell is written as BrtCellBlank with a non-zero style.
            byte[] sheet = ReadEntry(archive, "xl/worksheets/sheet1.bin");
            var sheetRecords = ReadRecords(sheet);
            Assert.Contains(sheetRecords, record =>
                record.Id == 0x01 && record.Length >= 8
                && BitConverter.ToUInt32(sheet, record.Offset + 4) != 0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void WriteSheet_Xlsb_RejectsAlignmentAndWrap()
    {
        var path = "test_style_alignment.xlsb";
        try
        {
            using (var writer = new XlsbWriter(path))
            {
                Assert.Throws<NotSupportedException>(() => writer.WriteSheet(
                    CreateTable().CreateDataReader(),
                    headerStyle: new CellStyle { Horizontal = "center" }));
                Assert.Throws<NotSupportedException>(() => writer.WriteSheet(
                    CreateTable().CreateDataReader(),
                    headerStyle: null,
                    styles: new Dictionary<string, CellStyle> { ["A1"] = new CellStyle { Wrap = true } }));

                writer.AddSheet("Data");
                writer.WriteSheet(CreateTable().CreateDataReader());
            }

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(path);
            Assert.Equal(3, reader.GetRowsOfSheet("Data").Count());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void WriteSheet_RejectsInvalidStyleInput()
    {
        var path = "test_style_invalid.xlsx";
        try
        {
            using (var writer = new XlsxWriter(path))
            {
                Assert.Throws<ArgumentException>(() => writer.WriteSheet(
                    CreateTable().CreateDataReader(),
                    headerStyle: null,
                    styles: new Dictionary<string, CellStyle> { ["A0"] = new CellStyle { Fill = "FFF2CC" } }));
                Assert.Throws<ArgumentException>(() => writer.WriteSheet(
                    CreateTable().CreateDataReader(),
                    headerStyle: null,
                    styles: new Dictionary<string, CellStyle> { ["C3:A1"] = new CellStyle { Fill = "FFF2CC" } }));
                Assert.Throws<ArgumentException>(() => writer.WriteSheet(
                    CreateTable().CreateDataReader(),
                    headerStyle: null,
                    styles: new Dictionary<string, CellStyle> { ["not a cell"] = new CellStyle { Fill = "FFF2CC" } }));
                Assert.Throws<ArgumentNullException>(() => writer.WriteSheet(
                    CreateTable().CreateDataReader(),
                    headerStyle: null,
                    styles: new Dictionary<string, CellStyle> { ["A1"] = null! }));

                writer.AddSheet("Data");
                writer.WriteSheet(CreateTable().CreateDataReader());
            }

            Assert.True(File.Exists(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void CellStyle_ValidatesProperties()
    {
        Assert.Throws<ArgumentException>(() => new CellStyle { Fill = "green" });
        Assert.Throws<ArgumentException>(() => new CellStyle { FontColor = "#12" });
        Assert.Throws<ArgumentException>(() => new CellStyle { Horizontal = "middle" });
        Assert.Throws<ArgumentException>(() => new CellStyle { Vertical = "middle" });

        var style = new CellStyle { Fill = "#fff2cc", FontColor = "ffffff", Horizontal = "CENTER" };
        Assert.Equal("FFF2CC", style.Fill);
        Assert.Equal("FFFFFF", style.FontColor);
        Assert.Equal("center", style.Horizontal);
    }

    [Fact]
    public void CellStyle_Merge_OverlaysOnlySetProperties()
    {
        var baseStyle = new CellStyle { Fill = "008000", Bold = true, NumberFormat = F.THOUSANDS_SEP };
        CellStyle merged = CellStyle.Merge(baseStyle, new CellStyle { Fill = "FFF2CC", Italic = true });

        Assert.Equal("FFF2CC", merged.Fill);
        Assert.True(merged.Bold);
        Assert.True(merged.Italic);
        Assert.Equal(F.THOUSANDS_SEP, merged.NumberFormat);
        Assert.Null(merged.Horizontal);

        Assert.Equal(baseStyle, CellStyle.Merge(baseStyle, new CellStyle()));
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteSheet_FormattedCellFormatWinsOverStyleNumberFormat(string extension)
    {
        var path = $"test_style_format{extension}";
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Value", typeof(object));
            dt.Rows.Add(new FormattedCell(1234.5, F.PERCENTAGE));

            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Data");
                writer.WriteSheet(dt.CreateDataReader(), headers: false,
                    headerStyle: null,
                    styles: new Dictionary<string, CellStyle>
                    {
                        ["A1"] = new CellStyle { NumberFormat = F.THOUSANDS_SEP, Fill = "FFF2CC" },
                    });
            }

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(path);
            var rows = ReadRows(reader, "Data");
            Assert.Single(rows);
            Assert.Equal(1234.5, Convert.ToDouble(rows[0][0]));

            if (extension == ".xlsx")
            {
                using var archive = ZipFile.OpenRead(path);
                using var stream = archive.GetEntry("xl/styles.xml")!.Open();
                var doc = XDocument.Load(stream);
                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                var formatCodes = doc.Root!.Element(ns + "numFmts")!.Elements(ns + "numFmt")
                    .Select(fmt => fmt.Attribute("formatCode")!.Value)
                    .ToList();
                Assert.Contains(F.PERCENTAGE, formatCodes);
                Assert.DoesNotContain(F.THOUSANDS_SEP, formatCodes);
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
    public void WriteSheet_StyleRangesUseAbsoluteCoordinates(string extension)
    {
        var path = $"test_style_offsets{extension}";
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Value", typeof(string));
            dt.Rows.Add("offset");

            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Data");
                writer.WriteSheet(dt.CreateDataReader(), headers: false, startingRow: 1, startingColumn: 1,
                    headerStyle: null,
                    styles: new Dictionary<string, CellStyle> { ["B2"] = new CellStyle { Fill = "FFF2CC" } });
            }

            if (extension == ".xlsx")
            {
                using var archive = ZipFile.OpenRead(path);
                var sheetXml = System.Text.Encoding.UTF8.GetString(ReadEntry(archive, "xl/worksheets/sheet1.xml"));
                // B2 is the only written cell and must carry the style.
                Assert.Contains("<row r=\"2\"><c r=\"B2\" s=\"4\" t=\"s\">", sheetXml);

                // A range keyed relative to the write start must not style B2.
                using (var writer = new XlsxWriter(path))
                {
                    writer.AddSheet("Data");
                    writer.WriteSheet(dt.CreateDataReader(), headers: false, startingRow: 1, startingColumn: 1,
                        headerStyle: null,
                        styles: new Dictionary<string, CellStyle> { ["A1"] = new CellStyle { Fill = "FFF2CC" } });
                }
                using var archive2 = ZipFile.OpenRead(path);
                var sheetXml2 = System.Text.Encoding.UTF8.GetString(ReadEntry(archive2, "xl/worksheets/sheet1.xml"));
                Assert.Contains("<row r=\"2\"><c r=\"B2\" t=\"s\">", sheetXml2);
            }
            else
            {
                using var archive = ZipFile.OpenRead(path);
                byte[] sheet = ReadEntry(archive, "xl/worksheets/sheet1.bin");
                var records = ReadRecords(sheet);
                var row = Assert.Single(records, record => record.Id == 0x00);
                Assert.Equal(1, BitConverter.ToInt32(sheet, row.Offset));
                var cell = Assert.Single(records, record => record.Id == 0x07);
                Assert.Equal(1, BitConverter.ToInt32(sheet, cell.Offset));
                Assert.NotEqual(0, BitConverter.ToInt32(sheet, cell.Offset + 4));
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
    public void WriteSheet_NumberFormatSharedByStyleAndFormattedCell(string extension)
    {
        var path = $"test_style_shared_format{extension}";
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Style", typeof(double));
            dt.Columns.Add("Tuple", typeof(object));
            dt.Rows.Add(1234.5, new FormattedCell(5678.9, F.THOUSANDS_SEP));

            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Data");
                writer.WriteSheet(dt.CreateDataReader(), headers: false,
                    headerStyle: null,
                    styles: new Dictionary<string, CellStyle> { ["A1"] = new CellStyle { NumberFormat = F.THOUSANDS_SEP } });
            }

            using (var reader = new XlsxOrXlsbReadOrEdit())
            {
                reader.Open(path);
                var rows = ReadRows(reader, "Data");
                Assert.Single(rows);
                Assert.Equal(1234.5, Convert.ToDouble(rows[0][0]));
                Assert.Equal(5678.9, Convert.ToDouble(rows[0][1]));
            }

            if (extension == ".xlsx")
            {
                using var archive = ZipFile.OpenRead(path);
                using var stream = archive.GetEntry("xl/styles.xml")!.Open();
                var doc = XDocument.Load(stream);
                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                var numFmts = doc.Root!.Element(ns + "numFmts")!;
                Assert.Single(numFmts.Elements(ns + "numFmt"));
                int numFmtId = int.Parse(numFmts.Elements(ns + "numFmt").Single().Attribute("numFmtId")!.Value);
                var cellXfs = doc.Root!.Element(ns + "cellXfs")!.Elements(ns + "xf").ToList();
                // Both cells use one shared format id on two different XFs.
                Assert.Equal(numFmtId, int.Parse(cellXfs[4].Attribute("numFmtId")!.Value));
                Assert.Equal(numFmtId, int.Parse(cellXfs[5].Attribute("numFmtId")!.Value));
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
    public void WriteSheet_StyledDateKeepsDateFormat(string extension)
    {
        var path = $"test_style_date{extension}";
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Date", typeof(DateTime));
            dt.Rows.Add(new DateTime(2026, 6, 1));

            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Data");
                writer.WriteSheet(dt.CreateDataReader(), headers: false,
                    headerStyle: null,
                    styles: new Dictionary<string, CellStyle> { ["A1"] = new CellStyle { Fill = "FFF2CC" } });
            }

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(path);
            var rows = ReadRows(reader, "Data");
            Assert.Single(rows);
            Assert.Equal(new DateTime(2026, 6, 1), Convert.ToDateTime(rows[0][0]));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteSheet_HeaderStyleAppliesToNumericHeaderCells(string extension)
    {
        var path = $"test_style_header{extension}";
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Amount", typeof(int));
            dt.Columns.Add("Price", typeof(double));
            dt.Rows.Add(10, 1.5);

            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Data");
                writer.WriteSheet(dt.CreateDataReader(), headers: false,
                    headerStyle: new CellStyle { Fill = "008000", FontColor = "FFFFFF" });
            }

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(path);
            var readRows = ReadRows(reader, "Data");
            Assert.Single(readRows);
            Assert.Equal(10, Convert.ToInt32(readRows[0][0]));
            Assert.Equal(1.5, Convert.ToDouble(readRows[0][1]));

            if (extension == ".xlsx")
            {
                using var archive = ZipFile.OpenRead(path);
                var sheetXml = System.Text.Encoding.UTF8.GetString(ReadEntry(archive, "xl/worksheets/sheet1.xml"));
                // Both numeric first-row cells carry the header style (xf 4: bold font + green fill).
                Assert.Contains("<c s=\"4\"><v>10</v></c>", sheetXml);
                Assert.Contains("<c s=\"4\"><v>1.5</v></c>", sheetXml);
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
    public void WriteSheet_StyleRangeOnHeaderRowKeepsDefaultBold(string extension)
    {
        var path = $"test_style_header_range{extension}";
        try
        {
            var dt = new DataTable();
            dt.Columns.Add("Value", typeof(string));
            dt.Rows.Add("row");

            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Data");
                writer.WriteSheet(dt.CreateDataReader(),
                    headerStyle: null,
                    styles: new Dictionary<string, CellStyle> { ["A1"] = new CellStyle { Fill = "FFF2CC" } });
            }

            if (extension == ".xlsx")
            {
                using var archive = ZipFile.OpenRead(path);
                using var stream = archive.GetEntry("xl/styles.xml")!.Open();
                var doc = XDocument.Load(stream);
                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                var xf = doc.Root!.Element(ns + "cellXfs")!.Elements(ns + "xf")
                    .Single(item => item.Attribute("fillId")?.Value == "2");
                // Font 1 is the built-in bold font; the header range must not drop the bold.
                Assert.Equal("1", xf.Attribute("fontId")!.Value);
                Assert.Equal("1", xf.Attribute("applyFont")!.Value);
                Assert.Equal("1", xf.Attribute("applyFill")!.Value);
            }
            else
            {
                using var archive = ZipFile.OpenRead(path);
                byte[] styles = ReadEntry(archive, "xl/styles.bin");
                var records = ReadRecords(styles);
                var cellXfs = new List<(int Offset, int Length)>();
                bool inCellXfs = false;
                foreach (var record in records)
                {
                    if (record.Id == 0x269) { inCellXfs = true; continue; }
                    if (record.Id == 0x26A) { inCellXfs = false; continue; }
                    if (inCellXfs && record.Id == 0x2F) { cellXfs.Add((record.Offset, record.Length)); }
                }
                // Find the XF with the custom fill and check it still uses the built-in bold font.
                var xf = Assert.Single(cellXfs, item => BitConverter.ToUInt16(styles, item.Offset + 6) == 2);
                Assert.Equal(1, BitConverter.ToUInt16(styles, xf.Offset + 4));
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
    public void WriteSheet_ConvenienceOverloadsPassStyles(string extension)
    {
        var path = $"test_style_overloads{extension}";
        try
        {
            var table = new DataTable();
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Count", typeof(int));
            table.Rows.Add("Table row", 3);

            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Table");
                writer.WriteSheet(table,
                    headerStyle: new CellStyle { Fill = "008000", FontColor = "FFFFFF" },
                    styles: new Dictionary<string, CellStyle> { ["B2"] = new CellStyle { Fill = "FFF2CC" } });

                writer.AddSheet("Objects");
                writer.WriteSheet(new object?[][] { new object?[] { "Object row", 7 } },
                    headerStyle: null,
                    styles: new Dictionary<string, CellStyle> { ["B2"] = new CellStyle { Fill = "FF0000" } });

                writer.AddSheet("List");
                writer.WriteSheet(
                    new List<string> { "Name", "Count" },
                    new List<TypeCode> { TypeCode.String, TypeCode.Int32 },
                    new List<object?[]> { new object?[] { "List row", 9 } },
                    headerStyle: new CellStyle { Fill = "008000" },
                    styles: new Dictionary<string, CellStyle> { ["B2"] = new CellStyle { Fill = "0000FF" } });

                writer.AddSheet("Formatted");
                writer.WriteSheet(
                    new FormattedCell?[][] { new FormattedCell?[] { new FormattedCell(1234, F.THOUSANDS_SEP) } },
                    new[] { "Value" },
                    headerStyle: new CellStyle { Fill = "0000FF" });
            }

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(path);
            Assert.Equal("Table row", ReadRows(reader, "Table")[1][0]);
            Assert.Equal(7, Convert.ToInt32(ReadRows(reader, "Objects")[1][1]));
            Assert.Equal("List row", ReadRows(reader, "List")[1][0]);
            Assert.Equal(9, Convert.ToInt32(ReadRows(reader, "List")[1][1]));
            Assert.Equal(1234, Convert.ToInt32(ReadRows(reader, "Formatted")[1][0]));
        }
        finally
        {
            File.Delete(path);
        }
    }
}
