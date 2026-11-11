using SpreadSheetTasks;
using System.Data;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using Sylvan.Data.Excel;

namespace Tests;

/// <summary>
/// Differential tests for the fast XLSX worksheet scanner: every fixture is read twice, once with
/// the scanner and once with the XmlReader path, and both must return exactly the same rows.
/// The corpus covers workbooks written by this library (shared and inline strings), Sylvan.Data.Excel,
/// Excel COM (when available) and handcrafted SpreadsheetML edge cases.
/// </summary>
[Collection("Sequential")]
public class XlsxFastReaderParityTests
{
    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private static string TempFile() => Path.Combine(Path.GetTempPath(), $"sst_fast_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void OurWriter_SharedStrings_MatchesXmlReader()
    {
        string path = TempFile();
        try
        {
            DataTable table = BuildTypedTable(400);
            using (var writer = new XlsxWriter(path, useScharedStrings: true))
            {
                writer.AddSheet("Data");
                writer.WriteSheet(table.CreateDataReader());
            }

            AssertParity(path);
            AssertTableValues(table, path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void OurWriter_InlineStrings_MatchesXmlReader()
    {
        string path = TempFile();
        try
        {
            DataTable table = BuildTypedTable(250);
            using (var writer = new XlsxWriter(path, useScharedStrings: false))
            {
                writer.AddSheet("Data");
                writer.WriteSheet(table.CreateDataReader());
            }

            AssertParity(path);
            AssertTableValues(table, path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void OurWriter_ManyColumnsAndSheets_MatchesXmlReader()
    {
        string path = TempFile();
        try
        {
            var table = new DataTable();
            for (int c = 0; c < 80; c++)
            {
                table.Columns.Add($"C{c}", c % 3 == 0 ? typeof(string) : typeof(double));
            }
            for (int r = 0; r < 20; r++)
            {
                var values = new object[80];
                for (int c = 0; c < 80; c++)
                {
                    values[c] = c % 3 == 0 ? $"text {r}/{c}" : r * 1000.5 + c;
                }
                table.Rows.Add(values);
            }

            using (var writer = new XlsxWriter(path))
            {
                writer.AddSheet("First");
                writer.WriteSheet(table.CreateDataReader());
                writer.AddSheet("Second", hidden: true);
                writer.WriteSheet(table.CreateDataReader());
            }

            AssertParity(path);
            AssertTableValues(table, path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SylvanWriter_MatchesXmlReader()
    {
        string path = TempFile();
        try
        {
            DataTable table = BuildTypedTable(150);
            using (var writer = ExcelDataWriter.Create(path))
            {
                writer.Write(table.CreateDataReader(), "Data");
            }

            AssertParity(path);

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(path);
            reader.ActualSheetName = reader.GetSheetNames()[0];
            Assert.True(reader.Read());
            Assert.Equal("Text", reader.GetString(0));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void TreatAllColumnsAsText_MatchesXmlReader()
    {
        string path = TempFile();
        try
        {
            DataTable table = BuildTypedTable(50);
            using (var writer = new XlsxWriter(path))
            {
                writer.AddSheet("Data");
                writer.WriteSheet(table.CreateDataReader());
            }

            AssertParity(path, configure: r => r.TreatAllColumnsAsText = true);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void HandcraftedWorksheets_MatchXmlReader()
    {
        foreach ((string name, byte[] sheetXml, string? styles, string? strings, int expectedRows) in HandcraftedCases())
        {
            string path = TempFile();
            try
            {
                CreatePackage(path, [(name, sheetXml)], styles, strings);

                using var reader = new XlsxOrXlsbReadOrEdit();
                reader.Open(path);
                reader.ActualSheetName = name;

                // Rows are emitted lazily; switch away and back mid-stream to prove the sheet
                // restarts cleanly, then run the fast-vs-XmlReader parity check.
                // (The workbook has a single sheet, so this also exercises the same-name path.)
                if (expectedRows > 1)
                {
                    Assert.True(reader.Read());
                    reader.ActualSheetName = "Missing";
                    reader.ActualSheetName = name;
                }

                AssertParity(path);

                reader.ActualSheetName = name;
                int rows = 0;
                while (reader.Read())
                {
                    rows++;
                }
                Assert.True(rows == expectedRows, $"{name}: read {rows} rows, expected {expectedRows}");
            }
            finally
            {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void HandcraftedValues_AreReadExactly()
    {
        string styles = StylesXml();
        const string strings = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
              <si><t>Alpha</t></si>
              <si><t>A&amp;B &lt;tag&gt; &quot;q&quot; &apos;a&apos;</t></si>
              <si><t xml:space="preserve">  spaced  </t></si>
              <si><t>Zaz&#243;&#322;&#263; &#128512;</t></si>
            </sst>
            """;

        const string sheet = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <dimension ref="A1:I6"/>
              <sheetData>
                <row r="1">
                  <c r="A1" t="s"><v>0</v></c>
                  <c r="B1" t="s"><v>1</v></c>
                  <c r="C1" t="s"><v>2</v></c>
                  <c r="D1" t="s"><v>3</v></c>
                  <c r="E1"><v>42</v></c>
                  <c r="F1"><v>-7</v></c>
                  <c r="G1"><v>1.5</v></c>
                  <c r="H1"><v>1.25E+3</v></c>
                  <c r="I1" s="1"><v>45000</v></c>
                </row>
                <row r="2">
                  <c r="A2" t="b"><v>1</v></c>
                  <c r="B2" t="b"><v>0</v></c>
                  <c r="C2" t="e"><v>#DIV/0!</v></c>
                  <c r="D2" t="str"><v>formula text</v></c>
                  <c r="E2" t="inlineStr"><is><t>inline</t></is></c>
                  <c r="F2"><f>A1+1</f><v>43</v></c>
                  <c r="G2" t="inlineStr"><is><t xml:space="preserve"> padded </t></is></c>
                  <c r="H2"/>
                  <c r="I2" s="1"/>
                </row>
                <row r="4">
                  <c r="C4"><v>9</v></c>
                </row>
              </sheetData>
            </worksheet>
            """;

        string path = TempFile();
        try
        {
            CreatePackage(path, [("Values", Encoding.UTF8.GetBytes(sheet))], styles, strings);
            AssertParity(path);

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(path);
            reader.ActualSheetName = "Values";

            Assert.True(reader.Read());
            Assert.Equal("Alpha", reader.GetValue(0));
            Assert.Equal("A&B <tag> \"q\" 'a'", reader.GetValue(1));
            Assert.Equal("  spaced  ", reader.GetValue(2));
            Assert.Equal("Zazółć 😀", reader.GetValue(3));
            Assert.Equal(42L, reader.GetValue(4));
            Assert.Equal(-7L, reader.GetValue(5));
            Assert.Equal(1.5, reader.GetValue(6));
            Assert.Equal(1250.0, reader.GetValue(7));
            Assert.Equal(new DateTime(2023, 3, 15), reader.GetValue(8));

            Assert.True(reader.Read());
            Assert.Equal(true, reader.GetValue(0));
            Assert.Equal(false, reader.GetValue(1));
            Assert.Equal("error in cell", reader.GetValue(2));
            Assert.Equal("formula text", reader.GetValue(3));
            Assert.Equal("inline", reader.GetValue(4));
            Assert.Equal(43L, reader.GetValue(5));
            Assert.Equal(" padded ", reader.GetValue(6));
            Assert.Equal(DBNull.Value, reader.GetValue(7));
            Assert.Equal(DBNull.Value, reader.GetValue(8));

            // Row 3 is missing in the file; both readers must surface it as an empty row.
            Assert.True(reader.Read());
            Assert.Equal(DBNull.Value, reader.GetValue(0));

            Assert.True(reader.Read());
            Assert.Equal(9L, reader.GetValue(2));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void AttributeOrderAndWhitespace_AreReadExactly()
    {
        string sheet = $$"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="{{MainNs}}">
              <sheetData>
                <row r="1">
                  <c t="s" r="A1"><v>0</v></c>
                  <c s="1" r="B1"><v>45000</v></c>
                  <c r="C1"><v>2</v></c>
                </row>
                <row r="2">
                  <c r="A2" t="inlineStr"><is><t>   </t></is></c>
                  <c r="B2" t="inlineStr"><is><t xml:space="preserve">   </t></is></c>
                  <c r="C2" t="inlineStr"><is><t>kept</t></is></c>
                </row>
              </sheetData>
            </worksheet>
            """;

        string path = TempFile();
        try
        {
            CreatePackage(path, [("Values", Encoding.UTF8.GetBytes(sheet))], StylesXml(), StringsXml("shared"));

            AssertParity(path);

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(path);
            reader.ActualSheetName = "Values";

            Assert.True(reader.Read());
            Assert.Equal("shared", reader.GetValue(0));
            Assert.Equal(new DateTime(2023, 3, 15), reader.GetValue(1));
            Assert.Equal(2L, reader.GetValue(2));

            Assert.True(reader.Read());
            Assert.Equal(string.Empty, (string)reader.GetValue(0));
            Assert.Equal("   ", (string)reader.GetValue(1));
            Assert.Equal("kept", (string)reader.GetValue(2));

            Assert.False(reader.Read());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void NumericFormats_UseFastParsersAndMatchXmlReader()
    {
        string sheet = $$"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="{{MainNs}}">
              <sheetData>
                <row r="1">
                  <c r="A1"><v>0</v></c>
                  <c r="B1"><v>-0</v></c>
                  <c r="C1"><v>1.5</v></c>
                  <c r="D1"><v>.5</v></c>
                  <c r="E1"><v>5.</v></c>
                  <c r="F1"><v>1e5</v></c>
                  <c r="G1"><v>1E-5</v></c>
                  <c r="H1"><v>1.5e3</v></c>
                  <c r="I1"><v>01.50</v></c>
                  <c r="J1"><v>0.1</v></c>
                  <c r="K1"><v>35.840000000000003</v></c>
                  <c r="L1"><v>165258.23999999999</v></c>
                  <c r="M1"><v>1.2345678901234567</v></c>
                  <c r="N1"><v>1e300</v></c>
                </row>
                <row r="2">
                  <c r="A2"><v>123456789012345678</v></c>
                  <c r="B2"><v>-9223372036854775808</v></c>
                  <c r="C2"><v>9223372036854775807</v></c>
                  <c r="D2"><v>00042</v></c>
                  <c r="E2"><v>-00042</v></c>
                  <c r="F2"><v>1e22</v></c>
                  <c r="G2"><v>1e23</v></c>
                  <c r="H2"><v>1e-22</v></c>
                  <c r="I2"><v>0.000001</v></c>
                  <c r="J2" s="1"><v>45000.25</v></c>
                  <c r="K2" s="1"><v>2.5</v></c>
                </row>
              </sheetData>
            </worksheet>
            """;

        string path = TempFile();
        try
        {
            CreatePackage(path, [("Values", Encoding.UTF8.GetBytes(sheet))], StylesXml(), sharedStringsXml: null);
            AssertParity(path);

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(path);
            reader.ActualSheetName = "Values";

            Assert.True(reader.Read());
            Assert.Equal(0L, reader.GetValue(0));
            Assert.Equal(0L, reader.GetValue(1));
            Assert.Equal(1.5, reader.GetValue(2));
            Assert.Equal(0.5, reader.GetValue(3));
            Assert.Equal(5.0, reader.GetValue(4));
            Assert.Equal(100000.0, reader.GetValue(5));
            Assert.Equal(1e-05, reader.GetValue(6));
            Assert.Equal(1500.0, reader.GetValue(7));
            Assert.Equal(1.5, reader.GetValue(8));
            Assert.Equal(0.1, reader.GetValue(9));
            Assert.Equal(35.840000000000003, reader.GetValue(10));
            Assert.Equal(165258.23999999999, reader.GetValue(11));
            Assert.Equal(1.2345678901234567, reader.GetValue(12));
            Assert.Equal(1e300, reader.GetValue(13));

            Assert.True(reader.Read());
            Assert.Equal(123456789012345678L, reader.GetValue(0));
            Assert.Equal(long.MinValue, reader.GetValue(1));
            Assert.Equal(long.MaxValue, reader.GetValue(2));
            Assert.Equal(42L, reader.GetValue(3));
            Assert.Equal(-42L, reader.GetValue(4));
            Assert.Equal(1e22, reader.GetValue(5));
            Assert.Equal(1e23, reader.GetValue(6));
            Assert.Equal(1e-22, reader.GetValue(7));
            Assert.Equal(0.000001, reader.GetValue(8));
            Assert.Equal(new DateTime(2023, 3, 15, 6, 0, 0), reader.GetValue(9));
            Assert.Equal(new DateTime(1900, 1, 1, 12, 0, 0), reader.GetValue(10));

            Assert.False(reader.Read());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void FormulaStringXstring_IsDecodedConsistently()
    {
        string sheet = $$"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="{{MainNs}}">
              <sheetData>
                <row r="1">
                  <c r="A1" t="str"><f>"x"</f><v>_x0041_</v></c>
                  <c r="B1" t="str"><v>plain</v></c>
                </row>
              </sheetData>
            </worksheet>
            """;

        string path = TempFile();
        try
        {
            CreatePackage(path, [("Formula", Encoding.UTF8.GetBytes(sheet))], stylesXml: null, sharedStringsXml: null);
            AssertParity(path);

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(path);
            reader.ActualSheetName = "Formula";

            Assert.True(reader.Read());
            Assert.Equal("A", reader.GetValue(0));
            Assert.Equal("plain", reader.GetValue(1));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void XmlSpaceInheritance_MatchesXmlReader()
    {
        string sheet = $$"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="{{MainNs}}" xml:space="preserve">
              <sheetData>
                <row r="1"><c r="A1" t="inlineStr"><is><t>   </t></is></c></row>
                <row r="2" xml:space="default"><c r="A2" t="inlineStr"><is><t>   </t></is></c></row>
                <row r="3" xml:space="default"><c r="A3" xml:space="preserve" t="inlineStr"><is><t>   </t></is></c></row>
                <row r="4" xml:space="default"><c r="A4" t="inlineStr"><is xml:space="preserve"><t>   </t></is></c></row>
                <row r="5"><c r="A5" t="inlineStr"><is><t xml:space="default">   </t></is></c></row>
                <row r="6"><c r="A6" t="inlineStr"><is><t>   </t></is></c></row>
              </sheetData>
            </worksheet>
            """;

        string path = TempFile();
        try
        {
            CreatePackage(path, [("Space", Encoding.UTF8.GetBytes(sheet))], stylesXml: null, sharedStringsXml: null);
            AssertParity(path);

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(path);
            reader.ActualSheetName = "Space";

            string[] expected = ["   ", "", "   ", "   ", "", "   "];
            foreach (string expectedValue in expected)
            {
                Assert.True(reader.Read());
                Assert.Equal(expectedValue, (string)reader.GetValue(0));
            }
            Assert.False(reader.Read());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SharedStrings_RichRunsEmptyItemsAndFallback_AreReadExactly()
    {
        const string strings = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="8" uniqueCount="8" xml:space="preserve">
              <si><t>plain</t></si>
              <si><r><t>ri</t></r><r><rPr><b/></rPr><t>ch</t></r><rPh sb="0" eb="1"><t>PHONETIC</t></rPh></si>
              <si><r><t xml:space="preserve"> a </t></r><r><t xml:space="preserve"> b </t></r></si>
              <si/>
              <si><t/></si>
              <si><t>A&amp;B&#32;&#x43;</t></si>
              <si><t>line&#13;break</t></si>
              <si><t>   </t></si>
            </sst>
            """;
        const string sheet = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <sheetData>
                <row r="1">
                  <c r="A1" t="s"><v>0</v></c>
                  <c r="B1" t="s"><v>1</v></c>
                  <c r="C1" t="s"><v>2</v></c>
                  <c r="D1" t="s"><v>3</v></c>
                  <c r="E1" t="s"><v>4</v></c>
                  <c r="F1" t="s"><v>5</v></c>
                  <c r="G1" t="s"><v>6</v></c>
                  <c r="H1" t="s"><v>7</v></c>
                </row>
              </sheetData>
            </worksheet>
            """;

        string path = TempFile();
        try
        {
            CreatePackage(path, [("Values", Encoding.UTF8.GetBytes(sheet))], stylesXml: null, sharedStringsXml: strings);

            using var fast = new XlsxOrXlsbReadOrEdit();
            fast.Open(path);
            fast.ActualSheetName = "Values";
            Assert.True(fast._fastSharedStringsUsed);
            List<object?[]> fastRows = ReadAll(fast);

            using var xml = new XlsxOrXlsbReadOrEdit { _disableFastSharedStringsReader = true };
            xml.Open(path);
            xml.ActualSheetName = "Values";
            Assert.False(xml._fastSharedStringsUsed);
            List<object?[]> xmlRows = ReadAll(xml);

            AssertRowsEqual(xmlRows, fastRows);
            AssertRow(
                ["plain", "rich", " a  b ", "", "", "A&B C", "line\rbreak", "   "],
                fastRows[0]);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SharedStrings_UniqueCountZeroAndWhitespace_AreHandled()
    {
        const string strings = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="5" uniqueCount="0">
              <si><t>only</t></si>
              <si><t>   </t></si>
              <si><t xml:space="default">   </t></si>
              <si><t xml:space="preserve">   </t></si>
              <si><t>&#32;</t></si>
            </sst>
            """;
        const string sheet = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <sheetData>
                <row r="1">
                  <c r="A1" t="s"><v>0</v></c>
                  <c r="B1" t="s"><v>1</v></c>
                  <c r="C1" t="s"><v>2</v></c>
                  <c r="D1" t="s"><v>3</v></c>
                  <c r="E1" t="s"><v>4</v></c>
                </row>
              </sheetData>
            </worksheet>
            """;

        string path = TempFile();
        try
        {
            CreatePackage(path, [("Values", Encoding.UTF8.GetBytes(sheet))], stylesXml: null, sharedStringsXml: strings);

            using var fast = new XlsxOrXlsbReadOrEdit();
            fast.Open(path);
            fast.ActualSheetName = "Values";
            Assert.True(fast._fastSharedStringsUsed);
            Assert.True(fast.Read());
            Assert.Equal("only", fast.GetValue(0));
            Assert.Equal(string.Empty, fast.GetValue(1));
            Assert.Equal(string.Empty, fast.GetValue(2));
            Assert.Equal("   ", fast.GetValue(3));
            Assert.Equal(string.Empty, fast.GetValue(4));

            using var xml = new XlsxOrXlsbReadOrEdit { _disableFastSharedStringsReader = true };
            xml.Open(path);
            xml.ActualSheetName = "Values";
            Assert.True(xml.Read());
            for (int column = 0; column < 5; column++)
            {
                Assert.True(Equals(fast.GetValue(column), xml.GetValue(column)),
                    $"column {column}: '{Format(fast.GetValue(column))}' vs '{Format(xml.GetValue(column))}'");
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void SharedStrings_UnsupportedItemMarkup_FallsBackToXmlReader()
    {
        const string strings = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
              <si><t><![CDATA[cd&ata]]></t></si>
            </sst>
            """;
        const string sheet = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>
            </worksheet>
            """;

        string path = TempFile();
        try
        {
            CreatePackage(path, [("Values", Encoding.UTF8.GetBytes(sheet))], stylesXml: null, sharedStringsXml: strings);

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(path);
            reader.ActualSheetName = "Values";
            Assert.False(reader._fastSharedStringsUsed);
            Assert.True(reader.Read());
            Assert.Equal("cd&ata", reader.GetValue(0));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InlineStringEntities_AreDecodedWithoutRowFallbacks()
    {
        var sheet = new StringBuilder();
        sheet.Append($$"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{{MainNs}}"><sheetData>""");
        const int rowCount = 20;
        for (int row = 1; row <= rowCount; row++)
        {
            sheet.Append($"""<row r="{row}"><c r="A{row}" t="inlineStr"><is><t>row {row} &amp; &#65; &lt;x&gt; &#x1F600;</t></is></c></row>""");
        }
        sheet.Append("</sheetData></worksheet>");

        string path = TempFile();
        try
        {
            CreatePackage(path, [("Values", Encoding.UTF8.GetBytes(sheet.ToString()))], stylesXml: null, sharedStringsXml: null);

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(path);
            reader.ActualSheetName = "Values";

            int rows = 0;
            while (reader.Read())
            {
                rows++;
                Assert.Equal($"row {rows} & A <x> 😀", reader.GetValue(0));
            }

            Assert.Equal(rowCount, rows);
            Assert.True(reader._fastXlsxReaderUsed);
            Assert.Equal(rowCount, reader._fastRowsParsed);
            Assert.Equal(0, reader._fastRowsFallback);

            AssertParity(path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void RichInlineStrings_SwitchToXmlReaderMidSheet_AndPreserveValues()
    {
        var sheet = new StringBuilder();
        sheet.Append($$"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{{MainNs}}"><sheetData>""");
        for (int row = 1; row <= 5; row++)
        {
            sheet.Append($"""<row r="{row}"><c r="A{row}" t="inlineStr"><is><t>simple {row}</t></is></c></row>""");
        }
        for (int row = 6; row <= 45; row++)
        {
            sheet.Append($"""<row r="{row}"><c r="A{row}" t="inlineStr"><is><r><t>x</t></r><r><t>{row}</t></r></is></c></row>""");
        }
        for (int row = 46; row <= 50; row++)
        {
            sheet.Append($"""<row r="{row}"><c r="A{row}" t="inlineStr"><is><t>simple {row}</t></is></c></row>""");
        }
        sheet.Append("</sheetData></worksheet>");

        string path = TempFile();
        try
        {
            CreatePackage(path, [("Values", Encoding.UTF8.GetBytes(sheet.ToString()))], stylesXml: null, sharedStringsXml: null);

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(path);
            reader.ActualSheetName = "Values";

            for (int row = 1; row <= 50; row++)
            {
                Assert.True(reader.Read(), $"row {row} missing");
                string expected = row <= 5 || row >= 46 ? $"simple {row}" : $"x{row}";
                Assert.Equal(expected, reader.GetValue(0));
            }
            Assert.False(reader.Read());

            Assert.True(reader._fastXlsxReaderUsed);
            Assert.False(reader._fastActive);
            Assert.Equal(32, reader._fastRowsFallback);
            Assert.Equal(37, reader._fastRowsParsed);

            AssertParity(path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void EscapedXmlSpaceValues_MatchXmlReader()
    {
        const string strings = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1" xml:space="pre&#115;erve">
              <si><t>  </t></si>
            </sst>
            """;
        const string sheet = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <sheetData>
                <row r="1">
                  <c r="A1" t="inlineStr"><is><t xml:space="pre&#115;erve">  </t></is></c>
                  <c r="B1" t="inlineStr"><is><t xml:space="def&#97;ult">  </t></is></c>
                  <c r="C1" t="s"><v>0</v></c>
                </row>
              </sheetData>
            </worksheet>
            """;

        string path = TempFile();
        try
        {
            CreatePackage(path, [("Values", Encoding.UTF8.GetBytes(sheet))], stylesXml: null, sharedStringsXml: strings);
            AssertParity(path);

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(path);
            reader.ActualSheetName = "Values";
            Assert.True(reader.Read());
            Assert.Equal("  ", reader.GetValue(0));
            Assert.Equal(string.Empty, reader.GetValue(1));
            Assert.Equal("  ", reader.GetValue(2));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void EmptyFirstRowAndLeadingColumnChanges_AlignColumns()
    {
        const string sheet = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <sheetData>
                <row r="1"><c r="B1" t="inlineStr"><is><t>b1</t></is></c></row>
                <row r="2">
                  <c r="A2" t="inlineStr"><is><t>a2</t></is></c>
                  <c r="B2" t="inlineStr"><is><t>b2</t></is></c>
                </row>
                <row r="3"><c r="A3" t="inlineStr"><is><t>a3</t></is></c></row>
              </sheetData>
            </worksheet>
            """;

        string path = TempFile();
        try
        {
            CreatePackage(path, [("Values", Encoding.UTF8.GetBytes(sheet))], stylesXml: null, sharedStringsXml: null);
            AssertParity(path);

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(path);
            reader.ActualSheetName = "Values";

            Assert.True(reader.Read());
            Assert.Equal(1, reader.FieldCount);
            Assert.Equal("b1", reader.GetValue(0));

            Assert.True(reader.Read());
            Assert.Equal(2, reader.FieldCount);
            Assert.Equal("a2", reader.GetValue(0));
            Assert.Equal("b2", reader.GetValue(1));

            Assert.True(reader.Read());
            Assert.Equal("a3", reader.GetValue(0));
            Assert.Equal(DBNull.Value, reader.GetValue(1));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Date1904Workbook_ConvertsDates()
    {
        const string sheet = """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <sheetData><row r="1"><c r="A1" s="1"><v>100</v></c></row></sheetData>
            </worksheet>
            """;

        string path = TempFile();
        try
        {
            CreatePackage(
                path,
                [("Values", Encoding.UTF8.GetBytes(sheet))],
                StylesXml(),
                sharedStringsXml: null,
                workbookPrXml: """<workbookPr date1904="1"/>""");
            AssertParity(path);

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(path);
            reader.ActualSheetName = "Values";

            Assert.True(reader.Read());
            Assert.Equal(SpreadSheetTasks.ExcelDataType.DateTime, reader.GetExcelDataType(0));
            Assert.Equal(new DateTime(1904, 4, 10), reader.GetValue(0));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(0u, false)]
    [InlineData(1u, true)]
    public void XlsbWorkbookProperties_Date1904Flag_IsParsed(uint flags, bool expected)
    {
        byte[] record =
        [
            0x99, 0x01, // BrtWbProp (0x0099) as a BIFF12 varint
            0x04,       // payload length
            (byte)(flags & 0xFF),
            (byte)((flags >> 8) & 0xFF),
            (byte)((flags >> 16) & 0xFF),
            (byte)((flags >> 24) & 0xFF),
        ];

        using var reader = new BiffReaderWriter(record);
        Assert.True(reader.ReadWorkbook());
        Assert.Equal(expected, reader._uses1904DateSystem);
    }

    [Fact]
    public void InvalidUtf8Worksheet_ThrowsLikeXmlReader()
    {
        var sheet = new StringBuilder();
        sheet.Append($$"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{{MainNs}}"><sheetData>""");
        const int rowCount = 20_000;
        for (int row = 1; row <= rowCount; row++)
        {
            sheet.Append($"""<row r="{row}"><c r="A{row}"><v>{row}</v></c></row>""");
        }
        sheet.Append($"""<row r="{rowCount + 1}"><c r="A{rowCount + 1}" t="inlineStr"><is><t>bad""");

        byte[] prefix = Encoding.UTF8.GetBytes(sheet.ToString());
        byte[] suffix = Encoding.UTF8.GetBytes("</t></is></c></row></sheetData></worksheet>");
        byte[] invalid = new byte[prefix.Length + 1 + suffix.Length];
        Buffer.BlockCopy(prefix, 0, invalid, 0, prefix.Length);
        invalid[prefix.Length] = 0x80; // continuation byte without a lead byte
        Buffer.BlockCopy(suffix, 0, invalid, prefix.Length + 1, suffix.Length);

        string path = TempFile();
        try
        {
            CreatePackage(path, [("Values", invalid)], stylesXml: null, sharedStringsXml: null);

            Assert.ThrowsAny<Exception>(() =>
            {
                using var reader = new XlsxOrXlsbReadOrEdit();
                reader.Open(path);
                reader.ActualSheetName = "Values";
                while (reader.Read())
                {
                }
            });

            Assert.ThrowsAny<Exception>(() =>
            {
                using var reader = new XlsxOrXlsbReadOrEdit { _disableFastXlsxReader = true };
                reader.Open(path);
                reader.ActualSheetName = "Values";
                while (reader.Read())
                {
                }
            });
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void TruncatedUtf8Worksheet_ThrowsLikeXmlReader()
    {
        byte[] prefix = Encoding.UTF8.GetBytes($$"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="{{MainNs}}"><sheetData>
            <row r="1"><c r="A1" t="inlineStr"><is><t>bad
            """);
        byte[] invalid = new byte[prefix.Length + 1];
        Buffer.BlockCopy(prefix, 0, invalid, 0, prefix.Length);
        invalid[^1] = 0xE2; // first byte of a three-byte sequence, cut off at EOF

        string path = TempFile();
        try
        {
            CreatePackage(path, [("Values", invalid)], stylesXml: null, sharedStringsXml: null);

            Assert.ThrowsAny<Exception>(() =>
            {
                using var reader = new XlsxOrXlsbReadOrEdit();
                reader.Open(path);
                reader.ActualSheetName = "Values";
                while (reader.Read())
                {
                }
            });

            Assert.ThrowsAny<Exception>(() =>
            {
                using var reader = new XlsxOrXlsbReadOrEdit { _disableFastXlsxReader = true };
                reader.Open(path);
                reader.ActualSheetName = "Values";
                while (reader.Read())
                {
                }
            });
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ExcelComFixtures_ValuesMatchManifest()
    {
        string dir = Path.Combine(FindRepositoryRoot(), "test_output", "fast-reader-fixtures");
        string manifestPath = Path.Combine(dir, "manifest.json");
        if (!File.Exists(manifestPath))
        {
            return; // run tools/Create-XlsxFastReaderFixtures.ps1 (requires Excel COM) to generate it
        }

        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(manifestPath));
        int checkedCells = 0;

        foreach (JsonElement file in manifest.RootElement.GetProperty("files").EnumerateArray())
        {
            string name = file.GetProperty("name").GetString()!;
            string path = Path.Combine(dir, name);
            Assert.True(File.Exists(path), $"missing fixture {path}");

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(path);
            string[] sheets = reader.GetSheetNames();

            foreach (JsonElement sheet in file.GetProperty("sheets").EnumerateArray())
            {
                string sheetName = sheet.GetProperty("name").GetString()!;
                Assert.Contains(sheetName, sheets);
                reader.ActualSheetName = sheetName;

                var rows = new List<object?[]>();
                while (reader.Read())
                {
                    var values = new object?[reader.FieldCount];
                    for (int i = 0; i < values.Length; i++)
                    {
                        values[i] = reader.GetValue(i);
                    }
                    rows.Add(values);
                }

                foreach (JsonElement cell in sheet.GetProperty("cells").EnumerateArray())
                {
                    int row = cell.GetProperty("row").GetInt32();
                    int column = cell.GetProperty("col").GetInt32();
                    Assert.True(row <= rows.Count, $"{name}/{sheetName}: row {row} missing (read {rows.Count})");

                    object? actual = column <= rows[row - 1].Length ? rows[row - 1][column - 1] : DBNull.Value;
                    string kind = cell.GetProperty("kind").GetString()!;
                    JsonElement expected = cell.GetProperty("value");

                    switch (kind)
                    {
                        case "string":
                            Assert.Equal(expected.GetString(), actual);
                            break;
                        case "number":
                            Assert.Equal(expected.GetDouble(), Convert.ToDouble(actual, CultureInfo.InvariantCulture), 9);
                            break;
                        case "bool":
                            Assert.Equal(expected.GetBoolean(), Convert.ToBoolean(actual, CultureInfo.InvariantCulture));
                            break;
                        case "date":
                            Assert.True(actual is DateTime, $"{name}/{sheetName} R{row}C{column}: expected date, got {Format(actual)}");
                            DateTime expectedDate = DateTime.Parse(expected.GetString()!, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
                            Assert.True(Math.Abs(((DateTime)actual).ToOADate() - expectedDate.ToOADate()) < 1e-6,
                                $"{name}/{sheetName} R{row}C{column}: {((DateTime)actual):O} vs {expectedDate:O}");
                            break;
                        case "error":
                            Assert.Equal("error in cell", actual);
                            break;
                        case "blank":
                            Assert.True(actual is DBNull or null, $"{name}/{sheetName} R{row}C{column}: expected blank, got {Format(actual)}");
                            break;
                        default:
                            Assert.Fail($"unknown manifest kind '{kind}'");
                            break;
                    }
                    checkedCells++;
                }
            }
        }

        Assert.True(checkedCells > 800, $"only {checkedCells} cells verified");
    }

    [Fact]
    public void WriterFixtures_ArePartOfTheCorpus()
    {
        string dir = Path.Combine(FindRepositoryRoot(), "test_output", "fast-reader-fixtures");
        if (!Directory.Exists(dir))
        {
            return; // the COM corpus has not been generated
        }

        DataTable table = BuildTypedTable(80);

        string shared = Path.Combine(dir, "ours-shared-strings.xlsx");
        using (var writer = new XlsxWriter(shared, useScharedStrings: true))
        {
            writer.AddSheet("Data");
            writer.WriteSheet(table.CreateDataReader());
        }

        string inline = Path.Combine(dir, "ours-inline-strings.xlsx");
        using (var writer = new XlsxWriter(inline, useScharedStrings: false))
        {
            writer.AddSheet("Data");
            writer.WriteSheet(table.CreateDataReader());
        }

        string sylvan = Path.Combine(dir, "sylvan-writer.xlsx");
        using (var writer = ExcelDataWriter.Create(sylvan))
        {
            writer.Write(table.CreateDataReader(), "Data");
        }

        AssertParity(shared);
        AssertParity(inline);
        AssertParity(sylvan);
        AssertTableValues(table, shared);
        AssertTableValues(table, inline);
    }

    [Fact]
    public void SwitchingSheetsMidRead_StillMatchesXmlReader()
    {
        AssertParityOnOwnWriterSheetSwitch();
    }

    private static void AssertParityOnOwnWriterSheetSwitch()
    {
        string path = Path.Combine(Path.GetTempPath(), $"sst_fast_switch_{Guid.NewGuid():N}.xlsx");
        try
        {
            DataTable table = BuildTypedTable(30);
            using (var writer = new XlsxWriter(path))
            {
                writer.AddSheet("A");
                writer.WriteSheet(table.CreateDataReader());
                writer.AddSheet("B");
                writer.WriteSheet(table.CreateDataReader());
            }

            using var fast = new XlsxOrXlsbReadOrEdit();
            fast.Open(path);

            fast.ActualSheetName = "A";
            Assert.True(fast.Read());
            object? first = fast.GetValue(0);

            fast.ActualSheetName = "B";
            Assert.True(fast.Read());
            Assert.Equal("Text", fast.GetValue(0));

            fast.ActualSheetName = "A";
            Assert.True(fast.Read());
            Assert.True(Equals(first, fast.GetValue(0)), "sheet switch must restart the sheet cleanly");

            // Full parity after a mid-stream switch: compare A against a fresh XmlReader read.
            // (Re-set the sheet: the row above already consumed A's first row on the fast reader.)
            fast.ActualSheetName = "B";
            fast.ActualSheetName = "A";
            using var slow = new XlsxOrXlsbReadOrEdit { _disableFastXlsxReader = true };
            slow.Open(path);
            slow.ActualSheetName = "A";

            int rows = 0;
            while (true)
            {
                bool fastRead = fast.Read();
                bool slowRead = slow.Read();
                Assert.True(fastRead == slowRead, $"Read() mismatch at row {rows}");
                if (!fastRead)
                {
                    break;
                }
                rows++;
                Assert.Equal(slow.FieldCount, fast.FieldCount);
                for (int c = 0; c < slow.FieldCount; c++)
                {
                    Assert.True(fast.GetExcelDataType(c) == slow.GetExcelDataType(c), $"row {rows} col {c}: type mismatch");
                    Assert.True(Equals(fast.GetValue(c), slow.GetValue(c)), $"row {rows} col {c}: value mismatch");
                }
            }
            Assert.True(rows > 0, "expected at least one row");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ExcelComFixtures_MatchXmlReader()
    {
        string dir = Path.Combine(FindRepositoryRoot(), "test_output", "fast-reader-fixtures");
        if (!Directory.Exists(dir))
        {
            return; // run tools/Create-XlsxFastReaderFixtures.ps1 to generate the COM corpus
        }

        string[] files = Directory.GetFiles(dir, "*.xlsx");
        Assert.NotEmpty(files);
        foreach (string file in files)
        {
            AssertParity(file);
        }
    }

    // ── Parity helpers ─────────────────────────────────────────────────────────

    private static void AssertParity(string path, Action<XlsxOrXlsbReadOrEdit>? configure = null)
    {
        using var fast = new XlsxOrXlsbReadOrEdit();
        using var slow = new XlsxOrXlsbReadOrEdit { _disableFastXlsxReader = true };
        configure?.Invoke(fast);
        configure?.Invoke(slow);

        fast.Open(path);
        slow.Open(path);

        Assert.Equal(slow.GetSheetNames(), fast.GetSheetNames());

        foreach (string sheet in slow.GetSheetNames())
        {
            fast.ActualSheetName = sheet;
            slow.ActualSheetName = sheet;

            int row = 0;
            while (true)
            {
                bool fastRead = fast.Read();
                bool slowRead = slow.Read();
                Assert.True(fastRead == slowRead, $"{sheet}: Read() mismatch at row {row} (fast={fastRead}, xml={slowRead})");
                if (!fastRead)
                {
                    break;
                }
                row++;

                Assert.True(slow.FieldCount == fast.FieldCount,
                    $"{sheet} row {row}: FieldCount {fast.FieldCount} vs {slow.FieldCount}");

                for (int column = 0; column < slow.FieldCount; column++)
                {
                    SpreadSheetTasks.ExcelDataType fastType = fast.GetExcelDataType(column);
                    SpreadSheetTasks.ExcelDataType slowType = slow.GetExcelDataType(column);
                    Assert.True(fastType == slowType,
                        $"{sheet} row {row} col {column}: type {fastType} vs {slowType}");

                    object? fastValue = fast.GetValue(column);
                    object? slowValue = slow.GetValue(column);
                    Assert.True(Equals(fastValue, slowValue),
                        $"{sheet} row {row} col {column}: value '{Format(fastValue)}' vs '{Format(slowValue)}'");
                }
            }

            // Read the dimension-backed row count after the sheet has been initialised.
            Assert.True(slow.RowCount == fast.RowCount,
                $"{sheet}: RowCount fast={fast.RowCount} xml={slow.RowCount}");
        }
    }

    private static string Format(object? value) => value switch
    {
        null => "<null>",
        DBNull => "<DBNull>",
        DateTime dateTime => dateTime.ToString("O", CultureInfo.InvariantCulture),
        double number => number.ToString("R", CultureInfo.InvariantCulture),
        string text => text.Length > 100 ? text[..100] + "..." : text,
        _ => value.ToString() ?? "<null>",
    };

    private static List<object?[]> ReadAll(XlsxOrXlsbReadOrEdit reader)
    {
        var rows = new List<object?[]>();
        while (reader.Read())
        {
            var values = new object?[reader.FieldCount];
            for (int column = 0; column < values.Length; column++)
            {
                values[column] = reader.GetValue(column);
            }
            rows.Add(values);
        }
        return rows;
    }

    private static void AssertRowsEqual(IReadOnlyList<object?[]> expected, IReadOnlyList<object?[]> actual)
    {
        Assert.Equal(expected.Count, actual.Count);
        for (int row = 0; row < expected.Count; row++)
        {
            Assert.Equal(expected[row].Length, actual[row].Length);
            for (int column = 0; column < expected[row].Length; column++)
            {
                Assert.True(Equals(expected[row][column], actual[row][column]),
                    $"row {row + 1} col {column}: '{Format(expected[row][column])}' vs '{Format(actual[row][column])}'");
            }
        }
    }

    private static void AssertRow(IReadOnlyList<object?> expected, IReadOnlyList<object?> actual)
    {
        Assert.Equal(expected.Count, actual.Count);
        for (int column = 0; column < expected.Count; column++)
        {
            Assert.True(Equals(expected[column], actual[column]),
                $"col {column}: '{Format(expected[column])}' vs '{Format(actual[column])}'");
        }
    }

    private static void AssertTableValues(DataTable table, string path)
    {
        using var reader = new XlsxOrXlsbReadOrEdit();
        reader.Open(path);
        reader.ActualSheetName = reader.GetSheetNames()[0];

        Assert.True(reader.Read()); // header
        for (int c = 0; c < table.Columns.Count; c++)
        {
            Assert.Equal(table.Columns[c].ColumnName, reader.GetValue(c));
        }

        for (int r = 0; r < table.Rows.Count; r++)
        {
            Assert.True(reader.Read(), $"row {r + 1} missing");
            for (int c = 0; c < table.Columns.Count; c++)
            {
                AssertExpectedCell(table.Rows[r][c], table.Columns[c].DataType, reader.GetValue(c), r, c);
            }
        }
        Assert.False(reader.Read());
    }

    private static void AssertExpectedCell(object expected, Type columnType, object? actual, int row, int column)
    {
        if (expected is DBNull)
        {
            Assert.Equal(DBNull.Value, actual);
            return;
        }

        if (columnType == typeof(int) || columnType == typeof(long) || columnType == typeof(short) || columnType == typeof(byte))
        {
            Assert.Equal(Convert.ToInt64(expected, CultureInfo.InvariantCulture), Convert.ToInt64(actual, CultureInfo.InvariantCulture));
        }
        else if (columnType == typeof(double) || columnType == typeof(float) || columnType == typeof(decimal))
        {
            // The reader reports "0" without a decimal point as Int64, so compare numerically.
            Assert.Equal(Convert.ToDouble(expected, CultureInfo.InvariantCulture), Convert.ToDouble(actual, CultureInfo.InvariantCulture));
        }
        else if (columnType == typeof(bool))
        {
            Assert.Equal(expected, actual);
        }
        else if (columnType == typeof(DateTime))
        {
            Assert.Equal(expected, actual);
        }
        else
        {
            Assert.True(Equals(expected, actual), $"row {row} col {column}: '{Format(expected)}' vs '{Format(actual)}'");
        }
    }

    private static DataTable BuildTypedTable(int rows)
    {
        var table = new DataTable();
        table.Columns.Add("Text", typeof(string));
        table.Columns.Add("Int", typeof(int));
        table.Columns.Add("Long", typeof(long));
        table.Columns.Add("Double", typeof(double));
        table.Columns.Add("Decimal", typeof(decimal));
        table.Columns.Add("When", typeof(DateTime));
        table.Columns.Add("Flag", typeof(bool));

        string[] tricky =
        [
            "plain",
            "A&B <tag> \"quoted\" 'apos'",
            "line\nbreak",
            "carriage\rreturn",
            "  leading and trailing  ",
            "Zażółć gęślą jaźń",
            "日本語テキスト",
            "emoji 😀🚀",
            "_x0041_ escape",
            new string('x', 20_000),
            "0123456789",
        ];

        for (int r = 0; r < rows; r++)
        {
            table.Rows.Add(
                tricky[r % tricky.Length],
                r - rows / 2,
                (long)r * 1_000_003,
                r == 0 ? 1.5 : r * -0.125,
                r == 0 ? 0m : (decimal)r / 7m,
                new DateTime(1900, 1, 1).AddDays(r * 37).AddSeconds(r),
                r % 2 == 0);
        }

        table.Rows[1]["Double"] = DBNull.Value;
        table.Rows[2]["When"] = DBNull.Value;
        table.Rows[3]["Text"] = DBNull.Value;
        return table;
    }

    // ── Handcrafted SpreadsheetML corpus ───────────────────────────────────────

    private static IEnumerable<(string Name, byte[] SheetXml, string? Styles, string? Strings, int ExpectedRows)> HandcraftedCases()
    {
        string styles = StylesXml();

        yield return ("NoRowRefs", Encoding.UTF8.GetBytes(SheetXml("""
            <row><c><v>1</v></c><c t="s"><v>0</v></c><c><v>2.5</v></c></row>
            <row><c><v>3</v></c><c t="s"><v>1</v></c><c><v>4.5</v></c></row>
            """)), styles, StringsXml("one", "two"), 2);

        yield return ("WhitespaceAndComments", Encoding.UTF8.GetBytes(SheetXml("""
              <!-- before rows -->
              <row r="1">
                <c r="A1"><v>1</v></c>
                <!-- inside row -->
                <c r="B1" t="s"><v>0</v></c>
              </row>
              <?pi inside?>
              <row r="2">
                <c r="A2"><v>2</v></c>
              </row>
            """)), styles, StringsXml("text"), 2);

        yield return ("SparseRows", Encoding.UTF8.GetBytes(SheetXml("""
            <row r="1"><c r="A1"><v>1</v></c></row>
            <row r="3"><c r="A3"><v>3</v></c></row>
            <row r="6"><c r="A6"><v>6</v></c></row>
            """)), styles, null, 6);

        yield return ("WideColumns", Encoding.UTF8.GetBytes(SheetXml(BuildWideRow())), styles, null, 2);

        yield return ("SelfClosingAndEmptyValues", Encoding.UTF8.GetBytes(SheetXml("""
            <row r="1">
              <c r="A1"/>
              <c r="B1" s="1"/>
              <c r="C1"><v/></c>
              <c r="D1"><v></v></c>
              <c r="E1" t="inlineStr"><is/></c>
              <c r="F1" t="inlineStr"><is><t/></is></c>
            </row>
            """)), styles, null, 1);

        yield return ("FormulasAndErrors", Encoding.UTF8.GetBytes(SheetXml("""
            <row r="1">
              <c r="A1"><f>1+1</f><v>2</v></c>
              <c r="B1" t="str"><f>"x"</f><v>x</v></c>
              <c r="C1" t="e"><v>#DIV/0!</v></c>
              <c r="D1" t="b"><f>TRUE()</f><v>1</v></c>
            </row>
            """)), styles, null, 1);

        yield return ("InlineStrings", Encoding.UTF8.GetBytes(SheetXml("""
            <row r="1">
              <c r="A1" t="inlineStr"><is><t>simple</t></is></c>
              <c r="B1" t="inlineStr"><is><t xml:space="preserve"> padded </t></is></c>
              <c r="C1" t="inlineStr"><is><t>uni zażółć 😀</t></is></c>
              <c r="D1" t="inlineStr"><is><t>a&amp;b&lt;c&gt;d&quot;e&apos;f</t></is></c>
            </row>
            """)), styles, null, 1);

        yield return ("RichInlineStrings", Encoding.UTF8.GetBytes(SheetXml("""
            <row r="1">
              <c r="A1" t="inlineStr"><is><r><t>first</t></r><r><t>second</t></r></is></c>
            </row>
            """)), styles, null, 1);

        yield return ("EntityStrings", Encoding.UTF8.GetBytes(SheetXml("""
            <row r="1">
              <c r="A1" t="s"><v>0</v></c>
              <c r="B1" t="s"><v>1</v></c>
              <c r="C1" t="inlineStr"><is><t>&#65;&#x42;&amp;</t></is></c>
            </row>
            """)), styles, StringsXml("A&amp;B", "&lt;x&gt;"), 1);

        yield return ("DateStyles", Encoding.UTF8.GetBytes(SheetXml("""
            <row r="1">
              <c r="A1" s="1"><v>45000</v></c>
              <c r="B1" s="2"><v>45000.5</v></c>
              <c r="C1" s="3"><v>45001</v></c>
              <c r="D1" s="4"><v>12.5</v></c>
            </row>
            """)), styles, null, 1);

        yield return ("LargeNumbers", Encoding.UTF8.GetBytes(SheetXml("""
            <row r="1">
              <c r="A1"><v>9007199254740993</v></c>
              <c r="B1"><v>-1234567890</v></c>
              <c r="C1"><v>1.5E+300</v></c>
              <c r="D1"><v>1E-300</v></c>
              <c r="E1"><v>0.0000001</v></c>
            </row>
            """)), styles, null, 1);

        yield return ("EmptyFirstRow", Encoding.UTF8.GetBytes(SheetXml("""
            <row r="1"/>
            <row r="2"><c r="A2"><v>7</v></c></row>
            """)), styles, null, 2);

        yield return ("UnknownElements", Encoding.UTF8.GetBytes($$"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="{{MainNs}}">
              <sheetData>
                <extLst><ext uri="x"><foo>ignored</foo></ext></extLst>
                <row r="1"><c r="A1"><v>1</v></c><extLst><ext/></extLst><c r="B1"><v>2</v></c></row>
                <row r="2"><c r="A2"><v>3</v></c></row>
              </sheetData>
            </worksheet>
            """), styles, null, 2);

        // Excel can save parts as UTF-16; the scanner must decline and let XmlReader decode them.
        yield return ("Utf16Sheet", Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(SheetXml("""
            <row r="1"><c r="A1"><v>5</v></c><c t="s"><v>0</v></c></row>
            """).Replace("encoding=\"UTF-8\"", "encoding=\"UTF-16\""))).ToArray(), styles, StringsXml("wide"), 1);

        yield return ("Utf8Bom", Encoding.UTF8.GetPreamble().Concat(Encoding.UTF8.GetBytes(SheetXml("""
            <row r="1"><c r="A1"><v>5</v></c></row>
            """))).ToArray(), styles, null, 1);

        // The XmlReader path matches qualified names ("x:row"), so prefixed parts yield no rows.
        // The scanner declines them and must reproduce that behaviour exactly.
        yield return ("PrefixedTags", Encoding.UTF8.GetBytes($$"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <x:worksheet xmlns:x="{{MainNs}}">
              <x:sheetData>
                <x:row r="1"><x:c r="A1"><x:v>5</x:v></x:c></x:row>
              </x:sheetData>
            </x:worksheet>
            """), styles, null, 0);

        yield return ("EmptySheetData", Encoding.UTF8.GetBytes($$"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="{{MainNs}}"><sheetData/></worksheet>
            """), styles, null, 0);

        yield return ("SpecialWhitespaceStrings", Encoding.UTF8.GetBytes(SheetXml("""
            <row r="1">
              <c r="A1" t="inlineStr"><is><t xml:space="preserve">  two  spaces  </t></is></c>
              <c r="B1" t="inlineStr"><is><t xml:space="preserve">tab	inside</t></is></c>
              <c r="C1" t="inlineStr"><is><t>_x0041_ literal</t></is></c>
              <c r="D1" t="s"><v>0</v></c>
            </row>
            """)), styles, StringsXml("_x000D_"), 1);

        // r, s and t are written in a different order than Excel uses; both paths must read the
        // type and the style, not parse the shared-string index as a number.
        yield return ("AttributeOrder", Encoding.UTF8.GetBytes(SheetXml("""
            <row r="1">
              <c t="s" r="A1"><v>0</v></c>
              <c s="1" r="B1"><v>45000</v></c>
              <c t="inlineStr" r="C1"><is><t>inline</t></is></c>
              <c r="D1" t="s"><v>0</v></c>
            </row>
            """)), styles, StringsXml("shared"), 1);

        // A whitespace-only inline string is only significant with xml:space="preserve";
        // XmlReader drops it otherwise and the scanner must agree.
        yield return ("WhitespaceOnlyInlineStrings", Encoding.UTF8.GetBytes(SheetXml("""
            <row r="1">
              <c r="A1" t="inlineStr"><is><t>   </t></is></c>
              <c r="B1" t="inlineStr"><is><t xml:space="preserve">   </t></is></c>
              <c r="C1" t="inlineStr"><is xml:space="preserve"><t>  </t></is></c>
              <c r="D1" t="inlineStr"><is><t>text</t></is></c>
            </row>
            """)), styles, null, 1);

        yield return ("WhitespaceRootPreserve", Encoding.UTF8.GetBytes($$"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="{{MainNs}}" xml:space="preserve">
              <sheetData>
                <row r="1"><c r="A1" t="inlineStr"><is><t>  </t></is></c><c r="B1"><v>1</v></c></row>
              </sheetData>
            </worksheet>
            """), styles, null, 1);

        // A comment or CDATA section can contain the literal "</row>"; the row boundary must
        // still be the real closing tag, and the cell then takes the XmlReader fallback.
        yield return ("RowEndingInsideCommentAndCData", Encoding.UTF8.GetBytes(SheetXml("""
            <row r="1">
              <!-- fake </row> here -->
              <?fake </row> pi?>
              <c r="A1"><v>1</v></c>
            </row>
            <row r="2">
              <c r="A2" t="inlineStr"><is><t><![CDATA[cd </row> text]]></t></is></c>
            </row>
            """)), styles, null, 2);

        // Stray text and CDATA inside sheetData make the scanner hand the part to XmlReader.
        // Before the initialization guard this re-entered the scanner and overflowed the stack,
        // even though no row had been produced yet.
        yield return ("StrayTextAndCDataInSheetData", Encoding.UTF8.GetBytes($$"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="{{MainNs}}">
              <sheetData>
                text before the first row
                <row r="1"><c r="A1"><v>1</v></c></row>
                <![CDATA[cdata between rows]]>
                <row r="2"><c r="A2"><v>2</v></c></row>
              </sheetData>
            </worksheet>
            """), styles, null, 2);
    }

    private static string BuildWideRow()
    {
        var builder = new StringBuilder("<row r=\"1\">");
        for (int c = 0; c < 60; c++)
        {
            builder.Append($"<c r=\"{ColumnName(c)}1\"><v>{c}</v></c>");
        }
        builder.Append("</row><row r=\"2\">");
        for (int c = 0; c < 60; c += 2)
        {
            builder.Append($"<c r=\"{ColumnName(c)}2\"><v>{c * 10}</v></c>");
        }
        builder.Append("</row>");
        return builder.ToString();
    }

    private static string ColumnName(int index)
    {
        string name = string.Empty;
        index++;
        while (index > 0)
        {
            int remainder = (index - 1) % 26;
            name = (char)('A' + remainder) + name;
            index = (index - 1) / 26;
        }
        return name;
    }

    private static string SheetXml(string body) => $$"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="{{MainNs}}">
          <sheetData>
        {{body}}
          </sheetData>
        </worksheet>
        """;

    private static string StringsXml(params string[] values)
    {
        var builder = new StringBuilder();
        builder.Append($$"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="{{MainNs}}" count="{{values.Length}}" uniqueCount="{{values.Length}}">""");
        foreach (string value in values)
        {
            builder.Append("<si><t xml:space=\"preserve\">").Append(value).Append("</t></si>");
        }
        builder.Append("</sst>");
        return builder.ToString();
    }

    private static string StylesXml() => $$"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="{{MainNs}}">
          <numFmts count="1"><numFmt numFmtId="165" formatCode="yyyy\-mm\-dd"/></numFmts>
          <cellXfs count="5">
            <xf numFmtId="0" xfId="0"/>
            <xf numFmtId="14" xfId="0" applyNumberFormat="1"/>
            <xf numFmtId="22" xfId="0" applyNumberFormat="1"/>
            <xf numFmtId="165" xfId="0" applyNumberFormat="1"/>
            <xf numFmtId="0" xfId="0"/>
          </cellXfs>
        </styleSheet>
        """;

    // ── Minimal XLSX package writer for the handcrafted tests ────────────────

    private static void CreatePackage(
        string path,
        IReadOnlyList<(string Name, byte[] Xml)> sheets,
        string? stylesXml,
        string? sharedStringsXml,
        string? workbookPrXml = null)
    {
        using var file = File.Create(path);
        using var zip = new ZipArchive(file, ZipArchiveMode.Create);

        var contentTypes = new StringBuilder();
        contentTypes.Append($$"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>""");
        for (int i = 0; i < sheets.Count; i++)
        {
            contentTypes.Append($"""<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>""");
        }
        if (stylesXml is not null)
        {
            contentTypes.Append("""<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>""");
        }
        if (sharedStringsXml is not null)
        {
            contentTypes.Append("""<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>""");
        }
        contentTypes.Append("</Types>");
        WriteEntry(zip, "[Content_Types].xml", contentTypes.ToString());

        WriteEntry(zip, "_rels/.rels", $$"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>""");

        var workbook = new StringBuilder();
        workbook.Append($$"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="{{MainNs}}" xmlns:r="{{RelNs}}">""");
        if (workbookPrXml is not null)
        {
            workbook.Append(workbookPrXml);
        }
        workbook.Append("<sheets>");
        for (int i = 0; i < sheets.Count; i++)
        {
            workbook.Append($"""<sheet name="{sheets[i].Name}" sheetId="{i + 1}" r:id="rId{i + 1}"/>""");
        }
        workbook.Append("</sheets></workbook>");
        WriteEntry(zip, "xl/workbook.xml", workbook.ToString());

        var rels = new StringBuilder();
        rels.Append($$"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">""");
        for (int i = 0; i < sheets.Count; i++)
        {
            rels.Append($"""<Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i + 1}.xml"/>""");
        }
        if (stylesXml is not null)
        {
            rels.Append("""<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>""");
        }
        if (sharedStringsXml is not null)
        {
            rels.Append("""<Relationship Id="rIdStrings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>""");
        }
        rels.Append("</Relationships>");
        WriteEntry(zip, "xl/_rels/workbook.xml.rels", rels.ToString());

        for (int i = 0; i < sheets.Count; i++)
        {
            WriteEntry(zip, $"xl/worksheets/sheet{i + 1}.xml", sheets[i].Xml);
        }
        if (stylesXml is not null)
        {
            WriteEntry(zip, "xl/styles.xml", stylesXml);
        }
        if (sharedStringsXml is not null)
        {
            WriteEntry(zip, "xl/sharedStrings.xml", sharedStringsXml);
        }
    }

    private static void WriteEntry(ZipArchive zip, string name, string content)
        => WriteEntry(zip, name, Encoding.UTF8.GetBytes(content));

    private static void WriteEntry(ZipArchive zip, string name, byte[] content)
    {
        ZipArchiveEntry entry = zip.CreateEntry(name, CompressionLevel.Fastest);
        using Stream stream = entry.Open();
        stream.Write(content, 0, content.Length);
    }

    private static string FindRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "source", "SpreadSheetTasks.slnx")))
            {
                return directory.FullName;
            }
            if (File.Exists(Path.Combine(directory.FullName, "SpreadSheetTasks.slnx")))
            {
                return directory.Parent?.FullName ?? directory.FullName;
            }
            directory = directory.Parent;
        }
        return AppContext.BaseDirectory;
    }
}
