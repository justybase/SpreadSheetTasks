using SpreadSheetTasks;
using System.Data;
using System.Security.Cryptography;
using System.Text;

namespace Tests;

[Collection("Sequential")]
public class XlsbStressTests
{
    private const string PathXlsbLarge = "stress_large.xlsb";
    private const string PathXlsbNumbers = "stress_numbers.xlsb";
    private const string PathXlsbSpecial = "stress_special.xlsb";
    private const string PathXlsbSheets = "stress_sheets.xlsb";
    private const string PathXlsxLarge = "stress_large.xlsx";
    private const string PathXlsbMixed = "stress_mixed.xlsb";
    private const string PathXlsxMixed = "stress_mixed.xlsx";

    private static DataTable CreateLargeTable(int rows)
    {
        DataTable dt = new DataTable();
        dt.Columns.Add("ID", typeof(long));
        dt.Columns.Add("INT", typeof(int));
        dt.Columns.Add("DBL", typeof(double));
        dt.Columns.Add("TXT", typeof(string));
        dt.Columns.Add("FLAG", typeof(bool));
        dt.Columns.Add("DATE", typeof(DateTime));
        var rnd = new Random(12345);
        for (int i = 0; i < rows; i++)
        {
            dt.Rows.Add(
                (long)i * 1_000_003L - 500_000L,
                rnd.Next(int.MinValue, int.MaxValue),
                rnd.NextDouble() * 1e9,
                "TXT_" + i + "_" + rnd.Next(0, 1_000_000),
                i % 2 == 0,
                new DateTime(2000 + (i % 40), 1 + (i % 12), 1 + (i % 27)));
        }
        return dt;
    }

    private static DataTable CreateNumbersTable()
    {
        DataTable dt = new DataTable();
        dt.Columns.Add("LONG", typeof(long));
        dt.Columns.Add("DBL", typeof(double));
        object[][] rows =
        [
            [long.MinValue, -1.7976931348623157E+308],
            [long.MaxValue, 1.7976931348623157E+308],
            [0L, 0.0],
            [-1L, -0.0],
            [1L, 1e-300],
            [-9_223_372_036_854_775_807L, double.MaxValue / 2],
            [9_223_372_036_854_775_806L, -double.MaxValue / 3],
            [123_456_789_012_345L, 3.141592653589793],
            [-123_456_789_012_345L, -3.141592653589793],
        ];
        foreach (var r in rows)
        {
            dt.Rows.Add(r);
        }
        return dt;
    }

    private static DataTable CreateSpecialCharsTable()
    {
        DataTable dt = new DataTable();
        dt.Columns.Add("TXT", typeof(string));
        var sb = new StringBuilder();
        sb.Append("base-");
        sb.Append('\u0001');
        sb.Append('\u0008');
        sb.Append('\u000B');
        sb.Append('\u000C');
        sb.Append('\u000E');
        sb.Append('\u001F');
        sb.Append("xml:<>&\"'");
        sb.Append("polish:zażółć gęślą jaźń");
        sb.Append("emoji:😀🚀👨‍👩‍👧‍👦");
        sb.Append("surrogates:\uD83D\uDE00\uD800\uDC00");
        sb.Append("rt:\r\nlf:\ncr:\r");
        sb.Append("tab:\t");
        sb.Append(new string('X', 20_000));
        sb.Append("-end");
        string huge = sb.ToString();

        dt.Rows.Add("");
        dt.Rows.Add("   ");
        dt.Rows.Add("null-like");
        dt.Rows.Add(huge);
        dt.Rows.Add(new string(' ', 100_000));
        dt.Rows.Add("a<b&c>d\"e'f");
        return dt;
    }

    private static DataTable CreateXmlSafeCharsTable()
    {
        DataTable dt = new DataTable();
        dt.Columns.Add("TXT", typeof(string));
        dt.Rows.Add("");
        dt.Rows.Add("   ");
        dt.Rows.Add("polish:zażółć gęślą jaźń");
        dt.Rows.Add("emoji:😀🚀👨‍👩‍👧‍👦");
        dt.Rows.Add("tab:\ttab-end");
        dt.Rows.Add("lf:\nlf-end");
        dt.Rows.Add(new string('Y', 20_000));
        return dt;
    }

    private static string DumpFile(string path, bool streaming = false)
    {
        using XlsxOrXlsbReadOrEdit excelFile = new XlsxOrXlsbReadOrEdit();
        excelFile.UseMemoryStreamInXlsb = !streaming;
        excelFile.Open(path);
        var sheetNames = excelFile.GetScheetNames();
        excelFile.ActualSheetName = sheetNames[0];
        var sb = new StringBuilder();
        object[]? row = null;
        while (excelFile.Read())
        {
            row ??= new object[excelFile.FieldCount];
            excelFile.GetValues(row);
            sb.AppendLine(string.Join('|', row));
        }
        return sb.ToString().Trim();
    }

    private static string Sha256Of(string content)
    {
        byte[] hash = SHA256.HashData(Encoding.UTF8.GetBytes(content));
        return Convert.ToHexString(hash);
    }

    private static void WriteFile(string path, DataTable dt)
    {
        using (var writer = ExcelWriter.CreateWriter(path))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }
    }

    [Fact]
    public void Stress_LargeFile_MemoryVsStreaming_Identical()
    {
        var dt = CreateLargeTable(100_000);
        WriteFile(PathXlsbLarge, dt);
        try
        {
            string memoryDump = DumpFile(PathXlsbLarge, streaming: false);
            string streamingDump = DumpFile(PathXlsbLarge, streaming: true);

            Assert.False(string.IsNullOrEmpty(memoryDump));
            Assert.Equal(100_001, memoryDump.Split('\n').Length);
            Assert.Equal(memoryDump, streamingDump);
        }
        finally
        {
            File.Delete(PathXlsbLarge);
        }
    }

    [Fact]
    public void Stress_ExtremeNumbers_RoundTrip_AllModes()
    {
        var dt = CreateNumbersTable();
        WriteFile(PathXlsbNumbers, dt);
        try
        {
            string memoryDump = DumpFile(PathXlsbNumbers, streaming: false);
            string streamingDump = DumpFile(PathXlsbNumbers, streaming: true);

            Assert.Contains(long.MinValue.ToString(), memoryDump);
            Assert.Contains(long.MaxValue.ToString(), memoryDump);
            Assert.Equal(memoryDump, streamingDump);
        }
        finally
        {
            File.Delete(PathXlsbNumbers);
        }
    }

    [Fact]
    public void Stress_SpecialCharacters_RoundTrip_XlsbOnly()
    {
        var dt = CreateSpecialCharsTable();
        WriteFile(PathXlsbSpecial, dt);
        try
        {
            string xlsbDump = DumpFile(PathXlsbSpecial, streaming: false);
            string xlsbDumpStreaming = DumpFile(PathXlsbSpecial, streaming: true);

            Assert.Contains("zażółć gęślą jaźń", xlsbDump);
            Assert.Contains("😀", xlsbDump);
            Assert.Contains("a<b&c>d\"e'f", xlsbDump);
            Assert.Contains('\u0001', xlsbDump);
            Assert.Contains('\u001F', xlsbDump);
            Assert.Equal(xlsbDump, xlsbDumpStreaming);
        }
        finally
        {
            File.Delete(PathXlsbSpecial);
        }
    }

    [Fact]
    public void Stress_XmlSafeSpecialChars_RoundTrip_XlsbVsXlsx()
    {
        var dt = CreateXmlSafeCharsTable();
        WriteFile(PathXlsbSpecial, dt);
        string xlsxPath = PathXlsbSpecial.Replace(".xlsb", ".xlsx");
        WriteFile(xlsxPath, dt);
        try
        {
            string xlsbDump = DumpFile(PathXlsbSpecial, streaming: false);
            string xlsbDumpStreaming = DumpFile(PathXlsbSpecial, streaming: true);
            string xlsxDump = DumpFile(xlsxPath, streaming: false);

            Assert.Contains("zażółć gęślą jaźń", xlsbDump);
            Assert.Contains("😀", xlsbDump);
            Assert.Equal(xlsbDump, xlsbDumpStreaming);
            Assert.Equal(xlsxDump, xlsbDump);
        }
        finally
        {
            File.Delete(PathXlsbSpecial);
            File.Delete(xlsxPath);
        }
    }

    [Fact]
    public void Stress_MultipleSheets_AllRead_Consistent()
    {
        var dt1 = CreateLargeTable(5_000);
        var dt2 = CreateLargeTable(3_000);
        using (var writer = ExcelWriter.CreateWriter(PathXlsbSheets))
        {
            writer.AddSheet("First");
            writer.WriteSheet(dt1.CreateDataReader());
            writer.AddSheet("Second");
            writer.WriteSheet(dt2.CreateDataReader());
            writer.AddSheet("Third");
            writer.WriteSheet(CreateSpecialCharsTable().CreateDataReader());
        }

        try
        {
            using XlsxOrXlsbReadOrEdit excelFile = new XlsxOrXlsbReadOrEdit();
            excelFile.Open(PathXlsbSheets);
            var names = excelFile.GetScheetNames();
            Assert.Equal(3, names.Length);

            excelFile.ActualSheetName = names[0];
            int count1 = 0;
            while (excelFile.Read())
            {
                count1++;
            }

            excelFile.ActualSheetName = names[1];
            int count2 = 0;
            while (excelFile.Read())
            {
                count2++;
            }

            excelFile.ActualSheetName = names[2];
            int count3 = 0;
            while (excelFile.Read())
            {
                count3++;
            }

            Assert.Equal(5_001, count1);
            Assert.Equal(3_001, count2);
            Assert.Equal(7, count3);
        }
        finally
        {
            File.Delete(PathXlsbSheets);
        }
    }

    [Fact]
    public void Stress_LargeMixedData_XlsbVsXlsx_Identical()
    {
        var dt = CreateLargeTable(30_000);
        WriteFile(PathXlsbMixed, dt);
        WriteFile(PathXlsxMixed, dt);
        try
        {
            string xlsbDump = DumpFile(PathXlsbMixed, streaming: false);
            string xlsbDumpStreaming = DumpFile(PathXlsbMixed, streaming: true);
            string xlsxDump = DumpFile(PathXlsxMixed, streaming: false);

            Assert.Equal(30_001, xlsbDump.Split('\n').Length);
            Assert.Equal(xlsbDump, xlsbDumpStreaming);
            Assert.Equal(xlsxDump, xlsbDump);
        }
        finally
        {
            File.Delete(PathXlsbMixed);
            File.Delete(PathXlsxMixed);
        }
    }

    [Fact]
    public void Stress_RepeatedReads_Deterministic()
    {
        var dt = CreateLargeTable(10_000);
        WriteFile(PathXlsbLarge, dt);
        try
        {
            string first = DumpFile(PathXlsbLarge, streaming: false);
            string second = DumpFile(PathXlsbLarge, streaming: false);
            string streaming = DumpFile(PathXlsbLarge, streaming: true);
            string streaming2 = DumpFile(PathXlsbLarge, streaming: true);

            Assert.Equal(first, second);
            Assert.Equal(streaming, streaming2);
            Assert.Equal(Sha256Of(first), Sha256Of(second));
        }
        finally
        {
            File.Delete(PathXlsbLarge);
        }
    }
}
