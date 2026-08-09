using SpreadSheetTasks;
using System.Data;

namespace AotSmoke;

internal static class Program
{
    private static int Main()
    {
        string workDirectory = Path.Combine(
            Path.GetTempPath(),
            "SpreadSheetTasks-AotSmoke-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workDirectory);

        try
        {
            RunListRoundTrip(workDirectory, ".xlsx");
            RunListRoundTrip(workDirectory, ".xlsb");
            RunDataTableRoundTrip(workDirectory);
            RunStreamWriterSmokeTest();
            Console.WriteLine("NativeAOT smoke test passed.");
            return 0;
        }
        finally
        {
            if (Directory.Exists(workDirectory))
                Directory.Delete(workDirectory, recursive: true);
        }
    }

    private static void RunListRoundTrip(string workDirectory, string extension)
    {
        string path = Path.Combine(workDirectory, "list" + extension);
        var headers = new List<string> { "Name", "Count", "Amount", "When" };
        var typeCodes = new List<TypeCode>
        {
            TypeCode.String,
            TypeCode.Int32,
            TypeCode.Double,
            TypeCode.DateTime
        };
        var rows = new List<object?[]>
        {
            new object?[] { "Alice", 42, 12.5, new DateTime(2025, 1, 2) },
            new object?[] { "Bob", 7, 99.25, new DateTime(2025, 1, 3) }
        };

        using (var writer = ExcelWriter.CreateWriter(path))
        {
            writer.AddSheet("Data");
            writer.WriteSheet(headers, typeCodes, rows, headers: true);
        }

        using var reader = new XlsxOrXlsbReadOrEdit();
        reader.Open(path);
        reader.ActualSheetName = "Data";

        Check(reader.Read(), $"Missing header row in {extension}.");
        Check(reader.GetString(0) == "Name", $"Unexpected header in {extension}.");
        Check(reader.Read(), $"Missing first data row in {extension}.");
        Check(reader.GetString(0) == "Alice", $"Unexpected string value in {extension}.");
        Check(reader.GetInt32(1) == 42, $"Unexpected integer value in {extension}.");
        Check(Math.Abs(reader.GetDouble(2) - 12.5) < 0.000001, $"Unexpected double value in {extension}.");
        Check(reader.GetDateTime(3) == new DateTime(2025, 1, 2), $"Unexpected date value in {extension}.");
    }

    private static void RunDataTableRoundTrip(string workDirectory)
    {
        string path = Path.Combine(workDirectory, "datatable.xlsx");
        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Active", typeof(bool));
        table.Rows.Add("NativeAOT", true);

        using (var writer = new XlsxWriter(path))
        {
            writer.AddSheet("Data");
            writer.WriteSheet(table.CreateDataReader());
        }

        using var reader = new XlsxOrXlsbReadOrEdit();
        reader.Open(path);
        reader.ActualSheetName = "Data";
        Check(reader.Read(), "Missing DataTable header row.");
        Check(reader.Read(), "Missing DataTable data row.");
        Check(reader.GetString(0) == "NativeAOT", "Unexpected DataTable string value.");
        Check(reader.GetValue(1) is bool value && value, "Unexpected DataTable boolean value.");
    }

    private static void RunStreamWriterSmokeTest()
    {
        using var stream = new MemoryStream();
        using (var writer = new XlsxWriter(stream, leaveExcelArchiveOpen: true))
        {
            writer.AddSheet("Data");
            writer.WriteSheet(new[] { "NativeAOT" });
        }

        Check(stream.Length > 0, "Stream writer produced no output.");
    }

    private static void Check(bool condition, string message)
    {
        if (!condition)
            throw new InvalidOperationException(message);
    }
}
