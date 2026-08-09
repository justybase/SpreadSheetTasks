using SpreadSheetTasks;

namespace Tests;

[Collection("Sequential")]
public class AotCompatibilityTests
{
    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void ListWriter_UsesStaticTypeCodeMapping(string extension)
    {
        string path = Path.Combine(Path.GetTempPath(), "SpreadSheetTasks-aot-test-" + Guid.NewGuid().ToString("N") + extension);

        try
        {
            var headers = new List<string> { "Text", "Number", "Amount", "Date" };
            var typeCodes = new List<TypeCode>
            {
                TypeCode.String,
                TypeCode.Int32,
                TypeCode.Double,
                TypeCode.DateTime
            };
            var rows = new List<object?[]>
            {
                new object?[] { "AOT", 10, 2.5, new DateTime(2025, 2, 3) }
            };

            using (var writer = ExcelWriter.CreateWriter(path))
            {
                writer.AddSheet("Data");
                writer.WriteSheet(headers, typeCodes, rows);
            }

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(path);
            reader.ActualSheetName = "Data";

            Assert.True(reader.Read());
            Assert.True(reader.Read());
            Assert.Equal("AOT", reader.GetString(0));
            Assert.Equal(10, reader.GetInt32(1));
            Assert.Equal(2.5, reader.GetDouble(2));
            Assert.Equal(new DateTime(2025, 2, 3), reader.GetDateTime(3));
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }
}
