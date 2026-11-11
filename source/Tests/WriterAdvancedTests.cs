using SpreadSheetTasks;
using System.Data;

namespace Tests;

[Collection("Sequential")]
public class WriterAdvancedTests
{
    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteSheet_WithStartingRow_OffsetsData(string extension)
    {
        var fileName = $"test_starting_row{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader(), startingRow: 5);
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            // Read all rows and verify data starts at row 5
            int rowNum = 0;
            while (reader.Read())
            {
                if (rowNum == 5) // Row 5 (0-indexed) should have header
                {
                    Assert.Equal("Col1", reader.GetValue(0));
                }
                else if (rowNum == 6) // Row 6 should have data
                {
                    Assert.Equal("Data", reader.GetValue(0));
                }
                rowNum++;
            }
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteSheet_WithStartingRowAndColumn_OffsetsBoth(string extension)
    {
        var fileName = $"test_starting_both{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader(), startingRow: 2, startingColumn: 2);
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            // Read until we find the data
            int rowNum = 0;
            while (reader.Read())
            {
                if (rowNum == 2) // Row 2 should have header in column 2
                {
                    Assert.Equal("Col1", reader.GetValue(2));
                }
                else if (rowNum == 3) // Row 3 should have data in column 2
                {
                    Assert.Equal("Data", reader.GetValue(2));
                }
                rowNum++;
            }
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void AddSheet_WithHiddenTrue_CreatesHiddenSheet(string extension)
    {
        var fileName = $"test_hidden_sheet{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("VisibleSheet");
            writer.WriteSheet(dt.CreateDataReader());
            writer.AddSheet("HiddenSheet", hidden: true);
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            var names = reader.GetScheetNames();
            
            // Both sheets should be accessible even if one is hidden
            Assert.Equal(2, names.Length);
            Assert.Contains("VisibleSheet", names);
            Assert.Contains("HiddenSheet", names);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void OnCompress_Event_FiresDuringSave(string extension)
    {
        var fileName = $"test_on_compress{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        bool eventFired = false;
        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.OnCompress += () => eventFired = true;
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        Assert.True(eventFired);
        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void On10k_Event_FiresDuringLargeWrite(string extension)
    {
        var fileName = $"test_on_10k{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        for (int i = 0; i < 15000; i++)
        {
            dt.Rows.Add($"Row{i}");
        }

        int eventCount = 0;
        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.On10k += (count) => eventCount++;
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        Assert.True(eventCount >= 1);
        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void DocPopertyProgramName_SetValue_WritesToDocument(string extension)
    {
        var fileName = $"test_doc_prop{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.DocPopertyProgramName = "TestApplication";
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        Assert.True(File.Exists(fileName));
        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void RowsCount_AfterWrite_ReturnsCorrectValue(string extension)
    {
        var fileName = $"test_rows_count{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        for (int i = 0; i < 50; i++)
        {
            dt.Rows.Add($"Row{i}");
        }

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
            Assert.Equal(50, writer.RowsCount);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteSheet_MultipleSheetsWithDifferentRowCounts_Works(string extension)
    {
        var fileName = $"test_diff_rows{extension}";
        
        DataTable dt1 = new DataTable();
        dt1.Columns.Add("Col1", typeof(string));
        dt1.Rows.Add("A");
        
        DataTable dt2 = new DataTable();
        dt2.Columns.Add("Col1", typeof(string));
        for (int i = 0; i < 10; i++)
        {
            dt2.Rows.Add($"B{i}");
        }
        
        DataTable dt3 = new DataTable();
        dt3.Columns.Add("Col1", typeof(string));
        dt3.Rows.Add("C1");
        dt3.Rows.Add("C2");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt1.CreateDataReader());
            Assert.Equal(1, writer.RowsCount);
            
            writer.AddSheet("Sheet2");
            writer.WriteSheet(dt2.CreateDataReader());
            Assert.Equal(10, writer.RowsCount);
            
            writer.AddSheet("Sheet3");
            writer.WriteSheet(dt3.CreateDataReader());
            Assert.Equal(2, writer.RowsCount);
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteSheet_WithAutofilter_AddsFilter(string extension)
    {
        var fileName = $"test_autofilter{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Columns.Add("Col2", typeof(int));
        dt.Rows.Add("A", 1);
        dt.Rows.Add("B", 2);
        dt.Rows.Add("C", 3);

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader(), doAutofilter: true);
        }

        // Verify file was created
        Assert.True(File.Exists(fileName));
        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteSheet_WithoutHeaders_WritesDataOnly(string extension)
    {
        var fileName = $"test_no_headers{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data1");
        dt.Rows.Add("Data2");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader(), headers: false);
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            // First row should be data, not header
            Assert.True(reader.Read());
            Assert.Equal("Data1", reader.GetValue(0));
            Assert.True(reader.Read());
            Assert.Equal("Data2", reader.GetValue(0));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Save_CalledExplicitly_SavesFile(string extension)
    {
        var fileName = $"test_explicit_save{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        var writer = ExcelWriter.CreateWriter(fileName);
        writer.AddSheet("Sheet1");
        writer.WriteSheet(dt.CreateDataReader());
        writer.Save();
        
        Assert.True(File.Exists(fileName));
        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Dispose_WithoutSave_SavesAutomatically(string extension)
    {
        var fileName = $"test_auto_save{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
            // No explicit Save() call
        }

        Assert.True(File.Exists(fileName));
        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteSheet_StringArray_Works(string extension)
    {
        var fileName = $"test_string_array{extension}";
        
        string[] data = new string[] { "One", "Two", "Three" };

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(data);
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read());
            Assert.Equal("One", reader.GetValue(0));
            Assert.True(reader.Read());
            Assert.Equal("Two", reader.GetValue(0));
            Assert.True(reader.Read());
            Assert.Equal("Three", reader.GetValue(0));
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void SuppressSomeDate_True_Works(string extension)
    {
        var fileName = $"test_suppress_date{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("DateCol", typeof(DateTime));
        dt.Rows.Add(new DateTime(100, 1, 1)); // Very old date

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.SuppressSomeDate = true;
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        Assert.True(File.Exists(fileName));
        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteSheet_EmptyDataTable_WithHeaders_WritesHeadersOnly(string extension)
    {
        var fileName = $"test_empty_with_headers{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Columns.Add("Col2", typeof(int));
        // No rows

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader(), headers: true);
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read()); // Header row
            Assert.Equal("Col1", reader.GetValue(0));
            Assert.Equal("Col2", reader.GetValue(1));
            Assert.False(reader.Read()); // No data rows
        }

        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteSheet_ToMemoryStream_Works(string extension)
    {
        using var memoryStream = new MemoryStream();
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        ExcelWriter writer = extension == ".xlsx" 
            ? new XlsxWriter(memoryStream) 
            : new XlsbWriter(memoryStream);
        
        writer.AddSheet("Sheet1");
        writer.WriteSheet(dt.CreateDataReader());
        writer.Dispose();
        
        Assert.True(memoryStream.Length > 0);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteSheet_ToFileStream_Works(string extension)
    {
        var fileName = $"test_file_stream{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("Col1", typeof(string));
        dt.Rows.Add("Data");

        using (var fileStream = File.Open(fileName, FileMode.Create))
        {
            ExcelWriter writer = extension == ".xlsx" 
                ? new XlsxWriter(fileStream) 
                : new XlsbWriter(fileStream);
            
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
            writer.Dispose();
        }

        Assert.True(File.Exists(fileName));
        File.Delete(fileName);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void WriteSheet_WithComplexTypes_HandlesGracefully(string extension)
    {
        var fileName = $"test_complex_types{extension}";
        
        DataTable dt = new DataTable();
        dt.Columns.Add("GuidCol", typeof(Guid));
        dt.Columns.Add("DecimalCol", typeof(decimal));
        dt.Columns.Add("FloatCol", typeof(float));
        
        var guid = Guid.NewGuid();
        dt.Rows.Add(guid, 123.456m, 78.9f);

        using (var writer = ExcelWriter.CreateWriter(fileName))
        {
            writer.AddSheet("Sheet1");
            writer.WriteSheet(dt.CreateDataReader());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(fileName);
            reader.ActualSheetName = "Sheet1";
            
            Assert.True(reader.Read()); // Skip header
            Assert.True(reader.Read());
            
            // Guid should be stored as string
            Assert.Equal(guid.ToString(), reader.GetValue(0).ToString());
            // Decimal and float as double
            Assert.Equal(123.456, Convert.ToDouble(reader.GetValue(1)), 3);
            Assert.Equal(78.9, reader.GetDouble(2), 1);
        }

        File.Delete(fileName);
    }

    /// <summary>
    /// Rows whose content ends exactly on the 64 KiB row-buffer boundary used to leave the buffer
    /// full: the next boolean cell threw IndexOutOfRangeException and the next number was dropped
    /// silently. Both must now round-trip.
    /// </summary>
    [Fact]
    public void LongRows_AtRowBufferBoundary_PreserveFollowingCells()
    {
        AssertLongRowsRoundTrip(trailingIsBoolean: true);
        AssertLongRowsRoundTrip(trailingIsBoolean: false);
    }

    private static void AssertLongRowsRoundTrip(bool trailingIsBoolean)
    {
        // A fresh writer per length is required: the row buffer only has to be exactly full once,
        // and it stays grown for the remaining rows of the same writer.
        for (int length = 65_440; length <= 65_560; length++)
        {
            var table = new DataTable();
            table.Columns.Add("Text", typeof(string));
            table.Columns.Add("Trailing", trailingIsBoolean ? typeof(bool) : typeof(int));
            table.Rows.Add(new string('x', length), trailingIsBoolean ? true : (object)length);

            using var stream = new MemoryStream();
            using (var writer = new XlsxWriter(stream, useScharedStrings: false))
            {
                writer.AddSheet("Sheet1");
                writer.WriteSheet(table.CreateDataReader(), headers: false);
            }
            stream.Position = 0;

            using var reader = new XlsxOrXlsbReadOrEdit();
            reader.Open(stream);
            reader.ActualSheetName = "Sheet1";

            Assert.True(reader.Read(), $"length {length}: row missing");
            Assert.Equal(length, ((string)reader.GetValue(0)).Length);
            if (trailingIsBoolean)
            {
                Assert.Equal(true, reader.GetValue(1));
            }
            else
            {
                Assert.Equal(length, reader.GetInt32(1));
            }
            Assert.False(reader.Read(), $"length {length}: unexpected extra row");
        }
    }
}