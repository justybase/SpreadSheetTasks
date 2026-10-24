# SpreadSheetTasks Usage Guide

## Table of Contents

1. [Installation](#installation)
2. [Writing Excel Files](#writing-excel-files)
   - [Write from DataTable](#write-from-datatable)
   - [Write from List](#write-from-list)
   - [Write from string array](#write-from-string-array)
   - [Write from object\[\]\[\]](#write-from-object)
   - [Write with per-cell formatting (FormattedCell\[\]\[\])](#write-with-per-cell-formatting-formattedcell)
   - [Cell styles](#cell-styles)
   - [Multi-sheet write](#multi-sheet-write)
   - [Write with Autofilter](#write-with-autofilter)
   - [Write with formatted cells](#write-with-formatted-cells)
   - [Write to Stream (MemoryStream / FileStream)](#write-to-stream)
   - [Write to Stream with custom buffer size](#write-to-stream-with-custom-buffer-size)
   - [Compression level and inline XLSB strings](#compression-level-and-inline-xlsb-strings)
   - [Hidden sheets](#hidden-sheets)
   - [Events: OnCompress / On10k](#events-oncompress--on10k)
   - [Document properties](#document-properties)
   - [SuppressYear1000Dates](#suppressyear1000dates)
3. [Reading Excel Files](#reading-excel-files)
   - [Basic read](#basic-read)
   - [Read from a stream](#read-from-a-stream)
   - [Read with typed getters](#read-with-typed-getters)
   - [GetRowsOfSheet](#getrowsofsheet)
   - [Get sheet names](#get-sheet-names)
   - [RowCount and ResultsCount](#rowcount-and-resultscount)
   - [TreatAllColumnsAsText](#treatallcolumnsastext)
   - [UseMemoryStreamInXlsb](#usememorystreaminxlsb)
   - [Read in update mode (legacy XLSX API)](#read-in-update-mode-legacy-xlsx-api)
   - [GetExcelDataType / GetNativeValue](#getexceldatatype--getnativevalue)
4. [Factory method](#factory-method-excelwritercreatewriter)
5. [Update existing XLSX/XLSB workbooks](#update-existing-xlsxxlsb-workbooks)
6. [Low-level existing XLSX writing](#low-level-existing-xlsx-writing)
   - [XlsxWriter.WriteToExisting](#xlsxwriterwritetoexisting)
   - [ReplaceSheetData + ReplacePivotTableDim](#replacesheetdata--replacepivottabledim)
7. [Breaking Changes in v1.0.0](#breaking-changes-in-v100)
8. [Format constants (F class)](#format-constants-f-class)
9. [Performance and NativeAOT](performance-and-aot.md)

---

## Installation

```bash
dotnet add package SpreadSheetTasks
```

```xml
<PackageReference Include="SpreadSheetTasks" Version="2.1.0" />
```

## Performance and NativeAOT

See the dedicated [performance and NativeAOT guide](performance-and-aot.md)
for the support matrix, publish smoke test and .NET 10 versus NativeAOT
benchmark commands and results.

---

## Writing Excel Files

### Write from DataTable

```csharp
using SpreadSheetTasks;
using System.Data;

var dt = new DataTable();
dt.Columns.Add("Name", typeof(string));
dt.Columns.Add("Age", typeof(int));
dt.Columns.Add("Salary", typeof(double));
dt.Rows.Add("Alice", 30, 5000.0);
dt.Rows.Add("Bob", 25, 4500.0);

using (var writer = new XlsxWriter("employees.xlsx"))
{
    writer.AddSheet("Employees");
    writer.WriteSheet(dt.CreateDataReader());
}

using (var writer = new XlsbWriter("employees.xlsb"))
{
    writer.AddSheet("Employees");
    writer.WriteSheet(dt.CreateDataReader());
}
```

### Write from List

```csharp
using SpreadSheetTasks;

var headers = new List<string> { "Product", "Price", "Quantity" };
var types = new List<TypeCode> { TypeCode.String, TypeCode.Double, TypeCode.Int32 };
var rows = new List<object?[]>
{
    new object?[] { "Apple", 1.99, 100 },
    new object?[] { "Banana", 0.99, 250 },
    new object?[] { "Cherry", 3.49, 75 },
};

using (var writer = ExcelWriter.CreateWriter("products.xlsx"))
{
    writer.AddSheet("Products");
    writer.WriteSheet(headers, types, rows, headers: true, doAutofilter: true);
}
```

### Write from string array

```csharp
using SpreadSheetTasks;

string[] data = ["Alpha", "Beta", "Gamma"];

using (var writer = new XlsxWriter("strings.xlsx"))
{
    writer.AddSheet("Sheet1");
    writer.WriteSheet(data);
}
```

### Write from object\[\]\[\]

```csharp
using SpreadSheetTasks;

var rows = new object[][]
{
    new object[] { "Apple", 1.99, 100 },
    new object[] { "Banana", 0.99, 250 },
    new object[] { "Cherry", 3.49, 75 },
};

using (var writer = new XlsxWriter("products.xlsx"))
{
    writer.AddSheet("Products");
    writer.WriteSheet(rows, headers: new[] { "Product", "Price", "Quantity" }, headers_row: true, doAutofilter: true);
}
```

### Write with per-cell formatting (FormattedCell\[\]\[\])

```csharp
using SpreadSheetTasks;

var rows = new FormattedCell?[][]
{
    new FormattedCell?[] { new(1234567, F.THOUSANDS_SEP), new(1234.56, F.CURRENCY_PLN) },
    new FormattedCell?[] { new(0.25, F.PERCENTAGE), new(12345.67, F.SCIENTIFIC) },
};

using (var writer = new XlsxWriter("formatted.xlsx"))
{
    writer.AddSheet("Sheet1");
    writer.WriteSheet(rows, headers: new[] { "Count", "Value" });
}
```

### Cell styles

Use `CellStyle` with `headerStyle` to format the first written row, and `styles`
to style arbitrary absolute A1 cells or rectangular ranges. Both `XlsxWriter`
and `XlsbWriter` support solid fill and font colors, bold, italic, underline
and number formats. Alignment and text wrapping are supported by XLSX; XLSB
throws `NotSupportedException` if those properties are requested.

```csharp
using SpreadSheetTasks;
using System.Data;

var table = new DataTable();
table.Columns.Add("Product", typeof(string));
table.Columns.Add("Quantity", typeof(int));
table.Columns.Add("Comment", typeof(string));
table.Rows.Add("Apples", 1200, "Restock soon");
table.Rows.Add("Pears", 8, "Fresh");

var header = new CellStyle { Fill = "#008000", FontColor = "#FFFFFF" };
var warning = new CellStyle { Fill = "FFF2CC", NumberFormat = F.THOUSANDS_SEP };
var note = new CellStyle { Fill = "FF0000", FontColor = "FFFFFF", Horizontal = "center", Wrap = true };

using var writer = new XlsxWriter("styled.xlsx"); // CellStyle also works with XlsbWriter
writer.AddSheet("Data");
writer.WriteSheet(table.CreateDataReader(),
    headerStyle: header,
    styles: new Dictionary<string, CellStyle>
    {
        ["B2:B3"] = warning,
        ["C2"] = note,
    });
```

`CellStyle` properties are `Fill`, `FontColor`, `Bold`, `Italic`, `Underline`,
`NumberFormat`, `Horizontal`, `Vertical` and `Wrap`. Colors are six-digit RGB
hex strings, with or without `#`. A `null` property means "leave unchanged";
overlapping ranges therefore compose, and later ranges override only the
properties they specify. A cell's `FormattedCell.Format` takes precedence over
`CellStyle.NumberFormat`. A header style is merged with the default bold header
style (set `Bold = false` to disable bold), and a range style that covers the
first written row of a headed sheet keeps that bold too. Pass `headerStyle:
null` when only range styles are needed. Range keys use absolute worksheet
coordinates, including when `startingRow` or `startingColumn` is set; `$A$1`
references are also accepted.

The same `headerStyle` and `styles` arguments are available on `WriteSheet`
overloads for `DataTable`, `IDataReader`, typed lists, `object[][]` and
`FormattedCell?[][]`.

### Multi-sheet write

```csharp
using SpreadSheetTasks;
using System.Data;

var dt = new DataTable();
dt.Columns.Add("Value", typeof(int));
dt.Rows.Add(1);
dt.Rows.Add(2);

using (var writer = ExcelWriter.CreateWriter("multi.xlsx"))
{
    writer.AddSheet("Sheet1");
    writer.WriteSheet(dt.CreateDataReader());

    writer.AddSheet("Sheet2");
    writer.WriteSheet(dt.CreateDataReader());

    writer.AddSheet("HiddenSheet", hidden: true);
    writer.WriteSheet(dt.CreateDataReader());
}
```

### Write with Autofilter

```csharp
using SpreadSheetTasks;
using System.Data;

var dt = new DataTable();
dt.Columns.Add("City", typeof(string));
dt.Columns.Add("Population", typeof(int));
dt.Rows.Add("New York", 8_400_000);
dt.Rows.Add("Los Angeles", 3_800_000);
dt.Rows.Add("Chicago", 2_700_000);

using (var writer = ExcelWriter.CreateWriter("cities.xlsx"))
{
    writer.AddSheet("Cities");
    writer.WriteSheet(dt.CreateDataReader(), doAutofilter: true);
}
```

### Write with formatted cells

```csharp
using SpreadSheetTasks;
using System.Data;

var dt = new DataTable();
dt.Columns.Add("Description", typeof(string));
dt.Columns.Add("Value", typeof(object));

dt.Rows.Add("Thousands", new FormattedCell(1234567, F.THOUSANDS_SEP));
dt.Rows.Add("Currency PLN", new FormattedCell(1234.56, F.CURRENCY_PLN));
dt.Rows.Add("Date", new FormattedCell(new DateTime(2026, 6, 1), F.DATE_SHORT));
dt.Rows.Add("DateTime ISO", new FormattedCell(new DateTime(2026, 6, 1, 14, 30, 0), F.DATETIME_ISO));
dt.Rows.Add("Percentage", new FormattedCell(0.25, F.PERCENTAGE));
dt.Rows.Add("Scientific", new FormattedCell(12345.67, F.SCIENTIFIC));
dt.Rows.Add("Time", new FormattedCell(new DateTime(2026, 6, 1, 8, 15, 0), F.TIME_HH_MM_SS));

using (var writer = ExcelWriter.CreateWriter("formatted.xlsx"))
{
    writer.AddSheet("Formatted");
    writer.WriteSheet(dt.CreateDataReader());
}
```

You can also use custom format strings directly:

```csharp
dt.Rows.Add("Custom", new FormattedCell(1234.56, "#,##0.00 \"USD\""));
dt.Rows.Add("ZIP", new FormattedCell(12345, "00000"));
```

### Write to Stream

```csharp
using SpreadSheetTasks;
using System.Data;

// FileStream
using (var fs = File.Open("stream_output.xlsx", FileMode.Create))
using (var writer = new XlsxWriter(fs))
{
    writer.AddSheet("Sheet1");
    var dt = new DataTable();
    dt.Columns.Add("Col1", typeof(string));
    dt.Rows.Add("Hello");
    writer.WriteSheet(dt.CreateDataReader());
}

// MemoryStream
byte[] excelBytes;
using (var ms = new MemoryStream())
using (var writer = new XlsxWriter(ms))
{
    writer.AddSheet("Sheet1");
    var dt = new DataTable();
    dt.Columns.Add("Col1", typeof(string));
    dt.Rows.Add("World");
    writer.WriteSheet(dt.CreateDataReader());

    // leaveExcelArchiveOpen: true (default for stream)
    // data is available after Dispose
}
```

### Write to Stream with custom buffer size

```csharp
using SpreadSheetTasks;
using System.Data;

// Larger buffer improves write throughput on large files
using (var writer = new XlsxWriter("large.xlsx", bufferSize: 65536))
{
    writer.AddSheet("Data");
    var dt = new DataTable();
    dt.Columns.Add("Index", typeof(int));
    for (int i = 0; i < 100_000; i++) dt.Rows.Add(i);
    writer.WriteSheet(dt.CreateDataReader());
}
```

Both `XlsxWriter` and `XlsbWriter` accept `bufferSize` in their constructors (default: 4096).

### Compression level and inline XLSB strings

The default ZIP compression is **zlib level 3**, which balances write speed and
file size (the same default as the JavaScript and Python libraries in
`comparison/`). Seekable outputs use the library's own ZIP writer for that
level; explicit `CompressionLevel` values and non-seekable streams use
`System.IO.Compression.ZipArchive`:

```csharp
using System.IO.Compression;
using SpreadSheetTasks;

// Default: zlib level 3
using (var writer = new XlsxWriter("default.xlsx")) { /* ... */ }

// Explicit enum level: previous ZipArchive behaviour
using (var writer = new XlsxWriter("fast.xlsx", _clvl: CompressionLevel.Fastest)) { /* ... */ }

// XLSB: write strings inline instead of building a shared-string table
using (var writer = new XlsbWriter("inline.xlsb", useSharedStrings: false)) { /* ... */ }
```

`useSharedStrings: false` avoids the per-workbook string dictionary and the
final shared-string pass. It is usually the best choice when most text values
are unique; repeated text produces a smaller file with the default
`useSharedStrings: true`. On .NET 8 and on non-seekable streams, level 3 maps
to the closest `CompressionLevel` enum value (typically `Optimal`).

### Hidden sheets

```csharp
using SpreadSheetTasks;
using System.Data;

var dt = new DataTable();
dt.Columns.Add("Data", typeof(string));
dt.Rows.Add("visible");
dt.Rows.Add("hidden data");

using (var writer = ExcelWriter.CreateWriter("hidden.xlsx"))
{
    writer.AddSheet("Visible");
    writer.WriteSheet(dt.CreateDataReader());

    writer.AddSheet("Hidden", hidden: true);
    writer.WriteSheet(dt.CreateDataReader());
}
```

### Events: OnCompress / On10k

```csharp
using SpreadSheetTasks;
using System.Data;

var dt = new DataTable();
dt.Columns.Add("Index", typeof(int));
for (int i = 0; i < 25_000; i++)
    dt.Rows.Add(i);

using (var writer = new XlsxWriter("events_demo.xlsx"))
{
    writer.OnCompress += () => Console.WriteLine("Starting compression...");
    writer.On10k += (row) => Console.WriteLine($"Written {row} rows...");

    writer.AddSheet("LargeData");
    writer.WriteSheet(dt.CreateDataReader());
}
```

### Document properties

```csharp
using SpreadSheetTasks;
using System.Data;

var dt = new DataTable();
dt.Columns.Add("Col1", typeof(string));
dt.Rows.Add("Data");

using (var writer = ExcelWriter.CreateWriter("docprop.xlsx"))
{
    writer.DocPropertyProgramName = "MyApplication";
    writer.AddSheet("Sheet1");
    writer.WriteSheet(dt.CreateDataReader());
}
```

> **Note:** The old name `DocPopertyProgramName` is still available with an `[Obsolete]` warning.

### SuppressYear1000Dates

Suppresses DateTime values where the year is 1000 (Excel serial date 0):

```csharp
using SpreadSheetTasks;
using System.Data;

var dt = new DataTable();
dt.Columns.Add("Date", typeof(DateTime));
dt.Rows.Add(new DateTime(1000, 1, 1)); // will be suppressed
dt.Rows.Add(new DateTime(2026, 6, 1)); // normal

using (var writer = ExcelWriter.CreateWriter("suppress.xlsx"))
{
    writer.SuppressYear1000Dates = true;
    writer.AddSheet("Sheet1");
    writer.WriteSheet(dt.CreateDataReader());
}
```

> **Note:** The old name `SuppressSomeDate` is still available with an `[Obsolete]` warning.

---

## Reading Excel Files

### Basic read

```csharp
using SpreadSheetTasks;

using (var reader = new XlsxOrXlsbReadOrEdit())
{
    reader.Open("file.xlsx"); // auto-detects xlsx/xlsb by extension
    reader.ActualSheetName = "Sheet1";

    object[]? row = null;
    while (reader.Read())
    {
        row ??= new object[reader.FieldCount];
        reader.GetValues(row);
        // process row...
    }
}
```

`Open` uses read-only, shared access (`FileAccess.Read`, `FileShare.Read`), so
read-only files can be opened and several readers in the same process can read
the same workbook concurrently. Pass `updateMode: true` only when the file must
be modified; that mode requires write access and exclusive sharing.

### Read from a stream

`Open(Stream)` accepts any seekable stream and detects XLSX/XLSB from the
package entries. The reader takes ownership of the stream and disposes it with
the reader. This is useful for in-memory data, cloud downloads or benchmarks
that should not measure disk I/O:

```csharp
using SpreadSheetTasks;

byte[] bytes = File.ReadAllBytes("data.xlsx");
using var input = new MemoryStream(bytes, writable: false);
using var reader = new XlsxOrXlsbReadOrEdit();
reader.Open(input); // format detected from xl/workbook.xml or xl/workbook.bin
reader.ActualSheetName = "Sheet1";

while (reader.Read())
{
    string name = reader.GetString(0);
}
```

### Read with typed getters

```csharp
using SpreadSheetTasks;

using (var reader = new XlsxOrXlsbReadOrEdit())
{
    reader.Open("data.xlsx");
    reader.ActualSheetName = "Sheet1";

    while (reader.Read())
    {
        string name = reader.GetString(0);
        int age = reader.GetInt32(1);
        double salary = reader.GetDouble(2);
        DateTime? hireDate = null;

        // GetDateTime throws InvalidCastException if cell is not a DateTime
        try { hireDate = reader.GetDateTime(3); }
        catch (InvalidCastException) { }
    }
}
```

### GetRowsOfSheet

```csharp
using SpreadSheetTasks;

using (var reader = new XlsxOrXlsbReadOrEdit())
{
    reader.Open("data.xlsx");
    int totalRows = reader.GetRowsOfSheet("Sheet1").Count();
}
```

`GetRowsOfSheet` (and `GetRowsOfXlsx` / `GetRowsOfXlsb`) returns a new `object[]` for every row, so rows can be stored or materialized safely.

For large sheets where each row is consumed before the next one is read, use `GetRowsOfSheetReadOnly`. It reuses a single array for all rows; a row is valid only until the next iteration, so copy the values you need inside the loop.

```csharp
foreach (var row in reader.GetRowsOfSheetReadOnly("Sheet1"))
{
    // Consume row[] here; do not keep a reference to it.
}
```

> **Breaking change:** earlier versions returned a shared buffer from `GetRowsOfSheet`. Code that relied on that buffer for speed should switch to `GetRowsOfSheetReadOnly`.

### Get sheet names

```csharp
using SpreadSheetTasks;

using (var reader = new XlsxOrXlsbReadOrEdit())
{
    reader.Open("multi.xlsx");
    string[] sheetNames = reader.GetSheetNames();

    foreach (string name in sheetNames)
    {
        reader.ActualSheetName = name;
        Console.WriteLine($"Sheet: {name}");
    }
}
```

> **Note:** The old name `GetScheetNames()` is still available with an `[Obsolete]` warning.

### RowCount and ResultsCount

```csharp
using SpreadSheetTasks;

using (var reader = new XlsxOrXlsbReadOrEdit())
{
    reader.Open("data.xlsx");
    reader.ActualSheetName = "Sheet1";

    int rowCount = reader.RowCount;        // estimated row count
    int sheetCount = reader.ResultsCount;  // number of sheets
}
```

### TreatAllColumnsAsText

```csharp
using SpreadSheetTasks;

using (var reader = new XlsxOrXlsbReadOrEdit())
{
    reader.Open("data.xlsx");
    reader.TreatAllColumnsAsText = true; // all values returned as strings
    reader.ActualSheetName = "Sheet1";

    while (reader.Read())
    {
        string val = reader.GetString(0);
    }
}
```

### UseMemoryStreamInXlsb

```csharp
using SpreadSheetTasks;

using (var reader = new XlsxOrXlsbReadOrEdit())
{
    reader.UseMemoryStreamInXlsb = false; // lower RAM usage, slightly slower
    reader.Open("data.xlsb");
    reader.ActualSheetName = "Sheet1";

    while (reader.Read()) { /* ... */ }
}
```

### Read in update mode (legacy XLSX API)

```csharp
using SpreadSheetTasks;
using System.Data;

// Open existing file in update mode and replace sheet data
using (var reader = new XlsxOrXlsbReadOrEdit())
{
    reader.Open("existing.xlsx", updateMode: true);

    var newData = new DataTable();
    newData.Columns.Add("Name", typeof(string));
    newData.Columns.Add("Score", typeof(int));
    newData.Rows.Add("Alice", 95);
    newData.Rows.Add("Bob", 87);

    string range = reader.ReplaceSheetData("Sheet1", newData.CreateDataReader());
    Console.WriteLine($"Replaced range: {range}");
}
```

For replacement of existing workbooks, prefer `XlsxUpdater`/`XlsbUpdater`
below. They preserve worksheet metadata and support both formats.

### GetExcelDataType / GetNativeValue

For advanced scenarios where you need raw cell type information or minimal overhead:

```csharp
using SpreadSheetTasks;

using (var reader = new XlsxOrXlsbReadOrEdit())
{
    reader.Open("data.xlsx");
    reader.ActualSheetName = "Sheet1";

    while (reader.Read())
    {
        // Recommended: use typed getters
        string name = reader.GetString(0);

        // Inspect cell type at column 1
        ExcelDataType type = reader.GetExcelDataType(1);

        // High‑performance escape hatch (check type first!)
        if (type == ExcelDataType.Double)
        {
            ref FieldInfo field = ref reader.GetNativeValue(1);
            double val = field.doubleValue;
        }

        // Or get the entire row as a ref to the internal buffer
        ref FieldInfo[] row = ref reader.GetNativeValues();

        // Non-boxing span over the current row (valid until the next Read())
        ReadOnlySpan<FieldInfo> current = reader.CurrentRow;
    }
}
```

`CurrentRow` is the preferred low-allocation API: it returns a
`ReadOnlySpan<FieldInfo>` over the same per-instance buffer, with no
`object[]` allocation and no boxing. Check `FieldInfo.type` before reading a
union field.

| `ExcelDataType` | Meaning | `FieldInfo` union field |
|---|---|---|
| `Null` | Empty cell | — |
| `Int32` | 32‑bit integer | `int32Value` |
| `Int64` | 64‑bit integer | `int64Value` |
| `Double` | Floating‑point number | `doubleValue` |
| `DateTime` | Date/time | `dtValue` |
| `Boolean` | True/false | `boolValue` |
| `String` | Text | `strValue` |

> **`GetNativeValue` / `GetNativeValues` are marked `[Obsolete]` and `[EditorBrowsable(Never)]`.** Use typed getters (`GetString`, `GetInt32`, etc.) for general code. Only use `GetNativeValue` in hot paths where every allocation counts. Always call `GetExcelDataType()` first to check the discriminator.

---

## Factory method (ExcelWriter.CreateWriter)

```csharp
using SpreadSheetTasks;
using System.Data;

// Automatically selects XlsxWriter or XlsbWriter based on file extension
using (var writer = ExcelWriter.CreateWriter("data.xlsx"))
{
    writer.AddSheet("Sheet1");
    var dt = new DataTable();
    dt.Columns.Add("Col1", typeof(string));
    dt.Rows.Add("Hello");
    writer.WriteSheet(dt.CreateDataReader());
}
```

---

## Update existing XLSX/XLSB workbooks

`XlsxUpdater` and `XlsbUpdater` replace the contents of one existing worksheet
without rebuilding the workbook. Other worksheets and package parts remain
untouched. The API is intentionally limited to `.xlsx` and `.xlsb`, which are
the two formats supported by this library.

```csharp
using SpreadSheetTasks;
using System.Data;

var table = new DataTable();
table.Columns.Add("Category", typeof(string));
table.Columns.Add("Amount", typeof(decimal));
table.Rows.Add("A", 100m);
table.Rows.Add("B", 200m);

using var updater = new XlsxUpdater("report.xlsx");
updater.ReplaceSheetData("Data", table,
    new ReplaceSheetDataOptions { Headers = ["Category", "Amount"] });
updater.Save("report-updated.xlsx");
```

For a binary workbook use the same calls with `XlsbUpdater` and `.xlsb` paths:

```csharp
using var updater = new XlsbUpdater("report.xlsb");
updater.ReplaceSheetData("Data", new object?[][]
{
    ["A", 100],
    ["B", 200]
}, new ReplaceSheetDataOptions { Headers = ["Category", "Amount"] });
updater.Save("report-updated.xlsb");
```

The overloads accept `IDataReader`, `DataTable`, `object?[][]` and
`IReadOnlyList<object?[]>`. `Headers` is optional; when supplied it becomes
row 1 and the data starts at row 2. Without it, the supplied rows start at
row 1. Trailing completely empty rows are omitted, while empty rows in the
middle are retained. Existing column styles are reused where possible;
`ReplaceSheetDataStyleFallback.General` selects the workbook's General style.

When a pivot cache uses the replaced worksheet as its source, the updater also
changes the source range and record count and marks the cache and dependent
pivot table for refresh on opening in Excel. It does not require Excel or COM
for normal operation. To validate generated files on Windows with Microsoft
Excel installed, run:

```powershell
.\tools\Validate-ExcelCom.ps1 -Path .\report-updated.xlsx -RequireExcel
```

`Save()` with no path overwrites the source safely; `ToArray()` returns the
complete updated package for callers that need byte-oriented output.

---

## Low-level existing XLSX writing

### XlsxWriter.WriteToExisting

Write data directly into an existing sheet's XML stream inside a `.xlsx` Zip archive:

```csharp
using SpreadSheetTasks;
using System.Data;
using System.IO;
using System.IO.Compression;

using (var archive = ZipFile.Open("existing.xlsx", ZipArchiveMode.Update))
{
    var entry = archive.GetEntry("xl/worksheets/sheet1.xml");
    using var sheetWriter = new StreamWriter(entry.Open());

    var dt = new DataTable();
    dt.Columns.Add("Name", typeof(string));
    dt.Columns.Add("Value", typeof(int));
    dt.Rows.Add("Alice", 100);

    XlsxWriter.WriteToExisting(sheetWriter, dt.CreateDataReader());
}
```

### ReplaceSheetData + ReplacePivotTableDim

For the built‑in update mode (handles full sheet replacement and pivot table references):

```csharp
using SpreadSheetTasks;
using System.Data;

using (var reader = new XlsxOrXlsbReadOrEdit())
{
    reader.Open("report.xlsx", updateMode: true);

    var newData = new DataTable();
    newData.Columns.Add("Product", typeof(string));
    newData.Columns.Add("Revenue", typeof(decimal));
    newData.Rows.Add("Widget", 50000m);
    newData.Rows.Add("Gadget", 75000m);

    // Replace sheet data and get the new cell range
    string range = reader.ReplaceSheetData("Sheet1", newData.CreateDataReader());

    // Point the pivot table to the new range
    reader.ReplacePivotTableDim("PivotTable1", range, doRefreshOnLoad: true);
}
```

---

## Breaking Changes in v1.0.0

When upgrading from v0.6.x to v1.0.0, the following renames may require updates:

| Old (v0.6.1) | New (v1.0.0) | Notes |
|---|---|---|
| `GetScheetNames()` | `GetSheetNames()` | Typo fix |
| `DocPopertyProgramName` | `DocPropertyProgramName` | Typo fix |
| `SuppressSomeDate` | `SuppressYear1000Dates` | Clarifies intent |
| `overLimit` parameter | `maxRows` parameter | Affects `WriteSheet` calls |
| `RowCount` returns `123123123` for unknown | `RowCount` returns `-1` for unknown | Sentinel value fix |
| `FormattingStreamWriter` (public) | `FormattingStreamWriter` (internal) | Was never intended for public use |
| `UseMemoryStreamInXlsb` (field) | `UseMemoryStreamInXlsb` (property) | Follows .NET conventions |
| Duplicate `F.SHORT_DATE` etc. | Marked `[Obsolete]` | Use canonical name |

All old names remain available with `[Obsolete]` warnings — code compiles but produces warnings.

---

## Format constants (F class)

| Constant | Format String | Description |
|----------|---------------|-------------|
| `F.THOUSANDS_SEP` | `#,##0` | Number with thousands separator |
| `F.CURRENCY_PLN` | `#,##0.00 "zł"` | Polish currency |
| `F.CURRENCY_EUR` | `#,##0.00 €` | Euro currency |
| `F.PERCENTAGE` | `0%` | Percentage |
| `F.SCIENTIFIC` | `0.00E+00` | Scientific notation |
| `F.TWO_DECIMALS` | `#,##0.00` | Number with 2 decimals |
| `F.TEXT` | `@` | Text format |
| `F.LEADING_ZEROS` | `000000000` | Leading zeros (9 digits) |
| `F.DATE_SHORT` | `dd.mm.yyyy` | Short date |
| `F.DATE_LONG` | `d mmmm yyyy` | Long date |
| `F.DATE_ISO` | `yyyy-mm-dd` | ISO date |
| `F.DATETIME_SHORT` | `dd.mm.yyyy hh:mm` | Short date/time |
| `F.DATETIME_LONG` | `d mmmm yyyy hh:mm:ss` | Long date/time |
| `F.DATETIME_ISO` | `yyyy-mm-dd"T"hh:mm:ss` | ISO date/time |
| `F.TIME_HH_MM` | `hh:mm` | Time (hours:minutes) |
| `F.TIME_HH_MM_SS` | `hh:mm:ss` | Time (hours:minutes:seconds) |
| `F.TIME_12H` | `h:mm AM/PM` | 12-hour time |
| `F.TIME_MS` | `hh:mm:ss.000` | Time with milliseconds |

> **Duplicate constants:** `F.SHORT_DATE`, `F.LONG_DATE`, `F.SHORT_DATE_TIME`, `F.LONG_DATE_TIME`, `F.ISO_DATE`, `F.ISO_DATE_TIME` are `[Obsolete]` — use `F.DATE_SHORT`, `F.DATE_LONG`, `F.DATETIME_SHORT`, `F.DATETIME_LONG`, `F.DATE_ISO`, `F.DATETIME_ISO` instead.
