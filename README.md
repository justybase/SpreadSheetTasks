# SpreadSheetTasks

[![CI](https://github.com/KrzysztofDusko/SpreadSheetTasks/actions/workflows/ci.yml/badge.svg)](https://github.com/KrzysztofDusko/SpreadSheetTasks/actions/workflows/ci.yml)
[![NuGet](https://img.shields.io/nuget/v/SpreadSheetTasks.svg)](https://www.nuget.org/packages/SpreadSheetTasks/)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

Fast, allocation-conscious .NET support for reading and writing Excel workbooks in
`.xlsx` and `.xlsb` formats.

SpreadSheetTasks is designed for data pipelines, backend services, command-line
tools and applications that need predictable streaming access to Excel data
without requiring Microsoft Excel.

## Highlights

- Read and write both XLSX and XLSB files.
- Stream rows with typed getters such as `GetString`, `GetInt32`, `GetDouble`
  and `GetDateTime`.
- Write from `DataTable`, `IDataReader`, object arrays or typed row lists.
- Support multiple sheets, autofilters, hidden sheets and cell formatting.
- Apply reusable cell styles (fills, font properties, number formats and XLSX alignment) by header or A1 range.
- Update existing `.xlsx` and `.xlsb` sheets while preserving the rest of the workbook.
- Update worksheet-backed pivot-cache ranges and mark dependent pivot tables for refresh.
- Targets `net8.0`, `net9.0` and `net10.0`.
- NativeAOT and trimming analyzers enabled for the library.

## Installation

```bash
dotnet add package SpreadSheetTasks --version 2.1.0
```

Or add the package reference manually:

```xml
<PackageReference Include="SpreadSheetTasks" Version="2.1.0" />
```

## Quick start

### Write an XLSX or XLSB file

```csharp
using SpreadSheetTasks;
using System.Data;

var table = new DataTable();
table.Columns.Add("Name", typeof(string));
table.Columns.Add("Age", typeof(int));
table.Rows.Add("Alice", 30);
table.Rows.Add("Bob", 25);

using var writer = new XlsxWriter("people.xlsx"); // or XlsbWriter
writer.AddSheet("People");
writer.WriteSheet(table.CreateDataReader());
```

### Compression and large writes

Both writers default to **zlib level 3**: the same balanced level used by the
JavaScript and Python implementations compared in `comparison/`. SST uses a
small built-in ZIP writer for that default, so writes are 13–21% faster than
the previous `CompressionLevel.Optimal` (zlib 6) path, at a modestly larger
file size. Explicit `CompressionLevel` values keep the standard .NET
`ZipArchive` path, and non-seekable streams fall back to it automatically
(where level 3 maps to `Optimal` on .NET 8, which has no numeric levels).

```csharp
using System.IO.Compression;
using SpreadSheetTasks;

// Default: zlib level 3 (seekable output)
using var balanced = new XlsxWriter("balanced.xlsx");

// Explicit levels keep the previous enum behaviour
using var fastest = new XlsxWriter("fastest.xlsx", _clvl: CompressionLevel.Fastest);
using var smallest = new XlsbWriter("smallest.xlsb", compressionLevel: CompressionLevel.SmallestSize);
```

XLSB can inline strings instead of maintaining a shared-string table. This is
useful when most text is unique or when the final shared-string pass matters:

```csharp
using var inline = new XlsbWriter("inline.xlsb", useSharedStrings: false);
```

### Read rows without loading the whole workbook

```csharp
using SpreadSheetTasks;

using var reader = new XlsxOrXlsbReadOrEdit();
reader.Open("people.xlsx");
reader.ActualSheetName = "People";

while (reader.Read())
{
    string name = reader.GetString(0);
    int age = reader.GetInt32(1);
    Console.WriteLine($"{name}: {age}");
}
```

On hot paths, `reader.CurrentRow` exposes the current row as a
`ReadOnlySpan<FieldInfo>` without boxing numbers, dates or booleans. The span
is valid until the next `Read()` call; see
[GetExcelDataType / GetNativeValue](docs/index.md#getexceldatatype--getnativevalue).

### Write typed rows

```csharp
using SpreadSheetTasks;

var headers = new List<string> { "Product", "Price", "Quantity" };
var types = new List<TypeCode> { TypeCode.String, TypeCode.Double, TypeCode.Int32 };
var rows = new List<object?[]>
{
    new object?[] { "Apple", 1.99, 100 },
    new object?[] { "Banana", 0.99, 250 }
};

using var writer = ExcelWriter.CreateWriter("products.xlsx");
writer.AddSheet("Products");
writer.WriteSheet(headers, types, rows, doAutofilter: true);
```

### Apply cell styles

`CellStyle` can style the first written row and arbitrary absolute A1 cells or
ranges. Styles are deduplicated in the workbook, and later overlapping ranges
override only the properties they specify.

```csharp
using SpreadSheetTasks;
using System.Data;

var table = new DataTable();
table.Columns.Add("Product", typeof(string));
table.Columns.Add("Quantity", typeof(int));
table.Columns.Add("Comment", typeof(string));
table.Rows.Add("Apples", 1200, "Restock soon");

using var writer = new XlsxWriter("report.xlsx"); // same styling API on XlsbWriter
writer.AddSheet("Data");
writer.WriteSheet(table.CreateDataReader(),
    headerStyle: new CellStyle { Fill = "008000", FontColor = "FFFFFF" },
    styles: new Dictionary<string, CellStyle>
    {
        ["B2"] = new CellStyle { Fill = "FFF2CC", NumberFormat = F.THOUSANDS_SEP },
        ["C2"] = new CellStyle { Fill = "FF0000", FontColor = "FFFFFF", Horizontal = "center", Wrap = true },
    });
```

Supported properties are `Fill`, `FontColor`, `Bold`, `Italic`, `Underline`,
`NumberFormat`, `Horizontal`, `Vertical` and `Wrap`. Colors are six-digit RGB
hex, with or without `#`. `null` properties leave existing values unchanged;
`FormattedCell.Format` takes precedence over a style number format. Alignment
and wrapping are available for XLSX; XLSB throws `NotSupportedException` for
those properties. A header style is merged with the default bold style (`Bold =
false` disables bold), and a range style that covers the first written row of a
headed sheet keeps that bold too. Pass `headerStyle: null` to apply only range
styles. Range coordinates refer to the worksheet, even when `startingRow` or
`startingColumn` is set.

### Update an existing XLSX or XLSB workbook

Use `XlsxUpdater` for `.xlsx` and `XlsbUpdater` for `.xlsb`. The updater keeps
the original package in memory, replaces only the selected worksheet data and
related pivot-cache metadata, and does not require Microsoft Excel at runtime:

```csharp
using SpreadSheetTasks;

var rows = new object?[][]
{
    ["Widget", 50000m],
    ["Gadget", 75000m]
};

using var updater = new XlsxUpdater("report.xlsx"); // XlsbUpdater for .xlsb
updater.ReplaceSheetData("Data", rows,
    new ReplaceSheetDataOptions { Headers = ["Product", "Revenue"] });
updater.Save("report-updated.xlsx");
```

`Save()` without an output path safely replaces the source file. `ToArray()`
can be used when the updated package must be handled as bytes. Only `.xlsx`
and `.xlsb` are supported by the updater API.

## NativeAOT and trimming

The package enables the .NET AOT and trimming analyzers and is tested with a
published NativeAOT application. The supported package target frameworks are
`net8.0`, `net9.0` and `net10.0`.

For an application, publish for a concrete runtime identifier:

```bash
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishAot=true
```

The repository contains a runnable AOT smoke test covering XLSX, XLSB,
`DataTable`, typed list rows and stream writing. See the complete
[performance and NativeAOT guide](docs/performance-and-aot.md).

## Documentation

- [Complete usage guide](docs/index.md)
- [Performance and NativeAOT](docs/performance-and-aot.md)
- [GitHub repository](https://github.com/KrzysztofDusko/SpreadSheetTasks)
- [NuGet package](https://www.nuget.org/packages/SpreadSheetTasks/)

## Performance

The repository includes BenchmarkDotNet projects for regular .NET 10 execution
and a focused comparison between .NET 10 JIT and .NET 10 NativeAOT. Benchmark
commands, environment details and measured results are maintained in
[docs/performance-and-aot.md](docs/performance-and-aot.md).

Benchmark results depend on CPU, storage, operating system, input files and
runtime version. Always reproduce measurements on the target environment before
making capacity or latency commitments.

## Compatibility notes

Version 1.0.0 introduced the following corrected names:

| Previous name | Current name |
|---|---|
| `GetScheetNames()` | `GetSheetNames()` |
| `DocPopertyProgramName` | `DocPropertyProgramName` |
| `SuppressSomeDate` | `SuppressYear1000Dates` |
| `overLimit` | `maxRows` |
| `UseMemoryStreamInXlsb` field | `UseMemoryStreamInXlsb` property |

The old members remain available with `[Obsolete]` warnings where applicable.

Version 2.1.0 changes the default ZIP compression to zlib level 3 and makes the
compression-level constructor parameters nullable (`null` = level 3). Pass
`CompressionLevel.Optimal` explicitly to keep the previous level-6 file sizes.
`XlsbWriter` also gained an optional `useSharedStrings` parameter, and
`ExcelReaderAbstract.CurrentRow` was added. Because the constructor signatures
changed, recompile against 2.1.0 rather than mixing binaries from 2.0.x.

## License

SpreadSheetTasks is released under the [MIT License](LICENSE).
