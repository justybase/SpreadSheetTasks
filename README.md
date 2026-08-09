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
- Update existing XLSX sheets and pivot-table ranges.
- Targets `net8.0`, `net9.0` and `net10.0`.
- NativeAOT and trimming analyzers enabled for the library.

## Installation

```bash
dotnet add package SpreadSheetTasks --version 1.0.1
```

Or add the package reference manually:

```xml
<PackageReference Include="SpreadSheetTasks" Version="1.0.1" />
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

## License

SpreadSheetTasks is released under the [MIT License](LICENSE).
