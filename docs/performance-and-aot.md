# Performance and NativeAOT

This page documents the reproducible performance and NativeAOT checks for
SpreadSheetTasks.

## Support matrix

The library package targets:

| Target framework | Library build | AOT analysis |
|---|---:|---:|
| `net8.0` | Yes | Yes |
| `net9.0` | Yes | Yes |
| `net10.0` | Yes | Yes |
| `net11.0` | Yes | Yes |

The package enables `IsAotCompatible`, trimming analysis, AOT analysis and
single-file analysis. The source does not require runtime reflection for its
supported read/write paths.

NativeAOT is a property of the consuming application. Publish for a concrete
runtime identifier:

```bash
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishAot=true
```

NativeAOT applications are platform-specific. Replace `win-x64` with the RID
required by the deployment target.

## Repository verification

The `source/AotSmoke` executable is intentionally separate from the xUnit test
project. It is published as a real NativeAOT application and exercises:

- XLSX and XLSB writing and reading,
- the `DataTable` writer path,
- the `List<string> + List<TypeCode> + List<object?[]>` writer path,
- typed getters,
- writing to an existing stream.

Run it locally with:

```powershell
dotnet publish source/AotSmoke/AotSmoke.csproj `
  -c Release -r win-x64 --self-contained true `
  -p:PublishAot=true -o artifacts/aot-smoke

& .\artifacts\aot-smoke\AotSmoke.exe
```

The CI workflow performs the same publish-and-run check before a package is
published.

## Benchmark projects

The repository has two benchmark purposes:

1. `source/Benchmark` keeps the broader .NET 11 comparison suite, including
   the existing Sylvan comparison.
2. `source/BenchmarkAot` is deliberately limited to SpreadSheetTasks APIs so
   that the NativeAOT job is not coupled to the AOT status of comparison
   libraries.

Run the focused .NET 11 versus NativeAOT benchmark:

```powershell
dotnet run --project source/BenchmarkAot/BenchmarkAot.csproj `
  -c Release -- --filter "*NativeAotReadWriteBenchmarks*"
```

The benchmark class contains both jobs:

- `RuntimeMoniker.Net11_0` — regular .NET 11 JIT,
- `RuntimeMoniker.NativeAot11_0` — .NET 11 NativeAOT.

The NativeAOT job is built by BenchmarkDotNet, so the first run can take
significantly longer than subsequent runs. BenchmarkDotNet documents the
NativeAOT toolchain and runtime-specific benchmark jobs in its
[toolchain guide](https://benchmarkdotnet.org/articles/configs/toolchains.html).

## Fast XLSX worksheet reader

XLSX worksheet parts are parsed by a byte-level row scanner by default. It handles
the SpreadsheetML subset produced by Excel and by this library: shared and inline
strings, formulas (cached values), booleans, errors, number formats, dates, empty
cells and sparse rows. Anything it is not certain about is handed to `XmlReader`
for that row only: comments inside a row, rich-text inline runs, unknown elements
and prefixed tags. XML entities and character references are decoded directly by
the scanner; an unknown or malformed reference falls back for that row. After 32
consecutive fallback rows, or once at least 512 rows were parsed and a quarter of
them needed the fallback, the remainder of the part switches to `XmlReader` while
preserving the rows already emitted. A part that is not UTF-8 (for example a
UTF-16 BOM or a non-UTF-8 encoding declaration) or that uses prefixed element
names is handed to `XmlReader` entirely. A construct at row level that the
scanner cannot skip (a CDATA section or stray text inside `sheetData`, for
example) also switches the remainder of that part to `XmlReader`.

Shared strings are read by a byte scanner as well. It handles plain and rich-text
items, concatenates rich runs, excludes phonetic-guide text (`<rPh>`), honors
`xml:space`, decodes entities and `_xHHHH_` escapes, and uses `uniqueCount` only
as a capped capacity hint. Markup it does not reproduce exactly (for example a
CDATA section inside `<si>`) falls back to the streaming `XmlReader` reader for the
whole part.

Inline strings follow the same `xml:space` rules as `XmlReader`: a whitespace-only
`<t>` is only significant when `xml:space="preserve"` is in scope.

The rest of the sheet keeps streaming through a small adapter that serves the
buffered read-ahead bytes and then hands reads through to the still-open deflate
stream, so the unread remainder is never materialised. Rows that take the
fallback are parsed from the buffered bytes with the same cell code as the regular
path, which keeps types and values identical.

Worksheet UTF-8 is validated incrementally as buffers are filled, so damaged files
fail instead of being replacement-decoded. A single row that grows past 64 MB
without closing throws `InvalidDataException`; more than 16,384 columns, or row
numbers outside 1–1,048,576, are rejected by the scanner and handled by
`XmlReader`, which applies the XML rules and raises the file's own error.
Workbooks using the 1904 date system are supported in both XLSX (`workbookPr
date1904`) and XLSB (`BrtWbProp` bit 0).

Measured with the `--sst` short BenchmarkDotNet job (.NET 11.0.0, one warmup and
three iterations, AMD Ryzen 7 7840HS):

| Benchmark | Mean | Allocated |
|---|---:|---:|
| `Open 200K with shared strings` (byte scanner) | 37.17 ms | 20.5 MB |
| `Open 200K without shared strings` | 0.050 ms | 0.14 MB |
| `Read 200K - GetValues` | 137.48 ms | 35.0 MB |
| `Read 65K - GetValues` | 87.46 ms | 14.4 MB |
| `Read 65K - CurrentRow` | 80.18 ms | 0.60 MB |
| `Read 20K entity inline strings` | 6.89 ms | 8.35 MB |

The `200kFile.xlsx` shared-strings part is 10.5 MB and holds 200,004 unique
strings. The byte scanner reads it in roughly 27–34 ms standalone; the rest of the
`Open` time is ZIP entry decompression plus the workbook and styles parts. With
`_disableFastSharedStringsReader = true` the same open takes about 100 ms and
allocates about 91 MB, so the byte scanner is roughly 2.2–2.4× faster and
allocates about a quarter as much.

`CurrentRow` exposes the row as `ReadOnlySpan<FieldInfo>` without boxing. On the
65K fixture that drops the per-read allocation from 14.4 MB (`GetValues` into
`object[]`) to 0.60 MB. The short job's wall-clock difference between the two 65K
rows is run-to-run noise; both parse exactly the same cells, and the reduction in
GC work matters on longer runs and memory-constrained targets.

Run the focused SST and access-pattern benchmarks with:

```powershell
dotnet run --project source/Benchmark/Benchmark.csproj -c Release -- --sst
```

XLSB read and both writers are unchanged by the worksheet scanner. The scanner
also removes the extra `BufferedStream` copy that used to sit in front of the
deflate stream.

### Numeric and date values

Numbers and dates go through a small decimal fast path before the general
`Utf8Parser` fallback. It accepts an optional sign, up to 19 digits, an optional
fraction and exponent, and only when the mantissa is at most 2^53 and the decimal
scale at most 22. In that range the mantissa and the tabulated power of ten are
both exact doubles, so one multiply or divide performs the single IEEE rounding
and the result is bit-identical to `Utf8Parser`. Values outside the range,
`NaN`/`Infinity`, whitespace and malformed text still fall back to
`Utf8Parser`, which keeps the accepted syntax and error behavior unchanged.

On the 65K fixture all 131,070 date values and 270,249 of the 320,048 plain
decimal values (89% of the values that reach the double parser) take the fast
path, with zero bit differences from `Utf8Parser` across the whole file. The
remaining 49,799 are 17-digit mantissas above 2^53 and fall back to
`Utf8Parser`.

In an isolated loop over representative values from the fixture the fast path
costs 8.0 ns per value. The built-in alternatives are `Utf8Parser` at 27.3 ns,
the .NET 11 `double.TryParsePartial` at 36.7 ns, `double.TryParse` at 37.3 ns
and `double.Parse` at 36.8 ns (all over `ReadOnlySpan<byte>` with
`NumberStyles` and the invariant culture). `Utf8Parser` is the fastest built-in
and is exactly what the fallback uses; `TryParsePartial` and the
`TryParse`/`Parse` overloads share the general `NumberStyles` parser and are
not competitive here. End to end, a repeated harness that opens the workbook
and calls `GetValues` for every row (25 warmups, 30 timed reads) dropped from
95.8-96.6 ms to 90.3-91.3 ms, about 5%; the 200K fixture stayed within
run-to-run noise.
Integers, style indexes and shared-string indexes use `Utf8Parser` directly. A
dedicated small-integer loop was about 1 ns faster per value (calamine uses the
SIMD `atoi_simd` crate for the same job), but that is only around 1 ms per 65K
read, so it did not justify the extra code.

### Prior art in other readers

The same two ideas used here - parsing `<v>` directly from bytes and using a
specialised float parser - are also used by other fast spreadsheet readers:

- [calamine](https://github.com/tafia/calamine) (Rust) parses `<v>` values for
  the plain `n`, `s`, `b`, `e` and untyped cells directly from raw bytes,
  bypassing its XML text decoder and `String` allocation. Its PR #622 reports
  an 8-10% gain on typical mixed workbooks. It parses every float with
  [`fast-float2`](https://crates.io/crates/fast-float2), a port of Lemire's
  `fast_float` (PR #608 reports about 3% on float-heavy sheets), uses
  `atoi_simd` for shared-string indexes and cell references, and precomputes
  style formats into an array instead of looking them up per cell.
- [Sylvan.Data.Excel](https://github.com/MarkPflug/Sylvan.Data.Excel) (.NET)
  reads through `XmlReader` into a reusable `char` buffer and parses numbers
  with `double.Parse(span, NumberStyles.Float, CultureInfo.InvariantCulture)`.
  It has no float fast path; since its PR #215 its XLSX reader parses field
  values lazily when a field is accessed, which helps when only a subset of
  columns is read but not for a full `GetValues` scan.
- [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) (.NET)
  calls `double.TryParse` with `NumberStyles.Any` on the `string` produced by
  `XmlReader`, and only has a small custom fast path for unsigned integers of
  at most nine digits.
- [SheetReader](https://github.com/polydbms/sheetreader-core) (C++/R/DuckDB)
  combines structure-aware parsing, parallel worksheet processing, `miniz`
  decompression and Lemire's `fast_double_parser`.

The main general-purpose alternative in .NET is
[csFastFloat](https://github.com/CarlVerret/csFastFloat), a direct port of
`fast_float` with a `ReadOnlySpan<byte>` overload. It is exact and handles every
input, but measured here it is slower than the narrow parser: 9.2 ns versus
6.5 ns on short values and 14.1 ns versus 10.5 ns on long ones (where the
narrow parser rejects the value and falls back to `Utf8Parser`). Replacing the
fast path with `FastDoubleParser.TryParseDouble` made the end-to-end 65K read
slower (91.5 ms to about 99-102 ms) and the 200K read slower as well, so it was
not adopted. An important part of the Eisel-Lemire algorithm used by these
libraries was merged into the .NET runtime for .NET 7, so the built-in
`double.Parse` and `Utf8Parser` already use it; the remaining gain comes from
specialising to Excel's short, invariant values rather than from a better
general algorithm.

### XLSB reads

XLSB cells already store numbers as binary IEEE doubles or RK integers, so the
reader decodes them with bit operations instead of parsing text. Two details
borrowed from calamine reduce the remaining overhead:

- The style-to-date classification is precomputed once after the styles part is
  loaded, the same way calamine builds its `CellFormat` array. Numeric cells no
  longer perform a `Dictionary<int, Type>` lookup for every styled cell.
- XLSB strings are UTF-16LE. On little-endian machines they are materialised
  with `new string(MemoryMarshal.Cast<byte, char>(...))` instead of the general
  `Encoding.Unicode.GetString` path, which is roughly 2.6x faster for the short
  strings in a shared-string table.

On the repeated harness `open200kXlsb` (decompress and decode the shared-string
part, without scanning the worksheet) dropped from about 44-49 ms to about
34-43 ms, roughly 18%. A full `read200kXlsb` with `GetValues` dropped from
about 89-91 ms to about 81-86 ms, roughly 5-8%. The 65K XLSB fixture only has a
5.6 KB shared-string part, so the UTF-16 change has no measurable effect there
and the style change is within run-to-run noise.

Streaming worksheet records directly from the deflate stream was also measured.
It reduces the per-read allocation from about 69 MB to 49 MB on the 200K
fixture, but is about 35% slower (111 ms versus 81 ms) because every record
header becomes a separate stream read. Reading the part into a byte array first
remains the default; the streaming variant is only used for parts that do not
fit in an `int`-sized array.

### Where the scan time goes

The measurements in this subsection are from the original .NET 10 scanner work
and are retained as an analysis of where the time goes; they are not directly
comparable to the .NET 11 table above. Two measurements bound what the scanner
can still win, both taken on `65K_Records_Data.xlsx` (32 MB of worksheet XML
after decompression):

| Measurement | Time |
|---|---:|
| decompress the sheet part only, no XML parsing | 20.1 ms |
| full read with the scanner (`GetValues` per row) | 99.7 ms |
| full read with the `XmlReader` path | 225.0 ms |

So decompression is roughly a fifth of the scanner path and sets a floor that no
parser change can go below. Of the remaining time, the largest single cost was
that every `<c>` start tag was walked once per attribute: `r`, `s` and `t` were
looked up by three independent scans of the same tag. All three names are one
character long, so they are now collected in a single pass
(`FastReadCellAttributes`), which is where the 109.1 ms → 92.9 ms step came from
(BenchmarkDotNet mean, same machine and SDK, allocations unchanged at 620 KB);
the current quick job reports 87.46 ms.

Buffer compaction was also measured and left alone: the scanner copies 32 MB in
4,901 small `Buffer.BlockCopy` calls per file, which sounds expensive but takes
0.24 ms in total. Removing it entirely made the 65k read 10% *slower* and raised
allocations tenfold, because the buffer then grows to hold the whole sheet
instead of staying hot in cache.

### Correctness corpus

`source/Tests/XlsxFastReaderParityTests.cs` reads every fixture twice, with the
scanner and with `XmlReader` (`_disableFastXlsxReader = true`), and compares
`FieldCount`, cell types, values and `RowCount`. The corpus contains:

- workbooks written by this library with shared and inline strings, including
  Unicode, XML metacharacters, control characters, a 20 000-character cell and
  more than 26 columns,
- a workbook written by `Sylvan.Data.Excel`,
- handcrafted SpreadsheetML for comments, processing instructions, missing row and
  cell references, sparse rows, self-closing and empty cells, formulas, error
  cells, rich inline strings, entities, date styles, large numbers, UTF-8 BOM and
  UTF-16 parts, and prefixed tags,
- an Excel-authored corpus generated by `tools/Create-XlsxFastReaderFixtures.ps1`.

The Excel script writes mixed-type, sparse, wide, formula/error, styled,
multi-sheet and 20 000-row workbooks to `test_output/fast-reader-fixtures/` and a
`manifest.json` with the expected cell values. The test verifies every manifest
value and then runs the fast-vs-XmlReader parity check on all files. The test
suite also writes this library's shared/inline-string workbooks and a
`Sylvan.Data.Excel` workbook into the same directory so they are part of the
corpus. Run it with:

```powershell
pwsh -File tools/Create-XlsxFastReaderFixtures.ps1
dotnet test source/Tests/Tests.csproj -c Release --filter "FullyQualifiedName~XlsxFastReaderParityTests"
```

## Measured results

Results below should be regenerated when the benchmark input files, SDK,
BenchmarkDotNet version or machine changes. They are intentionally recorded
with the environment instead of presented as universal performance claims.

### Environment

| Item | Value |
|---|---|
| OS | Windows 11 |
| Architecture | x64 |
| CPU | AMD Ryzen 7 7840HS, 16 logical / 8 physical cores |
| .NET SDK | 10.0.403 |
| .NET runtime | 10.0.14 |
| BenchmarkDotNet | 0.15.8 |
| Input data | 65K record XLSX/XLSB fixtures; 5,000-row generated write table |

### Results

The following short-run measurement was captured on the environment above.
Each job used one launch, one warmup and three iterations. `Error` is the
99.9% confidence-interval margin reported by BenchmarkDotNet.

| Benchmark | .NET 10 Mean | NativeAOT 10 Mean | .NET 10 Error | NativeAOT 10 Error | .NET 10 Allocated | NativeAOT 10 Allocated | Ratio |
|---|---:|---:|---:|---:|---:|---:|---:|
| `ReadXlsx` | 89.813 ms | 132.093 ms | 41.310 ms | 66.413 ms | 620.16 KB | 620.23 KB | 1.47 |
| `ReadXlsb` | 28.448 ms | 39.472 ms | 8.087 ms | 33.275 ms | 15,837.19 KB | 15,836.96 KB | 1.39 |
| `WriteXlsx` | 5.328 ms | 6.591 ms | 5.005 ms | 12.177 ms | 1,414.27 KB | 1,414.49 KB | 1.24 |
| `WriteXlsb` | 4.413 ms | 5.322 ms | 0.459 ms | 0.955 ms | 1,429.64 KB | 1,430.00 KB | 1.21 |

The measured values are a reproducibility sample, not a performance guarantee.
The complete BenchmarkDotNet report also includes `StdDev`, GC counts and
confidence intervals. This table was regenerated after the fast XLSX scanner
became the default, so `ReadXlsx` measures the scanner path rather than the
`XmlReader` path used by older artifacts.

### XLSX read comparison with `Sylvan.Data.Excel`

The `ReadBenchXlsx` full job compares both readers on the same fixtures with the
same access pattern per fixture: typed getters for the 65K file
(65,536 rows) and `GetValues` into an `object[]` for the 200K file
(200,001 rows). Both benchmarks used the default BenchmarkDotNet job on the
environment above.

| Fixture | SpreadSheetTasks | `Sylvan.Data.Excel` | Speed-up | Allocated (SpreadSheetTasks / Sylvan) |
|---|---:|---:|---:|---:|
| `65K_Records_Data.xlsx` | 93.70 ms | 168.63 ms | 1.80× | 620.17 KB / 664.80 KB |
| `200kFile.xlsx` | 162.92 ms | 351.63 ms | 2.16× | 35,059.50 KB / 52,326.09 KB |

Benchmark artifacts are generated under `BenchmarkDotNet.Artifacts` and are not
part of the NuGet package.
