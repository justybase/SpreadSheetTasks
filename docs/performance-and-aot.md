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

1. `source/Benchmark` keeps the broader .NET 10 comparison suite, including
   the existing Sylvan comparison.
2. `source/BenchmarkAot` is deliberately limited to SpreadSheetTasks APIs so
   that the NativeAOT job is not coupled to the AOT status of comparison
   libraries.

Run the focused .NET 10 versus NativeAOT benchmark:

```powershell
dotnet run --project source/BenchmarkAot/BenchmarkAot.csproj `
  -c Release -- --filter "*NativeAotReadWriteBenchmarks*"
```

The benchmark class contains both jobs:

- `RuntimeMoniker.Net10_0` — regular .NET 10 JIT,
- `RuntimeMoniker.NativeAot10_0` — .NET 10 NativeAOT.

The NativeAOT job is built by BenchmarkDotNet, so the first run can take
significantly longer than subsequent runs. BenchmarkDotNet documents the
NativeAOT toolchain and runtime-specific benchmark jobs in its
[toolchain guide](https://benchmarkdotnet.org/articles/configs/toolchains.html).

## Measured results

Results below should be regenerated when the benchmark input files, SDK,
BenchmarkDotNet version or machine changes. They are intentionally recorded
with the environment instead of presented as universal performance claims.

### Environment

| Item | Value |
|---|---|
| OS | Windows 11 |
| Architecture | x64 |
| .NET SDK | 10.0.302 |
| .NET runtime | 10.0.10 |
| BenchmarkDotNet | 0.15.8 |
| Input data | 65K record XLSX/XLSB fixtures; 5,000-row generated write table |

### Results

The following short-run measurement was captured on the environment above.
Each job used one launch, one warmup and three iterations. `Error` is the
99.9% confidence-interval margin reported by BenchmarkDotNet.

| Benchmark | .NET 10 Mean | NativeAOT 10 Mean | .NET 10 Error | NativeAOT 10 Error | .NET 10 Allocated | NativeAOT 10 Allocated | Ratio |
|---|---:|---:|---:|---:|---:|---:|---:|
| `ReadXlsx` | 175.034 ms | 271.037 ms | 122.721 ms | 47.866 ms | 594.05 KB | 594.44 KB | 1.55 |
| `ReadXlsb` | 30.265 ms | 48.308 ms | 19.866 ms | 69.610 ms | 15,776.32 KB | 15,776.79 KB | 1.60 |
| `WriteXlsx` | 9.178 ms | 10.859 ms | 4.295 ms | 10.104 ms | 1,649.59 KB | 1,649.66 KB | 1.18 |
| `WriteXlsb` | 3.242 ms | 8.393 ms | 2.865 ms | 1.162 ms | 1,244.52 KB | 1,244.60 KB | 2.59 |

The measured values are a reproducibility sample, not a performance guarantee.
The complete BenchmarkDotNet report also includes `StdDev`, GC counts and
confidence intervals.

Benchmark artifacts are generated under `BenchmarkDotNet.Artifacts` and are not
part of the NuGet package.
