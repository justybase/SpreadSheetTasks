using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Running;
using SpreadSheetTasks;
using Sylvan.Data.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.IO;
using System.IO.Compression;


namespace Benchmark;

internal static class Paths
{
    internal static readonly string FilesToTestDir = ResolveFilesToTestDir();

    private static string ResolveFilesToTestDir()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null)
        {
            var filesToTest = new DirectoryInfo(Path.Combine(dir.FullName, "FilesToTest"));
            if (filesToTest.Exists)
                return filesToTest.FullName;
            dir = dir.Parent;
        }
        throw new DirectoryNotFoundException("Cannot find FilesToTest directory. Run benchmarks from the SpreadSheetTasks repository root.");
    }
}

class Program
{
    static void Main(string[] args)
    {
#if RELEASE
        bool runFullBenchmarks = Array.Exists(args, a => string.Equals(a, "--full", StringComparison.OrdinalIgnoreCase));
        bool runSstBenchmarks = Array.Exists(args, a => string.Equals(a, "--sst", StringComparison.OrdinalIgnoreCase));
        if (runSstBenchmarks)
        {
            BenchmarkRunner.Run<ReadBenchXlsxSst>();
        }
        else if (runFullBenchmarks)
        {
            BenchmarkRunner.Run<ReadBenchXlsb>();
            BenchmarkRunner.Run<ReadBenchXlsx>();
            BenchmarkRunner.Run<ReadBenchXlsxSst>();
            BenchmarkRunner.Run<WriteBenchExcel>();
        }
        else
        {
            BenchmarkRunner.Run<ReadBenchXlsbQuick>();
            BenchmarkRunner.Run<ReadBenchXlsxQuick>();
            BenchmarkRunner.Run<WriteBenchExcelQuick>();
        }
#endif
#if DEBUG
        WriteBenchExcel writeBench = new WriteBenchExcel();
        writeBench.Setup();
        writeBench.XlsxSylvanWrite();
        writeBench.XlsxWrite();
        Console.WriteLine("PLEASE RUN IN RELEASE MODE");
#endif
    }
}
[SimpleJob(RuntimeMoniker.Net11_0)]
[MemoryDiagnoser]
public class ReadBenchXlsx
{
    readonly string filename65k = "65K_Records_Data.xlsx";
    readonly string filename200k = "200kFile.xlsx";

    [Benchmark]
    public void SpreadSheetTasks200K()
    {
        var path = Path.Combine(Paths.FilesToTestDir, filename200k);

        using XlsxOrXlsbReadOrEdit excelFile = new XlsxOrXlsbReadOrEdit();
        excelFile.Open(path);
        var sheetNames = excelFile.GetSheetNames();
        excelFile.ActualSheetName = sheetNames[0];
        object[] row = null;
        int rowNum = 0;
        while (excelFile.Read())
        {
            row ??= new object[excelFile.FieldCount];
            excelFile.GetValues(row);
            rowNum++;
#if DEBUG
                if (rowNum % 10_000 == 0)
                    Console.WriteLine("row " + rowNum.ToString("N0") + ": " + String.Join('|', row));
#endif
        }
    }

    [Benchmark]
    public void Sylvan200k()
    {
        var path = Path.Combine(Paths.FilesToTestDir, filename200k);

        var reader = ExcelDataReader.Create(path);

        object[] row = new object[reader.FieldCount];

        while (reader.Read())
        {
            reader.GetValues(row);
        }
    }

    [Benchmark]
    public void SpreadSheetTasks65k()
    {
        var path = Path.Combine(Paths.FilesToTestDir, filename65k);

        using XlsxOrXlsbReadOrEdit excelFile = new XlsxOrXlsbReadOrEdit();
        excelFile.Open(path);
        var sheetNames = excelFile.GetSheetNames();
        excelFile.ActualSheetName = sheetNames[0];

        excelFile.Read(); // = skip header
        while (excelFile.Read())
        {
            ProcessRecord(excelFile);
        }
    }

    [Benchmark]
    public void Sylvan65K()
    {
        var path = Path.Combine(Paths.FilesToTestDir, filename65k);

        var reader = Sylvan.Data.Excel.ExcelDataReader.Create(path);

        do
        {
            while (reader.Read())
            {
                ProcessRecordSylvan(reader);
            }

        } while (reader.NextResult());
    }

    //method from
    //https://github.com/MarkPflug/Benchmarks/blob/main/source/Benchmarks/XlsxDataReaderBenchmarks.cs
    internal static void ProcessRecordSylvan(IDataReader reader)
    {
        var region = reader.GetString(0);
        var country = reader.GetString(1);
        var type = reader.GetString(2);
        var channel = reader.GetString(3);
        var priority = reader.GetString(4);
        var orderDate = reader.GetDateTime(5);
        var id = reader.GetInt32(6);
        var shipDate = reader.GetDateTime(7);
        var unitsSold = reader.GetInt32(8);
        var unitPrice = reader.GetDouble(9);
        var unitCost = reader.GetDouble(10);
        var totalRevenue = reader.GetDouble(11);
        var totalCost = reader.GetDouble(12);
        var totalProfit = reader.GetDouble(13);
    }

    internal static void ProcessRecord(XlsxOrXlsbReadOrEdit excelFile)
    {
        var region = excelFile.GetString(0);
        var country = excelFile.GetString(1);
        var type = excelFile.GetString(2);
        var channel = excelFile.GetString(3);
        var priority = excelFile.GetString(4);
        var orderDate = excelFile.GetDateTime(5);
        var id = excelFile.GetInt32(6);
        var shipDate = excelFile.GetDateTime(7);
        var unitsSold = excelFile.GetInt32(8);
        var unitPrice = excelFile.GetDouble(9);
        var unitCost = excelFile.GetDouble(10);
        var totalRevenue = excelFile.GetDouble(11);
        var totalCost = excelFile.GetDouble(12);
        var totalProfit = excelFile.GetDouble(13);
    }
}

[SimpleJob(RuntimeMoniker.Net11_0)]
[MemoryDiagnoser]
public class ReadBenchXlsb
{
    [Params("65K_Records_Data.xlsb", "200kFile.xlsb")]
    public string FileName { get; set; }

    [Benchmark(Description = "SpreadSheetTasks - XLSB Read - v1")]
    public void ReadFile()
    {
        var path = Path.Combine(Paths.FilesToTestDir, FileName);
        using XlsxOrXlsbReadOrEdit excelFile = new XlsxOrXlsbReadOrEdit();
        excelFile.Open(path);
        excelFile.ActualSheetName = "sheet1";
        object[] row = null;
        int rowNum = 0;
        while (excelFile.Read())
        {
            row ??= new object[excelFile.FieldCount];
            excelFile.GetValues(row);
            rowNum++;
        }
    }

    [Benchmark(Description = "SpreadSheetTasks - XLSB Read - v2")]
    public void ReadFileInMemory()
    {
        var path = Path.Combine(Paths.FilesToTestDir, FileName);
        using XlsxOrXlsbReadOrEdit excelFile = new XlsxOrXlsbReadOrEdit();
        excelFile.Open(path);
        excelFile.UseMemoryStreamInXlsb = false;
        excelFile.ActualSheetName = "sheet1";
        object[] row = null;
        int rowNum = 0;
        while (excelFile.Read())
        {
            row ??= new object[excelFile.FieldCount];
            excelFile.GetValues(row);
            rowNum++;
        }
    }

    [Benchmark(Description = "Sylvan.Data.Excel - XLSB Read")]
    public void Sylvan()
    {
        var path = Path.Combine(Paths.FilesToTestDir, FileName);
        using ExcelDataReader reader = ExcelDataReader.Create(path);
        object[] row = new object[reader.FieldCount];
        while (reader.Read())
        {
            reader.GetValues(row);
        }
    }
}

[SimpleJob(RuntimeMoniker.Net11_0, warmupCount: 1, iterationCount: 3)]
[MemoryDiagnoser]
public class ReadBenchXlsxSst
{
    private readonly string _file200k = "200kFile.xlsx";
    private readonly string _file65k = "65K_Records_Data.xlsx";
    private string _entityFile = string.Empty;

    [GlobalSetup]
    public void Setup()
    {
        _entityFile = Path.Combine(Path.GetTempPath(), "sst_entity_rows.xlsx");
        var table = new DataTable();
        table.Columns.Add("Text", typeof(string));
        for (int row = 0; row < 20_000; row++)
        {
            table.Rows.Add($"row {row} & <tag> \"quote\" 'apos'");
        }

        using var writer = new XlsxWriter(_entityFile, useScharedStrings: false);
        writer.AddSheet("Data");
        writer.WriteSheet(table.CreateDataReader());
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        if (File.Exists(_entityFile))
        {
            File.Delete(_entityFile);
        }
    }

    [Benchmark(Description = "Open 200K with shared strings")]
    public int OpenWithSharedStrings()
    {
        using var reader = new XlsxOrXlsbReadOrEdit();
        reader.Open(Path.Combine(Paths.FilesToTestDir, _file200k), readSharedStrings: true);
        return reader.GetSheetNames().Length;
    }

    [Benchmark(Description = "Open 200K without shared strings")]
    public int OpenWithoutSharedStrings()
    {
        using var reader = new XlsxOrXlsbReadOrEdit();
        reader.Open(Path.Combine(Paths.FilesToTestDir, _file200k), readSharedStrings: false);
        return reader.GetSheetNames().Length;
    }

    [Benchmark(Description = "Read 200K - GetValues")]
    public long Read200kGetValues()
    {
        using var reader = new XlsxOrXlsbReadOrEdit();
        reader.Open(Path.Combine(Paths.FilesToTestDir, _file200k));
        reader.ActualSheetName = reader.GetSheetNames()[0];
        object[]? row = null;
        long rows = 0;
        while (reader.Read())
        {
            row ??= new object[reader.FieldCount];
            reader.GetValues(row);
            rows++;
        }
        return rows;
    }

    [Benchmark(Description = "Read 65K - GetValues")]
    public long Read65kGetValues()
    {
        using var reader = new XlsxOrXlsbReadOrEdit();
        reader.Open(Path.Combine(Paths.FilesToTestDir, _file65k));
        reader.ActualSheetName = reader.GetSheetNames()[0];
        object[]? row = null;
        long rows = 0;
        while (reader.Read())
        {
            row ??= new object[reader.FieldCount];
            reader.GetValues(row);
            rows++;
        }
        return rows;
    }

    [Benchmark(Description = "Read 65K - CurrentRow")]
    public long Read65kCurrentRow()
    {
        using var reader = new XlsxOrXlsbReadOrEdit();
        reader.Open(Path.Combine(Paths.FilesToTestDir, _file65k));
        reader.ActualSheetName = reader.GetSheetNames()[0];
        long cells = 0;
        while (reader.Read())
        {
            cells += reader.CurrentRow.Length;
        }
        return cells;
    }

    [Benchmark(Description = "Read 20K entity inline strings")]
    public long ReadEntityRows()
    {
        using var reader = new XlsxOrXlsbReadOrEdit();
        reader.Open(_entityFile);
        reader.ActualSheetName = reader.GetSheetNames()[0];
        object[]? row = null;
        long rows = 0;
        while (reader.Read())
        {
            row ??= new object[reader.FieldCount];
            reader.GetValues(row);
            rows++;
        }
        return rows;
    }
}

[SimpleJob(RuntimeMoniker.Net11_0)]
[MemoryDiagnoser]
public class WriteBenchExcel
{
    public int RowsCount = 200_000;
    private readonly Dictionary<string,DataTable> _dataTables = new Dictionary<string, DataTable>
    {
        {"GENERAL", new DataTable() },
        {"INT", new DataTable() },
        {"DOUBLE", new DataTable() },
        {"DATETIME", new DataTable() },
        {"TEXT", new DataTable() },
    };

    //[Params("GENERAL", "INT", "DOUBLE", "DATETIME", "TEXT")]
    [Params("GENERAL")]
    public string ReaderType { get; set; } = "GENERAL";
    private DataTable Dt => _dataTables[ReaderType];

    [GlobalSetup]
    public void Setup()
    {
        _dataTables["GENERAL"].Columns.Add("COL1_INT", typeof(int));
        _dataTables["GENERAL"].Columns.Add("COL2_TXT", typeof(string));
        _dataTables["GENERAL"].Columns.Add("COL3_DATETIME", typeof(DateTime));
        _dataTables["GENERAL"].Columns.Add("COL4_DOUBLE", typeof(double));

        Random rn = new Random(420);

        for (int i = 0; i < RowsCount; i++)
        {
            _dataTables["GENERAL"].Rows.Add(new object[]
            {
                rn.Next(0, 1_000_000),
                "TXT_" + rn.Next(0, 1_000_000),
                new DateTime(rn.Next(1900,2100), rn.Next(1,12), rn.Next(1, 28)),
                rn.NextDouble()
            });
        }

        _dataTables["INT"].Columns.Add("COL1", typeof(int));
        _dataTables["INT"].Columns.Add("COL2", typeof(int));
        _dataTables["INT"].Columns.Add("COL3", typeof(int));
        _dataTables["INT"].Columns.Add("COL4", typeof(int));

        for (int i = 0; i < RowsCount; i++)
        {
            _dataTables["INT"].Rows.Add(new object[]
            {
                rn.Next(),
                rn.Next(),
                rn.Next(),
                rn.Next(),
            });
        }


        _dataTables["DOUBLE"].Columns.Add("COL1", typeof(double));
        _dataTables["DOUBLE"].Columns.Add("COL2", typeof(double));
        _dataTables["DOUBLE"].Columns.Add("COL3", typeof(double));
        _dataTables["DOUBLE"].Columns.Add("COL4", typeof(double));

        for (int i = 0; i < RowsCount; i++)
        {
            _dataTables["DOUBLE"].Rows.Add(new object[]
            {
                rn.NextDouble(),
                rn.NextDouble(),
                rn.NextDouble(),
                rn.NextDouble(),
            });
        }

        _dataTables["DATETIME"].Columns.Add("COL1", typeof(DateTime));
        _dataTables["DATETIME"].Columns.Add("COL2", typeof(DateTime));
        _dataTables["DATETIME"].Columns.Add("COL3", typeof(DateTime));
        _dataTables["DATETIME"].Columns.Add("COL4", typeof(DateTime));

        for (int i = 0; i < RowsCount; i++)
        {
            _dataTables["DATETIME"].Rows.Add(new object[]
            {
                DateTime.Now.AddSeconds(i),
                DateTime.Now.AddSeconds(i+1),
                DateTime.Now.AddSeconds(i+2),
                DateTime.Now.AddSeconds(i+3),
            });
        }

        _dataTables["TEXT"].Columns.Add("COL1", typeof(string));
        _dataTables["TEXT"].Columns.Add("COL2", typeof(string));
        _dataTables["TEXT"].Columns.Add("COL3", typeof(string));
        _dataTables["TEXT"].Columns.Add("COL4", typeof(string));

        for (int i = 0; i < RowsCount; i++)
        {
            _dataTables["TEXT"].Rows.Add(new object[]
            {
                "TXT_" + rn.Next(0, 1_000_000),
                "TXT_" + rn.Next(0, 1_000_000),
                "TXT_" + rn.Next(0, 1_000_000),
                "TXT_" + rn.Next(0, 1_000_000),
            });
        }
    }

    private readonly CompressionLevel _cLvl = CompressionLevel.Fastest;

    [Benchmark(Description = "SpreadSheetTasks - XLSB Write")]
    public void XlsbWriteDefault()
    {
        using XlsbWriter xlsx = new XlsbWriter("file.xlsb", _cLvl);
        xlsx.AddSheet("sheetName");
        xlsx.WriteSheet(Dt.CreateDataReader(), doAutofilter: true);
    }

    [Benchmark]
    public void XlsbSylvanWrite()
    {
        using var edw = ExcelDataWriter.Create("fileSylvan.xlsb", new ExcelDataWriterOptions() {  CompressionLevel = _cLvl });
        DbDataReader dr;
        dr = Dt.CreateDataReader();
        edw.Write(dr, "sheetName");
    }

    [Benchmark(Description = "SpreadSheetTasks - XLSX Sylvan Write")]
    public void XlsxSylvanWrite()
    {
        using var edw = ExcelDataWriter.Create("fileSylvan.xlsx", new ExcelDataWriterOptions() { CompressionLevel = _cLvl });
        DbDataReader dr;
        dr = Dt.CreateDataReader();
        edw.Write(dr, "sheetName");
    }

    [Benchmark(Description = "SpreadSheetTasks - XLSX Write")]
    public void XlsxWrite()
    {
        using XlsxWriter xlsx = new XlsxWriter("file.xlsx", _clvl: this._cLvl);
        xlsx.AddSheet("sheetName");
        xlsx.WriteSheet(Dt.CreateDataReader());
    }
}

[SimpleJob(RuntimeMoniker.Net11_0, launchCount: 1, warmupCount: 1, iterationCount: 3)]
[MemoryDiagnoser]
public class ReadBenchXlsxQuick
{
    readonly string _fileName65k = "65K_Records_Data.xlsx";

    [Benchmark]
    public void SpreadSheetTasks65k()
    {
        var path = Path.Combine(Paths.FilesToTestDir, _fileName65k);

        using XlsxOrXlsbReadOrEdit excelFile = new XlsxOrXlsbReadOrEdit();
        excelFile.Open(path);
        var sheetNames = excelFile.GetSheetNames();
        excelFile.ActualSheetName = sheetNames[0];

        excelFile.Read();
        while (excelFile.Read())
        {
            ReadBenchXlsx.ProcessRecord(excelFile);
        }
    }
}

[SimpleJob(RuntimeMoniker.Net11_0, launchCount: 1, warmupCount: 1, iterationCount: 3)]
[MemoryDiagnoser]
public class ReadBenchXlsbQuick
{
    private readonly string _fileName = "65K_Records_Data.xlsb";

    [Benchmark(Description = "SpreadSheetTasks - XLSB Read - v1 (quick)")]
    public void ReadFile()
    {
        var path = Path.Combine(Paths.FilesToTestDir, _fileName);
        using XlsxOrXlsbReadOrEdit excelFile = new XlsxOrXlsbReadOrEdit();
        excelFile.Open(path);
        excelFile.ActualSheetName = "sheet1";
        object[] row = null;
        while (excelFile.Read())
        {
            row ??= new object[excelFile.FieldCount];
            excelFile.GetValues(row);
        }
    }

    [Benchmark(Description = "SpreadSheetTasks - XLSB Read - v2 (quick)")]
    public void ReadFileInMemory()
    {
        var path = Path.Combine(Paths.FilesToTestDir, _fileName);
        using XlsxOrXlsbReadOrEdit excelFile = new XlsxOrXlsbReadOrEdit();
        excelFile.Open(path);
        excelFile.UseMemoryStreamInXlsb = false;
        excelFile.ActualSheetName = "sheet1";
        object[] row = null;
        while (excelFile.Read())
        {
            row ??= new object[excelFile.FieldCount];
            excelFile.GetValues(row);
        }
    }
}

[SimpleJob(RuntimeMoniker.Net11_0, launchCount: 1, warmupCount: 1, iterationCount: 3)]
[MemoryDiagnoser]
public class WriteBenchExcelQuick : WriteBenchExcel
{
    public WriteBenchExcelQuick()
    {
        RowsCount = 50_000;
    }
}

