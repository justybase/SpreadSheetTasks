using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Running;
using SpreadSheetTasks;
using System.Data;

namespace BenchmarkAot;

internal static class Program
{
    public static void Main(string[] args)
        => BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
}

[SimpleJob(RuntimeMoniker.Net10_0, baseline: true, launchCount: 1, warmupCount: 1, iterationCount: 3)]
[SimpleJob(RuntimeMoniker.NativeAot10_0, launchCount: 1, warmupCount: 1, iterationCount: 3)]
[MemoryDiagnoser]
public class NativeAotReadWriteBenchmarks
{
    private readonly string _filesDirectory = Path.Combine(AppContext.BaseDirectory, "FilesToTest");
    private DataTable _data = null!;

    [GlobalSetup]
    public void Setup()
    {
        _data = new DataTable();
        _data.Columns.Add("Id", typeof(int));
        _data.Columns.Add("Name", typeof(string));
        _data.Columns.Add("Value", typeof(double));
        _data.Columns.Add("Date", typeof(DateTime));

        for (int i = 0; i < 5_000; i++)
            _data.Rows.Add(i, "Row_" + i, i * 0.25, new DateTime(2025, 1, 1).AddDays(i % 365));
    }

    [Benchmark]
    public void ReadXlsx()
    {
        string path = Path.Combine(_filesDirectory, "65K_Records_Data.xlsx");
        using var reader = new XlsxOrXlsbReadOrEdit();
        reader.Open(path);
        reader.ActualSheetName = reader.GetSheetNames()[0];
        while (reader.Read())
            _ = reader.GetValue(0);
    }

    [Benchmark]
    public void ReadXlsb()
    {
        string path = Path.Combine(_filesDirectory, "65K_Records_Data.xlsb");
        using var reader = new XlsxOrXlsbReadOrEdit();
        reader.Open(path);
        reader.ActualSheetName = "sheet1";
        while (reader.Read())
            _ = reader.GetValue(0);
    }

    [Benchmark]
    public void WriteXlsx()
    {
        string path = Path.Combine(Path.GetTempPath(), "SpreadSheetTasks-AotBenchmark.xlsx");
        using var writer = new XlsxWriter(path);
        writer.AddSheet("Data");
        writer.WriteSheet(_data.CreateDataReader());
    }

    [Benchmark]
    public void WriteXlsb()
    {
        string path = Path.Combine(Path.GetTempPath(), "SpreadSheetTasks-AotBenchmark.xlsb");
        using var writer = new XlsbWriter(path);
        writer.AddSheet("Data");
        writer.WriteSheet(_data.CreateDataReader());
    }
}
