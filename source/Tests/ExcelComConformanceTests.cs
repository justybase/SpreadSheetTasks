using System.Diagnostics;
using System.Globalization;
using System.Runtime.Versioning;
using System.Text.Json;
using SpreadSheetTasks;

namespace Tests;

// Excel COM conformance suite driven by the shared spec in comparison/com/cases.json.
//
// Writer tests: SpreadSheetTasks writes every case as XLSX and XLSB; Microsoft Excel (COM) opens each
// file and checks sheet names, visibility, AutoFilter, every cell value, the used range and number formats.
// Reader tests: Excel writes the same cases; SpreadSheetTasks reads them back and must return the spec values.
// Updater tests: an Excel-authored template is updated in place; Excel verifies the result.
//
// Requires Windows, Microsoft Excel with COM automation and PowerShell 7 (pwsh).
[Collection("Sequential")]
[SupportedOSPlatform("windows")]
public class ExcelComConformanceTests
{
    private static readonly string[] Formats = ["xlsx", "xlsb"];

    public static IEnumerable<object[]> CaseAndFormat() => SpecCases.Ids.SelectMany(id => Formats.Select(f => new object[] { id, f }));

    private readonly ComSession _session = ComSession.Instance;

    [Theory]
    [MemberData(nameof(CaseAndFormat))]
    public void Writer_output_has_expected_structure(string caseId, string format) =>
        AssertWriterChecks(caseId, format, check => !check.StartsWith("cells:", StringComparison.Ordinal) && !check.StartsWith("numberFormat:", StringComparison.Ordinal));

    [Theory]
    [MemberData(nameof(CaseAndFormat))]
    public void Writer_output_has_expected_cell_values(string caseId, string format) =>
        AssertWriterChecks(caseId, format, check => check.StartsWith("cells:", StringComparison.Ordinal) || check.StartsWith("numberFormat:", StringComparison.Ordinal));

    private void AssertWriterChecks(string caseId, string format, Func<string, bool> include)
    {
        if (!_session.ExcelAvailable) { Console.WriteLine("SKIP — Excel not installed or COM unavailable"); return; }
        var results = _session.WriterResults(caseId, format).Where(r => include(r.Check)).ToList();
        Assert.True(results.Count > 0, "Excel produced no checks for this file");
        var failures = results.Where(r => !r.Pass).Select(r => $"{r.Check} [{r.Sheet}]: expected '{r.Expected}', got '{r.Actual}' {r.Message}").ToList();
        Assert.True(failures.Count == 0, string.Join(Environment.NewLine, failures));
    }

    [Theory]
    [MemberData(nameof(CaseAndFormat))]
    public void Reader_reads_excel_authored_file(string caseId, string format)
    {
        if (!_session.ExcelAvailable) { Console.WriteLine("SKIP — Excel not installed or COM unavailable"); return; }
        var path = _session.FixturePath(caseId, format);
        var sheets = SpecCases.Get(caseId).Sheets;

        using var reader = new XlsxOrXlsbReadOrEdit();
        reader.Open(path);
        Assert.Equal(sheets.Select(s => s.Name), reader.GetSheetNames());

        foreach (var sheet in sheets)
        {
            var expected = sheet.Grid;
            if (expected.Count == 0) continue;
            int width = expected.Max(r => r.Length);
            var actual = reader.GetRowsOfSheet(sheet.Name).Select(r => r.ToArray()).ToList();
            Assert.True(actual.Count >= expected.Count, $"{sheet.Name}: expected at least {expected.Count} rows, got {actual.Count}");
            for (int r = 0; r < expected.Count; r++)
            {
                for (int c = 0; c < width; c++)
                {
                    object? exp = c < expected[r].Length ? expected[r][c] : null;
                    object? act = r < actual.Count && c < actual[r].Length ? actual[r][c] : null;
                    Assert.True(SpecCases.SameCell(exp, act), $"{sheet.Name} row {r + 1} col {c + 1}: expected '{exp}', got '{act}'");
                }
            }
        }
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsb")]
    public void Updater_replaces_sheet_and_excel_verifies(string format)
    {
        if (!_session.ExcelAvailable) { Console.WriteLine("SKIP — Excel not installed or COM unavailable"); return; }
        var target = Path.Combine(_session.OutDir, $"updated-{Guid.NewGuid():N}.{format}");
        File.Copy(_session.FixturePath("update_template", format), target);

        var rows = SpecCases.Get("update_result").Sheets[0].Grid.Select(r => (object?[])r).ToArray();
        if (format == "xlsx")
        {
            using var updater = new XlsxUpdater(target);
            updater.ReplaceSheetData("Data", rows);
            updater.Save(target);
        }
        else
        {
            using var updater = new XlsbUpdater(target);
            updater.ReplaceSheetData("Data", rows);
            updater.Save(target);
        }

        var results = ComSession.Verify(new[] { new VerifyTarget(target, "update_result", format) }, _session.OutDir);
        var failures = results.Where(r => !r.Pass).Select(r => $"{r.Check} [{r.Sheet}]: expected '{r.Expected}', got '{r.Actual}' {r.Message}").ToList();
        Assert.True(failures.Count == 0, string.Join(Environment.NewLine, failures));
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// Spec loading (mirror of Expand-SheetRows in comparison/com/ExcelComSpec.psm1)
// ─────────────────────────────────────────────────────────────────────────────

internal sealed record SpecSheet(string Name, bool Hidden, bool AutoFilter, IReadOnlyList<object?[]> Grid, IReadOnlyDictionary<string, string> NumberFormats);

internal sealed class SpecCase
{
    public required string Id { get; init; }
    public required IReadOnlyList<SpecSheet> Sheets { get; init; }
}

internal static class SpecCases
{
    private static readonly Lazy<Dictionary<string, SpecCase>> Cases = new(Load);

    public static IReadOnlyList<string> Ids => Cases.Value.Keys.ToList();

    public static SpecCase Get(string id) => Cases.Value[id];

    public static bool SameCell(object? expected, object? actual)
    {
        if (actual is DBNull) actual = null;
        if (expected is null) return actual is null || (actual is string s && s.Length == 0);
        if (expected is bool eb) return actual is bool ab && ab == eb;
        if (expected is string es) return actual is string aStr && string.Equals(aStr, es, StringComparison.Ordinal);
        if (expected is DateTime ed)
        {
            // GetRowsOfSheet returns date-formatted cells as raw serials (Int64 when whole days, Double otherwise).
            DateTime? ad = actual switch
            {
                DateTime d => d,
                double serial => DateTime.FromOADate(serial),
                long serial => DateTime.FromOADate(serial),
                _ => null,
            };
            return ad.HasValue && Math.Abs((ad.Value - ed).TotalSeconds) <= 1;
        }
        if (expected is double ef)
        {
            double? af = actual switch
            {
                double d => d,
                float f => f,
                int i => i,
                long l => l,
                decimal m => (double)m,
                _ => null,
            };
            return af.HasValue && Math.Abs(af.Value - ef) <= 1e-9 * Math.Max(1.0, Math.Abs(ef));
        }
        return Equals(expected, actual);
    }

    private static Dictionary<string, SpecCase> Load()
    {
        var specPath = FindSpecPath();
        using var doc = JsonDocument.Parse(File.ReadAllText(specPath));
        var result = new Dictionary<string, SpecCase>(StringComparer.Ordinal);
        foreach (var caseEl in doc.RootElement.GetProperty("cases").EnumerateArray())
        {
            var id = caseEl.GetProperty("id").GetString()!;
            var sheets = caseEl.GetProperty("sheets").EnumerateArray().Select(ParseSheet).ToList();
            result[id] = new SpecCase { Id = id, Sheets = sheets };
        }
        return result;
    }

    private static string FindSpecPath()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            var candidate = Path.Combine(dir.FullName, "comparison", "com", "cases.json");
            if (File.Exists(candidate)) return candidate;
            dir = dir.Parent;
        }
        throw new FileNotFoundException("comparison/com/cases.json not found above " + AppContext.BaseDirectory);
    }

    private static SpecSheet ParseSheet(JsonElement sheet)
    {
        var name = sheet.GetProperty("name").GetString()!;
        bool hidden = sheet.TryGetProperty("hidden", out var h) && h.GetBoolean();
        bool autoFilter = sheet.TryGetProperty("autoFilter", out var a) && a.GetBoolean();
        var numberFormats = new Dictionary<string, string>(StringComparer.Ordinal);
        var grid = ExpandGrid(sheet, numberFormats);
        return new SpecSheet(name, hidden, autoFilter, grid, numberFormats);
    }

    private static List<object?[]> ExpandGrid(JsonElement sheet, Dictionary<string, string> numberFormats)
    {
        var grid = new List<object?[]>();
        if (sheet.TryGetProperty("generator", out var gen))
        {
            int count = gen.GetProperty("rows").GetInt32();
            if (gen.TryGetProperty("columns", out var colsEl))
            {
                int cols = colsEl.GetInt32();
                grid.Add(Enumerable.Range(1, cols).Select(j => (object?)$"C{j:D2}").ToArray());
                for (int i = 1; i <= count; i++)
                    grid.Add(Enumerable.Range(1, cols).Select(j => (object?)(double)(i * 1000 + j)).ToArray());
                return grid;
            }
            grid.Add(gen.GetProperty("header").EnumerateArray().Select(e => (object?)e.GetString()).ToArray());
            for (int i = 1; i <= count; i++)
                grid.Add(new object?[] { (double)i, $"row-{i}", i * 0.25 });
            return grid;
        }

        int rowIndex = 0;
        foreach (var rowEl in sheet.GetProperty("rows").EnumerateArray())
        {
            rowIndex++;
            var row = new List<object?>();
            int colIndex = 0;
            foreach (var cell in rowEl.EnumerateArray())
            {
                colIndex++;
                row.Add(ParseCell(cell, rowIndex, colIndex, numberFormats));
            }
            grid.Add(row.ToArray());
        }
        return grid;
    }

    private static object? ParseCell(JsonElement cell, int row, int col, Dictionary<string, string> numberFormats)
    {
        switch (cell.ValueKind)
        {
            case JsonValueKind.Null: return null;
            case JsonValueKind.String: return cell.GetString();
            case JsonValueKind.True: return true;
            case JsonValueKind.False: return false;
            case JsonValueKind.Number: return cell.GetDouble();
            case JsonValueKind.Object:
                if (cell.TryGetProperty("date", out var d))
                    return DateTime.ParseExact(d.GetString()!, "yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);
                if (cell.TryGetProperty("numberFormat", out var nf))
                    numberFormats[CellAddress(row, col)] = nf.GetString()!;
                return ParseCell(cell.GetProperty("value"), row, col, numberFormats);
            default:
                throw new InvalidDataException($"Unsupported cell kind {cell.ValueKind}");
        }
    }

    internal static string CellAddress(int row, int col) => $"{ColumnName(col)}{row}";

    private static string ColumnName(int index)
    {
        var name = string.Empty;
        for (int n = index; n > 0; n = (n - 1) / 26)
            name = (char)('A' + (n - 1) % 26) + name;
        return name;
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// Excel COM session: writes all cases, verifies them, creates Excel fixtures once per test run.
// ─────────────────────────────────────────────────────────────────────────────

internal sealed record VerifyTarget(string File, string CaseId, string Format);

internal sealed record VerifyResult(string CaseId, string Format, string File, string Sheet, string Check, bool Pass, string? Expected, string? Actual, string Message);

internal sealed class ComSession
{
    private static readonly Lazy<ComSession> Lazy = new(() => new ComSession());

    public static ComSession Instance => Lazy.Value;

    public string OutDir { get; } = Path.Combine(Path.GetTempPath(), "spreadsheettasks-com-" + Guid.NewGuid().ToString("N"));

    public bool ExcelAvailable { get; }

    private readonly Dictionary<(string, string), string> _fixtures = new();
    private readonly List<VerifyResult> _writerResults = new();

    private ComSession()
    {
        ExcelAvailable = IsExcelInstalled();
        if (!ExcelAvailable) return;

        Directory.CreateDirectory(OutDir);
        var written = Path.Combine(OutDir, "written");
        Directory.CreateDirectory(written);

        var targets = new List<VerifyTarget>();
        foreach (var id in SpecCases.Ids)
            foreach (var fmt in new[] { "xlsx", "xlsb" })
            {
                var path = Path.Combine(written, $"{id}.{fmt}");
                WriteCase(SpecCases.Get(id), path, fmt);
                targets.Add(new VerifyTarget(path, id, fmt));
            }
        _writerResults.AddRange(Verify(targets, OutDir));

        var fixtureDir = Path.Combine(OutDir, "excel-fixtures");
        RunPwsh("New-ExcelFixture.ps1", "-SpecPath", SpecPath, "-OutDir", fixtureDir);
        using var doc = JsonDocument.Parse(File.ReadAllText(Path.Combine(fixtureDir, "targets.json")));
        foreach (var t in doc.RootElement.EnumerateArray())
            _fixtures[(t.GetProperty("caseId").GetString()!, t.GetProperty("format").GetString()!)] = t.GetProperty("file").GetString()!;
    }

    public IEnumerable<VerifyResult> WriterResults(string caseId, string format) =>
        _writerResults.Where(r => r.CaseId == caseId && r.Format == format);

    public string FixturePath(string caseId, string format) => _fixtures[(caseId, format)];

    private static string SpecPath => Path.Combine(ComDir, "cases.json");

    private static string ComDir
    {
        get
        {
            var dir = new DirectoryInfo(AppContext.BaseDirectory);
            while (dir is not null)
            {
                var candidate = Path.Combine(dir.FullName, "comparison", "com");
                if (Directory.Exists(candidate)) return candidate;
                dir = dir.Parent;
            }
            throw new DirectoryNotFoundException("comparison/com not found above " + AppContext.BaseDirectory);
        }
    }

    private static void WriteCase(SpecCase specCase, string path, string format)
    {
        if (format == "xlsx")
        {
            using var writer = new XlsxWriter(path);
            WriteSheets(writer, specCase);
        }
        else
        {
            using var writer = new XlsbWriter(path);
            WriteSheets(writer, specCase);
        }
    }

    private static void WriteSheets(ExcelWriter writer, SpecCase specCase)
    {
        foreach (var sheet in specCase.Sheets)
        {
            writer.AddSheet(sheet.Name, sheet.Hidden);
            var rows = sheet.Grid.ToArray();
            if (sheet.NumberFormats.Count == 0)
            {
                writer.WriteSheet(rows, startingRow: 0, startingColumn: 0, doAutofilter: sheet.AutoFilter, headers: false);
            }
            else
            {
                var styles = sheet.NumberFormats.ToDictionary(kv => kv.Key, kv => new CellStyle { NumberFormat = kv.Value });
                writer.WriteSheet(rows, headerStyle: null, styles: styles, doAutofilter: sheet.AutoFilter, headers: false);
            }
        }
    }

    internal static List<VerifyResult> Verify(IEnumerable<VerifyTarget> targets, string outDir)
    {
        var targetsPath = Path.Combine(outDir, $"verify-targets-{Guid.NewGuid():N}.json");
        var resultsPath = Path.Combine(outDir, $"verify-results-{Guid.NewGuid():N}.json");
        var payload = targets.Select(t => new { file = t.File, caseId = t.CaseId, format = t.Format });
        File.WriteAllText(targetsPath, JsonSerializer.Serialize(payload));
        RunPwsh("Invoke-ExcelVerify.ps1", "-SpecPath", SpecPath, "-TargetsPath", targetsPath, "-OutPath", resultsPath);
        using var doc = JsonDocument.Parse(File.ReadAllText(resultsPath));
        return doc.RootElement.EnumerateArray().Select(ParseResult).ToList();
    }

    private static VerifyResult ParseResult(JsonElement e) => new(
        CaseId: e.GetProperty("caseId").GetString()!,
        Format: e.GetProperty("format").GetString()!,
        File: e.GetProperty("file").GetString()!,
        Sheet: e.GetProperty("sheet").GetString() ?? string.Empty,
        Check: e.GetProperty("check").GetString()!,
        Pass: e.GetProperty("pass").GetBoolean(),
        Expected: e.GetProperty("expected").GetString(),
        Actual: e.GetProperty("actual").GetString(),
        Message: e.GetProperty("message").GetString() ?? string.Empty);

    private static void RunPwsh(string script, params string[] args)
    {
        var psi = new ProcessStartInfo("pwsh")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
        };
        foreach (var a in new[] { "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", Path.Combine(ComDir, script) }.Concat(args))
            psi.ArgumentList.Add(a);
        using var process = Process.Start(psi) ?? throw new InvalidOperationException("pwsh could not be started");
        var stdout = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();
        if (!process.WaitForExit(900_000)) { process.Kill(); throw new TimeoutException(script); }
        if (process.ExitCode != 0 && script != "Invoke-ExcelVerify.ps1")
            throw new InvalidOperationException($"{script} failed:\n{stdout}\n{stderr}");
    }

    private static bool IsExcelInstalled()
    {
        try { return Type.GetTypeFromProgID("Excel.Application") is not null; }
        catch { return false; }
    }
}
