using SpreadSheetTasks;
using System.Text.Json;

namespace Tests;

[Collection("Sequential")]
public class UpdaterPivotTests
{
    [Fact]
    public void UpdateExcelGeneratedPivotFixtures()
    {
        string directory = ResolveRepositoryTestOutput();
        string xlsxSource = Path.Combine(directory, "updater_pivot_source.xlsx");
        string xlsbSource = Path.Combine(directory, "updater_pivot_source.xlsb");
        if (!File.Exists(xlsxSource) || !File.Exists(xlsbSource))
        {
            Console.WriteLine("Excel pivot fixtures are not present; run Create-ExcelPivotFixture.ps1 first.");
            return;
        }

        string xlsxOutput = Path.Combine(directory, "updater_pivot_updated.xlsx");
        string xlsbOutput = Path.Combine(directory, "updater_pivot_updated.xlsb");

        using (var updater = new XlsxUpdater(xlsxSource))
        {
            updater.ReplaceSheetData("Data", new object?[][]
            {
                ["A", 100],
                ["B", 20],
                ["C", 30]
            }, new ReplaceSheetDataOptions { Headers = ["Category", "Amount"] });
            updater.Save(xlsxOutput);
        }

        using (var updater = new XlsbUpdater(xlsbSource))
        {
            updater.ReplaceSheetData("Data", new object?[][]
            {
                ["A", 100],
                ["B", 20],
                ["C", 30]
            }, new ReplaceSheetDataOptions { Headers = ["Category", "Amount"] });
            updater.Save(xlsbOutput);
        }

        var manifest = new
        {
            files = new[]
            {
                new
                {
                    path = xlsxOutput,
                    sheets = new[] { "Report", "Data" },
                    cells = new object[]
                    {
                        new { sheet = "Data", address = "A1", value = "Category" },
                        new { sheet = "Data", address = "A4", value = "C" },
                        new { sheet = "Data", address = "B4", value = 30 }
                    },
                    pivots = new[]
                    {
                        new
                        {
                            sheet = "Report",
                            name = "SalesPivot",
                            refreshCells = new object[]
                            {
                                new { address = "A6", value = "C" },
                                new { address = "B6", value = 30 },
                                new { address = "B7", value = 150 }
                            }
                        }
                    }
                },
                new
                {
                    path = xlsbOutput,
                    sheets = new[] { "Report", "Data" },
                    cells = new object[]
                    {
                        new { sheet = "Data", address = "A1", value = "Category" },
                        new { sheet = "Data", address = "A4", value = "C" },
                        new { sheet = "Data", address = "B4", value = 30 }
                    },
                    pivots = new[]
                    {
                        new
                        {
                            sheet = "Report",
                            name = "SalesPivot",
                            refreshCells = new object[]
                            {
                                new { address = "A6", value = "C" },
                                new { address = "B6", value = 30 },
                                new { address = "B7", value = 150 }
                            }
                        }
                    }
                }
            }
        };

        File.WriteAllText(
            Path.Combine(directory, "updater_pivot_manifest.json"),
            JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true }));

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(xlsxOutput);
            reader.ActualSheetName = "Data";
            Assert.True(reader.Read());
            Assert.Equal("Category", reader.GetValue(0));
            Assert.True(reader.Read());
            Assert.Equal("A", reader.GetValue(0));
            Assert.Equal(100L, reader.GetValue(1));
            Assert.True(reader.Read());
            Assert.Equal("B", reader.GetValue(0));
            Assert.True(reader.Read());
            Assert.Equal("C", reader.GetValue(0));
            Assert.False(reader.Read());
        }

        using (var reader = new XlsxOrXlsbReadOrEdit())
        {
            reader.Open(xlsbOutput);
            reader.ActualSheetName = "Data";
            Assert.True(reader.Read());
            Assert.True(reader.Read());
            Assert.Equal("A", reader.GetValue(0));
            Assert.Equal(100L, reader.GetValue(1));
            Assert.True(reader.Read());
            Assert.True(reader.Read());
            Assert.Equal("C", reader.GetValue(0));
            Assert.False(reader.Read());
        }
    }

    private static string ResolveRepositoryTestOutput()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        string? firstCandidate = null;
        while (directory is not null)
        {
            string candidate = Path.Combine(directory.FullName, "test_output");
            if (Directory.Exists(candidate))
            {
                firstCandidate ??= candidate;
                if (File.Exists(Path.Combine(candidate, "updater_pivot_source.xlsx")) &&
                    File.Exists(Path.Combine(candidate, "updater_pivot_source.xlsb")))
                    return candidate;
            }
            directory = directory.Parent;
        }
        if (firstCandidate is not null)
            return firstCandidate;
        throw new DirectoryNotFoundException("Cannot find repository test_output directory.");
    }
}
