using SpreadSheetTasks;
using System.Text.Json;

namespace Tests;

[Collection("Sequential")]
public class UpdaterComArtifactTests
{
    [Fact]
    public void CreateUpdatedFilesAndManifestForExcelComValidation()
    {
        string outputDirectory = Path.Combine(Directory.GetCurrentDirectory(), "test_output");
        Directory.CreateDirectory(outputDirectory);

        var manifestFiles = new List<object>();
        foreach (string extension in new[] { ".xlsx", ".xlsb" })
        {
            string source = Path.Combine(outputDirectory, "updater_com_source" + extension);
            string output = Path.Combine(outputDirectory, "updater_com" + extension);

            using (var writer = ExcelWriter.CreateWriter(source))
            {
                writer.AddSheet("Data");
                writer.WriteSheet(new object?[][]
                {
                    ["old", 1],
                    ["old2", 2]
                }, ["Name", "Amount"]);
                writer.AddSheet("Other");
                writer.WriteSheet(new object?[][] { ["preserved"] }, headers: false);
            }

            if (extension == ".xlsx")
            {
                using var updater = new XlsxUpdater(source);
                updater.ReplaceSheetData("Data", new object?[][]
                {
                    ["COM", 42.5],
                    ["validated", 7]
                }, new ReplaceSheetDataOptions { Headers = ["Name", "Amount"] });
                updater.Save(output);
            }
            else
            {
                using var updater = new XlsbUpdater(source);
                updater.ReplaceSheetData("Data", new object?[][]
                {
                    ["COM", 42.5],
                    ["validated", 7]
                }, new ReplaceSheetDataOptions { Headers = ["Name", "Amount"] });
                updater.Save(output);
            }

            File.Delete(source);
            manifestFiles.Add(new
            {
                path = output,
                sheets = new[] { "Data", "Other" },
                cells = new object[]
                {
                    new { sheet = "Data", address = "A1", value = "Name" },
                    new { sheet = "Data", address = "A2", value = "COM" },
                    new { sheet = "Data", address = "B2", value = 42.5 },
                    new { sheet = "Data", address = "A3", value = "validated" },
                    new { sheet = "Other", address = "A1", value = "preserved" }
                }
            });
        }

        string manifestPath = Path.Combine(outputDirectory, "updater_com_manifest.json");
        File.WriteAllText(manifestPath, JsonSerializer.Serialize(new { files = manifestFiles }, new JsonSerializerOptions
        {
            WriteIndented = true
        }));
    }
}
