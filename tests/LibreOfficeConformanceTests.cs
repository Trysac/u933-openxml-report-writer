using System.Diagnostics;
using Xunit;

namespace Trysac.OpenxmlReportWriter.Tests;

/// <summary>CI conformance test that verifies generated workbooks evaluate formulas cleanly in headless LibreOffice Calc.</summary>
public class LibreOfficeConformanceTests
{
    [Fact]
    [Trait("Category", "conformance")]
    public void LibreOfficeCalcConvertsReportAndEvaluatesFormulas()
    {
        string? sofficePath = FindSofficeExecutable();
        if (sofficePath is null)
        {
            return;
        }

        string tempDir = Path.Combine(Path.GetTempPath(), "openxml-report-writer-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        try
        {
            string xlsxPath = Path.Combine(tempDir, "sales.xlsx");
            string csvPath = Path.Combine(tempDir, "sales.csv");

            var template = new ReportTemplate(
                "Sales",
                [
                    new ColumnSpec("Product", "Product", CellKind.Text),
                    new ColumnSpec("Qty", "Qty", CellKind.Number, NumberFormat: "0", Total: Aggregate.Sum),
                    new ColumnSpec("Price", "Price", CellKind.Number, NumberFormat: "0.00"),
                    new ColumnSpec("Total", "Total", CellKind.Formula, NumberFormat: "0.00", Formula: "{Qty}*{Price}", Total: Aggregate.Sum)
                ],
                TotalsRow: true);

            var rows = new List<IReadOnlyDictionary<string, object?>>
            {
                new Dictionary<string, object?> { ["Product"] = "Alpha", ["Qty"] = 2, ["Price"] = 10.5m },
                new Dictionary<string, object?> { ["Product"] = "Beta", ["Qty"] = 3, ["Price"] = 4.0m },
                new Dictionary<string, object?> { ["Product"] = "Gamma", ["Qty"] = 5, ["Price"] = 2.0m }
            };

            using (var fs = File.Create(xlsxPath))
            {
                ReportWriter.Write(fs, [new SheetData(template, rows)]);
            }

            var psi = new ProcessStartInfo(sofficePath)
            {
                ArgumentList =
                {
                    "--headless",
                    "--convert-to",
                    "csv",
                    "--outdir",
                    tempDir,
                    xlsxPath
                },
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false
            };

            using var proc = Process.Start(psi)!;
            proc.WaitForExit(30_000);
            Assert.Equal(0, proc.ExitCode);
            Assert.True(File.Exists(csvPath));

            string[] lines = File.ReadAllLines(csvPath);
            Assert.Equal(5, lines.Length);
            Assert.Equal("Product,Qty,Price,Total", lines[0]);
            Assert.Equal("Alpha,2,10.50,21.00", lines[1]);
            Assert.Equal("Beta,3,4.00,12.00", lines[2]);
            Assert.Equal("Gamma,5,2.00,10.00", lines[3]);
            Assert.Equal(",10,,43.00", lines[4]);
        }
        finally
        {
            if (Directory.Exists(tempDir))
            {
                Directory.Delete(tempDir, recursive: true);
            }
        }
    }

    private static string? FindSofficeExecutable()
    {
        string pathEnv = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        foreach (string dir in pathEnv.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            string candidate = Path.Combine(dir, "soffice");
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }
        return null;
    }
}
