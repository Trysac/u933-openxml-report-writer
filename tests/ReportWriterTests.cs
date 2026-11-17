using System.Globalization;
using System.IO.Compression;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Trysac.OpenxmlReportWriter.Tests;

/// <summary>Comprehensive acceptance, oracle, validation and edge-case tests for ReportWriter and CellRef.</summary>
public class ReportWriterTests
{
    private static ReportTemplate CreateSalesTemplate(bool totalsRow = true) =>
        new(
            SheetName: "Sales",
            Columns:
            [
                new ColumnSpec("Product", "Product", CellKind.Text, Width: 20),
                new ColumnSpec("Qty", "Qty", CellKind.Number, NumberFormat: "0", Width: 10, Total: Aggregate.Sum),
                new ColumnSpec("Price", "Price", CellKind.Number, NumberFormat: "#,##0.00", Width: 14, Total: Aggregate.Average),
                new ColumnSpec("Total", "Total", CellKind.Formula, NumberFormat: "#,##0.00", Formula: "{Qty}*{Price}", Width: 16, Total: Aggregate.Sum)
            ],
            TotalsRow: totalsRow,
            FreezeHeader: true,
            AutoFilter: true);

    private static List<IReadOnlyDictionary<string, object?>> CreateSalesRows() =>
    [
        new Dictionary<string, object?> { ["Product"] = "Widget", ["Qty"] = 2, ["Price"] = 10.5m },
        new Dictionary<string, object?> { ["Product"] = "Gadget", ["Qty"] = 5L, ["Price"] = 3.25 },
        new Dictionary<string, object?> { ["Product"] = "Widget", ["Qty"] = 1, ["Price"] = 100.0 }
    ];

    [Fact]
    public void SalesReportWritesFormulasTotalsFreezeAutoFilterAndValidatesWithOpenXmlSdk()
    {
        var template = CreateSalesTemplate(totalsRow: true);
        var rows = CreateSalesRows();
        using var ms = new MemoryStream();

        ReportWriter.Write(ms, [new SheetData(template, rows)]);
        ms.Position = 0;

        using var document = SpreadsheetDocument.Open(ms, isEditable: false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2016);
        var errors = validator.Validate(document).ToList();
        Assert.Empty(errors);

        var workbookPart = document.WorkbookPart!;
        var sst = workbookPart.SharedStringTablePart!.SharedStringTable!;
        var sstItems = sst.Elements<SharedStringItem>().Select(i => i.InnerText).ToList();

        // "Product", "Qty", "Price", "Total", "Widget" (reused twice), "Gadget" => 6 unique, 7 total references
        Assert.Single(sstItems, s => s == "Product");
        Assert.Equal(6, (int)sst.UniqueCount!.Value);
        Assert.Equal(7, (int)sst.Count!.Value);

        var sheet = workbookPart.Workbook!.Sheets!.Elements<Sheet>().Single();
        Assert.Equal("Sales", sheet.Name!.Value);

        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        var worksheet = worksheetPart.Worksheet!;

        var cellsByRef = worksheet.Descendants<Cell>().ToDictionary(c => c.CellReference!.Value!, c => c);
        Assert.Equal("B2*C2", cellsByRef["D2"].CellFormula!.Text);
        Assert.Equal("B4*C4", cellsByRef["D4"].CellFormula!.Text);
        Assert.Equal("SUM(D2:D4)", cellsByRef["D5"].CellFormula!.Text);
        Assert.Equal("SUM(B2:B4)", cellsByRef["B5"].CellFormula!.Text);
        Assert.Equal("AVERAGE(C2:C4)", cellsByRef["C5"].CellFormula!.Text);

        // Frozen header pane
        var pane = worksheet.Descendants<Pane>().Single();
        Assert.Equal(1D, pane.VerticalSplit!.Value);
        Assert.Equal(PaneStateValues.Frozen, pane.State!.Value);

        // AutoFilter covers A1:D4 (header + 3 data rows, excluding totals row)
        var autoFilter = worksheet.Elements<AutoFilter>().Single();
        Assert.Equal("A1:D4", autoFilter.Reference!.Value);

        // Column widths appear as <col> elements
        var cols = worksheet.Elements<Columns>().Single().Elements<Column>().ToList();
        Assert.Equal(4, cols.Count);
        Assert.Equal(20D, cols[0].Width!.Value);
        Assert.Equal(10D, cols[1].Width!.Value);
        Assert.Equal(14D, cols[2].Width!.Value);
        Assert.Equal(16D, cols[3].Width!.Value);
    }

    [Fact]
    public void CellRefConvertsAndParsesExpectedValuesAndRejectsInvalidOnes()
    {
        Assert.Equal("A1", CellRef.ToA1(1, 1));
        Assert.Equal("AA1", CellRef.ToA1(1, 27));
        Assert.Equal("AZ10", CellRef.ToA1(10, 52));
        Assert.Equal("XFD1048576", CellRef.ToA1(1_048_576, 16_384));
        Assert.Equal((12, 28), CellRef.Parse("AB12"));
        Assert.Equal((1_048_576, 16_384), CellRef.Parse("XFD1048576"));

        Assert.Throws<ArgumentNullException>(() => CellRef.Parse(null!));

        string[] invalidRefs = ["", "A", "12", "A0", "A01", "XFE1", "AAAA1", "A1048577", "a1", "A1B", "A-1", "$A$1", "A123456789"];
        foreach (string invalid in invalidRefs)
        {
            var ex = Assert.Throws<ReportException>(() => CellRef.Parse(invalid));
            Assert.Equal("invalid_reference", ex.Code);
        }

        Assert.Equal("invalid_reference", Assert.Throws<ReportException>(() => CellRef.ToA1(0, 1)).Code);
        Assert.Equal("invalid_reference", Assert.Throws<ReportException>(() => CellRef.ToA1(1_048_577, 1)).Code);
        Assert.Equal("invalid_reference", Assert.Throws<ReportException>(() => CellRef.ToA1(1, 0)).Code);
        Assert.Equal("invalid_reference", Assert.Throws<ReportException>(() => CellRef.ToA1(1, 16_385)).Code);
    }

    [Fact]
    public void DateSerialAndNumberFormattingAreCultureIndependentUnderDeDe()
    {
        var template = new ReportTemplate(
            "Metrics",
            [
                new ColumnSpec("When", "When", CellKind.Date),
                new ColumnSpec("Timestamp", "Timestamp", CellKind.Date, NumberFormat: "yyyy-mm-dd hh:mm"),
                new ColumnSpec("Amount", "Amount", CellKind.Number),
                new ColumnSpec("Active", "Active", CellKind.Boolean, Total: Aggregate.Count),
                new ColumnSpec("LiteralFormulaText", "LiteralFormulaText", CellKind.Text)
            ],
            TotalsRow: true);

        var rows = new List<IReadOnlyDictionary<string, object?>>
        {
            new Dictionary<string, object?>
            {
                ["When"] = new DateOnly(2026, 1, 15),
                ["Timestamp"] = new DateTime(2026, 1, 15, 12, 0, 0, DateTimeKind.Utc),
                ["Amount"] = 1234.5,
                ["Active"] = true,
                ["LiteralFormulaText"] = "=SUM(A1:A2)"
            },
            new Dictionary<string, object?>
            {
                ["When"] = new DateTime(2026, 1, 15, 0, 0, 0, DateTimeKind.Utc),
                ["Timestamp"] = null,
                ["Amount"] = null,
                ["Active"] = false,
                ["LiteralFormulaText"] = "  leading and trailing  "
            }
        };

        var previousCulture = CultureInfo.CurrentCulture;
        using var ms = new MemoryStream();
        try
        {
            CultureInfo.CurrentCulture = new CultureInfo("de-DE");
            ReportWriter.Write(ms, [new SheetData(template, rows)]);
        }
        finally
        {
            CultureInfo.CurrentCulture = previousCulture;
        }

        ms.Position = 0;
        using var document = SpreadsheetDocument.Open(ms, isEditable: false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2016);
        Assert.Empty(validator.Validate(document));

        var wbPart = document.WorkbookPart!;
        var wsPart = wbPart.WorksheetParts.Single();
        var cells = wsPart.Worksheet!.Descendants<Cell>().ToDictionary(c => c.CellReference!.Value!, c => c);

        // 2026-01-15 is serial 46037
        Assert.Equal("46037", cells["A2"].CellValue!.Text);
        Assert.Equal("46037", cells["A3"].CellValue!.Text);
        Assert.Equal("46037.5", cells["B2"].CellValue!.Text);
        Assert.False(cells.ContainsKey("B3")); // null leaves cell out
        Assert.False(cells.ContainsKey("C3"));

        // Number 1234.5 written with invariant dot even under de-DE
        Assert.Equal("1234.5", cells["C2"].CellValue!.Text);

        // Boolean written as t="b" with 1 and 0
        Assert.Equal(CellValues.Boolean, cells["D2"].DataType!.Value);
        Assert.Equal("1", cells["D2"].CellValue!.Text);
        Assert.Equal(CellValues.Boolean, cells["D3"].DataType!.Value);
        Assert.Equal("0", cells["D3"].CellValue!.Text);
        Assert.Equal("COUNTA(D2:D3)", cells["D4"].CellFormula!.Text);

        // Text starting with "=" is stored as shared string, never as a formula
        Assert.Equal(CellValues.SharedString, cells["E2"].DataType!.Value);
        Assert.Null(cells["E2"].CellFormula);
        int sstIndex = int.Parse(cells["E2"].CellValue!.Text, CultureInfo.InvariantCulture);
        var sstItem = wbPart.SharedStringTablePart!.SharedStringTable!.Elements<SharedStringItem>().ElementAt(sstIndex);
        Assert.Equal("=SUM(A1:A2)", sstItem.InnerText);

        // Check that A2 has the built-in date style (numFmtId = 14)
        uint styleIdx = cells["A2"].StyleIndex!.Value;
        var xf = (CellFormat)wbPart.WorkbookStylesPart!.Stylesheet!.CellFormats!.ElementAt((int)styleIdx);
        Assert.Equal(14U, xf.NumberFormatId!.Value);
    }

    [Fact]
    public void WritingSameReportTwiceProducesByteIdenticalStreamsAndFixedZipMetadata()
    {
        var sheet = new SheetData(CreateSalesTemplate(), CreateSalesRows());
        using var ms1 = new MemoryStream();
        using var ms2 = new MemoryStream();

        ReportWriter.Write(ms1, [sheet]);
        ReportWriter.Write(ms2, [sheet]);

        byte[] bytes1 = ms1.ToArray();
        byte[] bytes2 = ms2.ToArray();
        Assert.Equal(bytes1, bytes2);

        ms1.Position = 0;
        using var zip = new ZipArchive(ms1, ZipArchiveMode.Read, leaveOpen: true);
        string[] expectedEntries =
        [
            "[Content_Types].xml",
            "_rels/.rels",
            "docProps/app.xml",
            "docProps/core.xml",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "xl/styles.xml",
            "xl/sharedStrings.xml",
            "xl/worksheets/sheet1.xml"
        ];
        Assert.Equal(expectedEntries, zip.Entries.Select(e => e.FullName).ToArray());
        foreach (var entry in zip.Entries)
        {
            Assert.Equal(new DateTime(1980, 1, 1, 0, 0, 0), entry.LastWriteTime.DateTime);
        }
    }

    [Fact]
    public void SupportsTwoSheetsEmptyRowsAndFormulaReferencingLaterColumn()
    {
        var sheet1Template = new ReportTemplate(
            "EmptySheet",
            [
                new ColumnSpec("Code", "Code", CellKind.Text, Total: Aggregate.Count),
                new ColumnSpec("Value", "Value", CellKind.Number, Total: Aggregate.Sum)
            ],
            TotalsRow: true,
            FreezeHeader: false,
            AutoFilter: true);

        var sheet2Template = new ReportTemplate(
            "ForwardFormula",
            [
                new ColumnSpec("Net", "Net", CellKind.Formula, Formula: "{Gross}-{Tax}"),
                new ColumnSpec("Gross", "Gross", CellKind.Number),
                new ColumnSpec("Tax", "Tax", CellKind.Number)
            ],
            TotalsRow: false,
            FreezeHeader: true,
            AutoFilter: false);

        var sheet2Rows = new List<IReadOnlyDictionary<string, object?>>
        {
            new Dictionary<string, object?> { ["Gross"] = 100, ["Tax"] = 15 }
        };

        using var ms = new MemoryStream();
        ReportWriter.Write(ms, [
            new SheetData(sheet1Template, []),
            new SheetData(sheet2Template, sheet2Rows)
        ]);

        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, isEditable: false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2016);
        Assert.Empty(validator.Validate(doc));

        var sheets = doc.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().ToList();
        Assert.Equal(2, sheets.Count);

        // Sheet 1 has 0 data rows, so only row 1 (header) is written and totals row is omitted
        var ws1 = ((WorksheetPart)doc.WorkbookPart.GetPartById(sheets[0].Id!.Value!)).Worksheet!;
        var rows1 = ws1.Descendants<Row>().ToList();
        Assert.Single(rows1);
        Assert.Equal("A1:B1", ws1.Elements<AutoFilter>().Single().Reference!.Value);

        // Sheet 2 column A references later columns B and C
        var ws2 = ((WorksheetPart)doc.WorkbookPart.GetPartById(sheets[1].Id!.Value!)).Worksheet!;
        var cellA2 = ws2.Descendants<Cell>().First(c => c.CellReference!.Value == "A2");
        Assert.Equal("B2-C2", cellA2.CellFormula!.Text);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("ThisSheetNameIs32CharactersLong!")]
    [InlineData("Bad[Name")]
    [InlineData("Bad]Name")]
    [InlineData("Bad:Name")]
    [InlineData("Bad*Name")]
    [InlineData("Bad?Name")]
    [InlineData("Bad/Name")]
    [InlineData("Bad\\Name")]
    [InlineData("'LeadingQuote")]
    [InlineData("TrailingQuote'")]
    [InlineData("Bad\u0001Char")]
    public void InvalidSheetNameThrowsExpectedCode(string invalidName)
    {
        var template = new ReportTemplate(invalidName, [new ColumnSpec("A", "A", CellKind.Text)]);
        using var ms = new MemoryStream();
        var ex = Assert.Throws<ReportException>(() => ReportWriter.Write(ms, [new SheetData(template, [])]));
        Assert.Equal("invalid_sheet_name", ex.Code);
    }

    [Fact]
    public void DuplicateSheetNameDifferingOnlyInCaseThrowsDuplicateSheetName()
    {
        var t1 = new ReportTemplate("Sales", [new ColumnSpec("A", "A", CellKind.Text)]);
        var t2 = new ReportTemplate("sAlEs", [new ColumnSpec("A", "A", CellKind.Text)]);
        using var ms = new MemoryStream();
        var ex = Assert.Throws<ReportException>(() => ReportWriter.Write(ms, [new SheetData(t1, []), new SheetData(t2, [])]));
        Assert.Equal("duplicate_sheet_name", ex.Code);
    }

    [Fact]
    public void CellTextOf32768CharactersThrowsCellTooLong()
    {
        var template = new ReportTemplate("Sheet1", [new ColumnSpec("A", "A", CellKind.Text)]);
        string longText = new('x', 32_768);
        var rows = new List<IReadOnlyDictionary<string, object?>>
        {
            new Dictionary<string, object?> { ["A"] = longText }
        };
        using var ms = new MemoryStream();
        var ex = Assert.Throws<ReportException>(() => ReportWriter.Write(ms, [new SheetData(template, rows)]));
        Assert.Equal("cell_too_long", ex.Code);
        Assert.True(ex.Message.Length <= 200);
    }

    [Fact]
    public void TextWithControlCharacterU0001ThrowsInvalidCharacter()
    {
        var template = new ReportTemplate("Sheet1", [new ColumnSpec("A", "A", CellKind.Text)]);
        var rows = new List<IReadOnlyDictionary<string, object?>>
        {
            new Dictionary<string, object?> { ["A"] = "Hello\u0001World" }
        };
        using var ms = new MemoryStream();
        var ex = Assert.Throws<ReportException>(() => ReportWriter.Write(ms, [new SheetData(template, rows)]));
        Assert.Equal("invalid_character", ex.Code);
    }

    [Fact]
    public void TooManyRowsThrowsTooManyRows()
    {
        var template = new ReportTemplate("Sheet1", [new ColumnSpec("A", "A", CellKind.Text)]);
        var emptyRow = new Dictionary<string, object?>();
        var rows = Enumerable.Repeat<IReadOnlyDictionary<string, object?>>(emptyRow, 1_048_576).ToList();
        using var ms = new MemoryStream();
        var ex = Assert.Throws<ReportException>(() => ReportWriter.Write(ms, [new SheetData(template, rows)]));
        Assert.Equal("too_many_rows", ex.Code);
    }

    [Fact]
    public void TooManyColumnsThrowsTooManyColumns()
    {
        var cols = Enumerable.Range(1, 16_385)
            .Select(i => new ColumnSpec($"H{i}", $"K{i}", CellKind.Text))
            .ToList();
        var template = new ReportTemplate("Sheet1", cols);
        using var ms = new MemoryStream();
        var ex = Assert.Throws<ReportException>(() => ReportWriter.Write(ms, [new SheetData(template, [])]));
        Assert.Equal("too_many_columns", ex.Code);
    }

    [Fact]
    public void FormulaReferencingUnknownColumnThrowsUnknownColumn()
    {
        var template = new ReportTemplate(
            "Sheet1",
            [
                new ColumnSpec("Qty", "Qty", CellKind.Number),
                new ColumnSpec("Bad", "Bad", CellKind.Formula, Formula: "{Unknown}*2")
            ]);
        using var ms = new MemoryStream();
        var ex = Assert.Throws<ReportException>(() => ReportWriter.Write(ms, [new SheetData(template, [])]));
        Assert.Equal("unknown_column", ex.Code);
    }

    [Theory]
    [InlineData("={Qty}*2")]
    [InlineData("SUM({Qty};10)")]
    [InlineData("sum({Qty})")]
    [InlineData("{Qty")
    ]
    [InlineData("Qty}")]
    [InlineData("{}")]
    [InlineData("({Qty}+1")]
    [InlineData("{Qty}+1)")]
    [InlineData("")]
    [InlineData("   ")]
    public void InvalidFormulaSyntaxThrowsInvalidFormula(string badFormula)
    {
        var template = new ReportTemplate(
            "Sheet1",
            [
                new ColumnSpec("Qty", "Qty", CellKind.Number),
                new ColumnSpec("Calc", "Calc", CellKind.Formula, Formula: badFormula)
            ]);
        using var ms = new MemoryStream();
        var ex = Assert.Throws<ReportException>(() => ReportWriter.Write(ms, [new SheetData(template, [])]));
        Assert.Equal("invalid_formula", ex.Code);
    }

    [Fact]
    public void InvalidValuesAndEmptySheetsThrowExpectedCodes()
    {
        using var ms = new MemoryStream();
        Assert.Equal("no_sheets", Assert.Throws<ReportException>(() => ReportWriter.Write(ms, [])).Code);

        var numTemplate = new ReportTemplate("S", [new ColumnSpec("N", "N", CellKind.Number)]);
        Assert.Equal(
            "invalid_value",
            Assert.Throws<ReportException>(() => ReportWriter.Write(ms, [new SheetData(numTemplate, [new Dictionary<string, object?> { ["N"] = double.NaN }])])).Code);
        Assert.Equal(
            "invalid_value",
            Assert.Throws<ReportException>(() => ReportWriter.Write(ms, [new SheetData(numTemplate, [new Dictionary<string, object?> { ["N"] = double.PositiveInfinity }])])).Code);
        Assert.Equal(
            "invalid_value",
            Assert.Throws<ReportException>(() => ReportWriter.Write(ms, [new SheetData(numTemplate, [new Dictionary<string, object?> { ["N"] = "not a number" }])])).Code);

        var dateTemplate = new ReportTemplate("S", [new ColumnSpec("D", "D", CellKind.Date)]);
        Assert.Equal(
            "invalid_value",
            Assert.Throws<ReportException>(() => ReportWriter.Write(ms, [new SheetData(dateTemplate, [new Dictionary<string, object?> { ["D"] = 12345 }])])).Code);

        var boolTemplate = new ReportTemplate("S", [new ColumnSpec("B", "B", CellKind.Boolean)]);
        Assert.Equal(
            "invalid_value",
            Assert.Throws<ReportException>(() => ReportWriter.Write(ms, [new SheetData(boolTemplate, [new Dictionary<string, object?> { ["B"] = "true" }])])).Code);

        var textTemplate = new ReportTemplate("S", [new ColumnSpec("T", "T", CellKind.Text)]);
        Assert.Equal(
            "invalid_value",
            Assert.Throws<ReportException>(() => ReportWriter.Write(ms, [new SheetData(textTemplate, [new Dictionary<string, object?> { ["T"] = 42 }])])).Code);
    }

    [Fact]
    public void ReportExceptionConstructorsAndTruncationWorkAsExpected()
    {
        var defaultEx = new ReportException();
        Assert.Equal("invalid_value", defaultEx.Code);

        var msgEx = new ReportException("short");
        Assert.Equal("short", msgEx.Message);

        var inner = new InvalidOperationException("inner");
        var wrapped = new ReportException("wrapped", inner);
        Assert.Same(inner, wrapped.InnerException);

        var codedInner = new ReportException("custom_code", "msg", inner);
        Assert.Equal("custom_code", codedInner.Code);
        Assert.Same(inner, codedInner.InnerException);

        string veryLongMessage = new('a', 300);
        var truncated = new ReportException("cell_too_long", veryLongMessage);
        Assert.Equal(200, truncated.Message.Length);
        Assert.EndsWith("...", truncated.Message, StringComparison.Ordinal);
    }
}
