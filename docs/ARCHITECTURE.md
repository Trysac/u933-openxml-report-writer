# Architecture

`Trysac.OpenxmlReportWriter` generates ECMA-376 5th Edition (ISO/IEC 29500-1:2016) SpreadsheetML packages directly from typed C# records (`ReportTemplate`, `ColumnSpec`, `SheetData`).

## Package Structure & Determinism

`ReportWriter.Write` writes OPC parts into a `ZipArchive` in a fixed order with every entry's `LastWriteTime` pinned to `1980-01-01 00:00:00Z`:

1. `[Content_Types].xml`
2. `_rels/.rels`
3. `docProps/app.xml`
4. `docProps/core.xml`
5. `xl/workbook.xml`
6. `xl/_rels/workbook.xml.rels`
7. `xl/styles.xml`
8. `xl/sharedStrings.xml`
9. `xl/worksheets/sheet1.xml` ... `xl/worksheets/sheetN.xml`

## Validation Pass Before Output

All sheets, columns, formulas, shared strings, and styles are validated and interned in memory before the first ZIP entry is created:

- **Shared Strings (`xl/sharedStrings.xml`)**: Every header and text cell is validated for XML 1.0 characters and the 32,767-character limit, then interned in first-seen order. Leading/trailing whitespace is preserved with `xml:space="preserve"`.
- **Styles (`xl/styles.xml`)**: Built-in number formats `0` (General) and `14` (Date) are combined with custom `numFmt` entries assigned sequentially starting at ID `164` in first-use order. Font `0` is regular Calibri 11pt; Font `1` is bold Calibri 11pt (used for the header row and totals row).
- **Formulas**: `{Key}` placeholders are replaced with the target column's letters and the current 1-based Excel row index. Cached formula values are omitted and `<calcPr fullCalcOnLoad="1"/>` is set on the workbook.
