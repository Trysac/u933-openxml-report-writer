# Trysac.OpenxmlReportWriter

`Trysac.OpenxmlReportWriter` is a lightweight, zero-dependency library for generating deterministic Excel `.xlsx` reports from typed C# templates. It produces valid ECMA-376 SpreadsheetML packages by manually writing zip entries, ensuring byte-identical output across runs.

## Why not closedxml/EPPlus?

Unlike `ClosedXML` or `EPPlus`, `Trysac.OpenxmlReportWriter` has no runtime dependencies, uses only the .NET base class library, and guarantees byte-identical output with deterministic metadata. It focuses on template-level validation of formulas and structure rather than arbitrary editing of existing files.

## Usage

```csharp
using Trysac.OpenxmlReportWriter;

var template = new ReportTemplate("Sales", [
    new ColumnSpec("Product", "Product", CellKind.Text, Width: 20),
    new ColumnSpec("Qty", "Qty", CellKind.Number, NumberFormat: "0", Total: Aggregate.Sum),
    new ColumnSpec("Price", "Price", CellKind.Number, NumberFormat: "#,##0.00"),
    new ColumnSpec("Total", "Total", CellKind.Formula, NumberFormat: "#,##0.00", Formula: "{Qty}*{Price}", Total: Aggregate.Sum)
], TotalsRow: true);
using var output = new MemoryStream();
ReportWriter.Write(output, [new SheetData(template, [new Dictionary<string, object?> { ["Product"] = "Widget", ["Qty"] = 3, ["Price"] = 19.95m }])]);
```

## API Summary

### Core Classes
- `ReportWriter.Write(Stream, IReadOnlyList<SheetData>)`: Generates the `.xlsx` package.
- `ReportTemplate`: Defines worksheet settings (name, totals, freeze headers, auto-filter).
- `ColumnSpec`: Defines column properties (key, kind, format, width, formula).
- `SheetData`: Binds a template to a collection of row data.
- `CellRef`: Helper for converting 1-based coordinates to A1 notation (`ToA1`, `Parse`).

### Limits & Options
- **Excel Limits**: Supports up to 255 sheets, 1,048,576 rows, 16,384 columns, 32,767 characters per cell, and 31-character sheet names.
- **Error Handling**: Throws `ReportException` for validation errors (e.g., `invalid_sheet_name`, `invalid_formula`, `too_many_rows`, `unknown_column`).

See [docs/SECURITY.md](docs/SECURITY.md) for details on threat modeling.