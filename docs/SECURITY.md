# Security & Threat Model

## Defended

- **Formula Injection**: `CellKind.Text` values are always serialized as Shared Strings (`t="s"`), preventing Excel from interpreting them as formulas. Formula templates (`CellKind.Formula`) are strictly validated against an allowlist of characters (`[A-Z0-9:+\-*/(). ,$]+`) and balanced parentheses.
- **Resource Exhaustion**: Inputs are checked against Excel specifications (max rows, columns, characters, sheet count). Invalid or excessive input triggers a `ReportException` before serialization starts.
- **XML Injection**: Every text input is validated against XML 1.0 character sets. Unpaired surrogates and invalid code points trigger `invalid_character` exceptions.

## Not Defended

- **Memory Exhaustion**: The library constructs the entire workbook in memory; callers must ensure the dataset (`IReadOnlyList<SheetData>`) fits within available heap space.
- **Excel Repair Alerts**: While the output is valid SpreadsheetML, Excel may still prompt for repair if the host system locale or specific Excel version imposes non-standard validation rules not covered by the ECMA-376 specification.