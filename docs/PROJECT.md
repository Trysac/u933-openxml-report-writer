# Project record

## Objective

Write Excel xlsx reports from typed templates with hand-written SpreadsheetML: shared strings, number formats, row formulas, totals and byte-identical output.

## Scope

- Typed `ReportTemplate` definitions (columns, aggregate functions, cell types, formulas).
- Deterministic package generation (fixed part ordering, static 1980 timestamps).
- Validation of SpreadsheetML rules (limits, character sets, formula syntax).
- Native integration with .NET stream I/O via `System.IO.Compression` and `System.Xml`.

## Out-of-scope

- Reading or modifying existing `.xlsx` files.
- Charts, images, conditional formatting, merged cells, rich text.
- Encryption or signing of files.
- Streaming support for datasets exceeding available system memory.

## Factory run u933 (measured from API usage)

| Date (UTC) | Phase | Model | Batch | Input | Output | Cache write | Cache read | Cost (USD) |
|---|---|---|---|---|---|---|---|---|
| 2026-11-17 | generation | google/gemini-4-argon | yes | 6023 | 25577 | 0 | 0 | 0.5356 |
| 2026-11-17 | fix-1 | google/gemini-3.8-flash | no | 39754 | 376 | 0 | 159844 | 0.0432 |
| 2026-11-17 | audit | google/gemini-3.8-flash | no | 151877 | 893 | 0 | 450833 | 0.1511 |
| 2026-11-17 | docs | google/gemini-3.1-flash-lite | no | 21715 | 1390 | 0 | 0 | 0.0075 |

Total: $0.7374. Prices: vertex line, as measured by its adapter. Providers serve the current version of each model id, so the date tells which version answered; versions, launch stages and price sources per line are in the factory's model register (docs/MODELS.md).
