# Hand-written ECMA-376 SpreadsheetML with ZipArchive and XmlWriter

Status: accepted (generated, pending human review)

## Context

Back-office jobs need to produce valid, deterministic .xlsx reports with shared strings, number formats, row formulas, and totals without pulling in large third-party spreadsheet dependencies or producing files that differ on every run.

## Decision

Implement the OPC package and SpreadsheetML parts directly on top of System.IO.Compression.ZipArchive and System.Xml.XmlWriter with a pre-write validation pass, fixed part order, and 1980-01-01 timestamps, using DocumentFormat.OpenXml 3.x exclusively as a dev-scoped test oracle.

## Consequences

The library has zero runtime dependencies, produces byte-identical output across runs and cultures, catches template/data errors early with typed error codes, and is verified against the Open XML SDK validator and LibreOffice Calc in CI.
