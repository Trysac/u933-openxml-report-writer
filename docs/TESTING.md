# Testing Strategy

- **Unit & Acceptance Tests**: Located in `tests/ReportWriterTests.cs`. These verify acceptance criteria including multi-sheet support, empty datasets, shared string reuse, culture invariance (`de-DE`), and formula expansion logic.
- **Schema Validation**: Uses the Open XML SDK (`DocumentFormat.OpenXml`) as a test-only oracle. Generated streams are passed to `OpenXmlValidator` to ensure compliance with the Office 2016 schema.
- **Property-Based Testing**: Located in `tests/TextPropertiesTests.cs`. Uses `FsCheck` to verify invariants for `CellRef.ToA1` and `CellRef.Parse` across the full valid Excel coordinate range.
- **Conformance Tier**: Located in `tests/LibreOfficeConformanceTests.cs`. Runs headless `soffice` (LibreOffice Calc) to verify that generated formulas and totals evaluate correctly in a real-world spreadsheet processor.