using System.Buffers;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace Trysac.OpenxmlReportWriter;

/// <summary>Writes deterministic SpreadsheetML (.xlsx) workbooks from typed templates and row dictionaries.</summary>
public static class ReportWriter
{
    /// <summary>Maximum number of worksheets in a workbook.</summary>
    public const int MaxSheets = 255;

    /// <summary>Maximum length of a worksheet name in characters.</summary>
    public const int MaxSheetNameLength = 31;

    /// <summary>Maximum number of characters in a single cell.</summary>
    public const int MaxCellTextLength = 32_767;

    /// <summary>Maximum length of a formula expression.</summary>
    public const int MaxFormulaLength = 8_192;

    private static readonly DateTimeOffset DeterministicZipTimestamp = new(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);
    private const string DeterministicIsoDate = "1980-01-01T00:00:00Z";

    private const string SpreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelationshipsNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelsNs = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string ContentTypesNs = "http://schemas.openxmlformats.org/package/2006/content-types";

    private static readonly SearchValues<char> ForbiddenSheetNameChars = SearchValues.Create(['[', ']', ':', '*', '?', '/', '\\']);
    private static readonly DateOnly ExcelEpoch = new(1899, 12, 30);
    private static readonly UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);

    /// <summary>Writes an Excel workbook containing <paramref name="sheets"/> to <paramref name="output"/>.</summary>
    /// <param name="output">A writable destination stream.</param>
    /// <param name="sheets">The ordered list of sheets to write (1 to 255).</param>
    /// <exception cref="ArgumentNullException"><paramref name="output"/> or <paramref name="sheets"/> is null.</exception>
    /// <exception cref="ReportException">Any validation limit, formula, sheet name or cell value rule is violated.</exception>
    public static void Write(Stream output, IReadOnlyList<SheetData> sheets)
    {
        ArgumentNullException.ThrowIfNull(output);
        ArgumentNullException.ThrowIfNull(sheets);

        if (sheets.Count == 0)
        {
            throw new ReportException("no_sheets", "At least one sheet must be provided.");
        }
        if (sheets.Count > MaxSheets)
        {
            throw new ReportException("invalid_value", $"Workbook cannot contain more than {MaxSheets} sheets.");
        }

        var preparedSheets = ValidateAndPrepare(sheets, out var stylesTable, out var sharedStrings);

        using var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true);

        WriteEntry(archive, "[Content_Types].xml", writer => WriteContentTypes(writer, preparedSheets.Count));
        WriteEntry(archive, "_rels/.rels", WriteRootRels);
        WriteEntry(archive, "docProps/app.xml", writer => WriteAppProps(writer, preparedSheets));
        WriteEntry(archive, "docProps/core.xml", WriteCoreProps);
        WriteEntry(archive, "xl/workbook.xml", writer => WriteWorkbook(writer, preparedSheets));
        WriteEntry(archive, "xl/_rels/workbook.xml.rels", writer => WriteWorkbookRels(writer, preparedSheets.Count));
        WriteEntry(archive, "xl/styles.xml", writer => WriteStyles(writer, stylesTable));
        WriteEntry(archive, "xl/sharedStrings.xml", writer => WriteSharedStrings(writer, sharedStrings));

        for (int i = 0; i < preparedSheets.Count; i++)
        {
            int sheetIndex = i + 1;
            var sheet = preparedSheets[i];
            WriteEntry(archive, $"xl/worksheets/sheet{sheetIndex}.xml", writer => WriteWorksheet(writer, sheet));
        }
    }

    private static void WriteEntry(ZipArchive archive, string entryName, Action<XmlWriter> writeXml)
    {
        var entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
        entry.LastWriteTime = DeterministicZipTimestamp;
        using var stream = entry.Open();
        var settings = new XmlWriterSettings
        {
            Encoding = Utf8NoBom,
            Indent = false,
            NewLineHandling = NewLineHandling.None,
            OmitXmlDeclaration = false,
            CloseOutput = false
        };
        using var writer = XmlWriter.Create(stream, settings);
        writeXml(writer);
        writer.Flush();
    }

    private static List<PreparedSheet> ValidateAndPrepare(
        IReadOnlyList<SheetData> sheets,
        out StylesBuilder styles,
        out SharedStringsBuilder sharedStrings)
    {
        styles = new StylesBuilder();
        sharedStrings = new SharedStringsBuilder();
        var seenNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var prepared = new List<PreparedSheet>(sheets.Count);

        foreach (var sheetData in sheets)
        {
            if (sheetData is null || sheetData.Template is null || sheetData.Rows is null)
            {
                throw new ReportException("invalid_value", "SheetData, Template, and Rows must not be null.");
            }

            var template = sheetData.Template;
            ValidateSheetName(template.SheetName);
            if (!seenNames.Add(template.SheetName))
            {
                throw new ReportException("duplicate_sheet_name", $"Duplicate sheet name '{template.SheetName}'.");
            }

            if (template.Columns is null || template.Columns.Count == 0)
            {
                throw new ReportException("invalid_value", "Sheet template must define at least one column.");
            }
            if (template.Columns.Count > CellRef.MaxColumn)
            {
                throw new ReportException("too_many_columns", $"Sheet exceeds maximum of {CellRef.MaxColumn} columns.");
            }

            int dataRowCount = sheetData.Rows.Count;
            bool hasTotalsRow = template.TotalsRow && dataRowCount > 0;
            long totalRows = 1L + dataRowCount + (hasTotalsRow ? 1L : 0L);
            if (totalRows > CellRef.MaxRow)
            {
                throw new ReportException("too_many_rows", $"Sheet total rows ({totalRows}) exceed {CellRef.MaxRow}.");
            }

            var keyToColLetter = new Dictionary<string, string>(StringComparer.Ordinal);
            for (int c = 0; c < template.Columns.Count; c++)
            {
                var col = template.Columns[c] ?? throw new ReportException("invalid_value", "Column specification cannot be null.");
                if (string.IsNullOrEmpty(col.Key))
                {
                    throw new ReportException("invalid_value", "Column Key cannot be null or empty.");
                }
                ValidateTextValue(col.Header);
                ValidateTextValue(col.Key);
                if (!double.IsFinite(col.Width) || col.Width <= 0 || col.Width > 255)
                {
                    throw new ReportException("invalid_value", "Column Width must be a finite number between 0 and 255.");
                }
                if (!keyToColLetter.TryAdd(col.Key, CellRef.ColumnToLetters(c + 1)))
                {
                    throw new ReportException("invalid_value", $"Duplicate column key '{col.Key}'.");
                }
            }

            var preparedColumns = new List<PreparedColumn>(template.Columns.Count);
            for (int c = 0; c < template.Columns.Count; c++)
            {
                var col = template.Columns[c];
                int dataStyleIndex = styles.GetDataStyleIndex(col.Kind, col.NumberFormat);
                int totalsStyleIndex = styles.GetTotalsStyleIndex(col.Kind, col.NumberFormat, col.Total);

                string? sampleFormula = null;
                if (col.Kind == CellKind.Formula)
                {
                    if (string.IsNullOrWhiteSpace(col.Formula))
                    {
                        throw new ReportException("invalid_formula", $"Formula column '{col.Key}' must specify a non-empty formula.");
                    }
                    sampleFormula = ExpandAndValidateFormula(col.Formula, keyToColLetter, 2);
                }
                else if (col.Formula is not null)
                {
                    throw new ReportException("invalid_formula", $"Non-formula column '{col.Key}' must not specify a formula.");
                }

                int headerStringId = sharedStrings.Intern(col.Header);
                preparedColumns.Add(new PreparedColumn(
                    col,
                    c + 1,
                    CellRef.ColumnToLetters(c + 1),
                    headerStringId,
                    dataStyleIndex,
                    totalsStyleIndex,
                    sampleFormula));
            }

            var preparedRows = new List<PreparedCell[]>(dataRowCount);
            for (int r = 0; r < dataRowCount; r++)
            {
                var rowDict = sheetData.Rows[r] ?? throw new ReportException("invalid_value", "Row dictionary cannot be null.");
                int excelRowNumber = r + 2;
                var rowCells = new PreparedCell[preparedColumns.Count];

                for (int c = 0; c < preparedColumns.Count; c++)
                {
                    var pCol = preparedColumns[c];
                    var spec = pCol.Spec;

                    if (spec.Kind == CellKind.Formula)
                    {
                        string expanded = ExpandAndValidateFormula(spec.Formula!, keyToColLetter, excelRowNumber);
                        rowCells[c] = PreparedCell.ForFormula(expanded, pCol.DataStyleIndex);
                        continue;
                    }

                    if (!rowDict.TryGetValue(spec.Key, out object? rawValue) || rawValue is null)
                    {
                        rowCells[c] = PreparedCell.Empty;
                        continue;
                    }

                    rowCells[c] = ConvertCellValue(spec.Kind, rawValue, pCol.DataStyleIndex, sharedStrings);
                }

                preparedRows.Add(rowCells);
            }

            prepared.Add(new PreparedSheet(template, preparedColumns, preparedRows, hasTotalsRow));
        }

        return prepared;
    }

    private static PreparedCell ConvertCellValue(
        CellKind kind,
        object rawValue,
        int styleIndex,
        SharedStringsBuilder sharedStrings)
    {
        switch (kind)
        {
            case CellKind.Text:
                if (rawValue is not string text)
                {
                    throw new ReportException("invalid_value", "Text column expects a string value.");
                }
                ValidateTextValue(text);
                int sstId = sharedStrings.Intern(text);
                return PreparedCell.ForSharedString(sstId, styleIndex);

            case CellKind.Number:
                return PreparedCell.ForNumber(FormatNumberValue(rawValue), styleIndex);

            case CellKind.Date:
                return PreparedCell.ForNumber(FormatDateValue(rawValue), styleIndex);

            case CellKind.Boolean:
                if (rawValue is not bool flag)
                {
                    throw new ReportException("invalid_value", "Boolean column expects a bool value.");
                }
                return PreparedCell.ForBoolean(flag, styleIndex);

            default:
                throw new ReportException("invalid_value", $"Unsupported cell kind '{kind}'.");
        }
    }

    private static string FormatNumberValue(object rawValue)
    {
        switch (rawValue)
        {
            case int i:
                return i.ToString(CultureInfo.InvariantCulture);
            case long l:
                return l.ToString(CultureInfo.InvariantCulture);
            case double d:
                if (double.IsNaN(d) || double.IsInfinity(d))
                {
                    throw new ReportException("invalid_value", "Numeric value cannot be NaN or Infinity.");
                }
                return d.ToString("G17", CultureInfo.InvariantCulture);
            case decimal m:
                return m.ToString(CultureInfo.InvariantCulture);
            default:
                throw new ReportException("invalid_value", $"Number column requires int, long, double, or decimal, but received {rawValue.GetType().Name}.");
        }
    }

    private static string FormatDateValue(object rawValue)
    {
        switch (rawValue)
        {
            case DateOnly dateOnly:
                {
                    int serial = dateOnly.DayNumber - ExcelEpoch.DayNumber;
                    return serial.ToString(CultureInfo.InvariantCulture);
                }
            case DateTime dateTime:
                {
                    var datePart = DateOnly.FromDateTime(dateTime);
                    int days = datePart.DayNumber - ExcelEpoch.DayNumber;
                    double fraction = dateTime.TimeOfDay.TotalDays;
                    if (fraction == 0d)
                    {
                        return days.ToString(CultureInfo.InvariantCulture);
                    }
                    double serial = days + fraction;
                    return serial.ToString("G17", CultureInfo.InvariantCulture);
                }
            default:
                throw new ReportException("invalid_value", $"Date column requires DateOnly or DateTime, but received {rawValue.GetType().Name}.");
        }
    }

    private static void ValidateSheetName(string sheetName)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
        {
            throw new ReportException("invalid_sheet_name", "Sheet name cannot be null, empty, or whitespace.");
        }
        if (sheetName.Length > MaxSheetNameLength)
        {
            throw new ReportException("invalid_sheet_name", $"Sheet name exceeds {MaxSheetNameLength} characters.");
        }
        if (sheetName[0] == '\'' || sheetName[^1] == '\'')
        {
            throw new ReportException("invalid_sheet_name", "Sheet name cannot start or end with an apostrophe.");
        }
        if (sheetName.AsSpan().IndexOfAny(ForbiddenSheetNameChars) >= 0)
        {
            throw new ReportException("invalid_sheet_name", "Sheet name contains a forbidden character ([]:*?/\\).");
        }
        ValidateXmlCharacters(sheetName, "invalid_sheet_name");
    }

    private static void ValidateTextValue(string? text)
    {
        if (text is null)
        {
            throw new ReportException("invalid_value", "Text value cannot be null.");
        }
        if (text.Length > MaxCellTextLength)
        {
            throw new ReportException("cell_too_long", $"Cell text length {text.Length} exceeds the {MaxCellTextLength} limit.");
        }
        ValidateXmlCharacters(text, "invalid_character");
    }

    private static void ValidateXmlCharacters(string text, string invalidCharCode)
    {
        for (int i = 0; i < text.Length; i++)
        {
            char ch = text[i];
            if (char.IsHighSurrogate(ch))
            {
                if (i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                {
                    i++;
                    continue;
                }
                throw new ReportException(invalidCharCode, "String contains an unpaired high surrogate.");
            }
            if (char.IsLowSurrogate(ch))
            {
                throw new ReportException(invalidCharCode, "String contains an unpaired low surrogate.");
            }
            if (!IsValidXmlChar(ch))
            {
                throw new ReportException(invalidCharCode, $"String contains invalid XML 1.0 character U+{(int)ch:X4}.");
            }
        }
    }

    private static bool IsValidXmlChar(char ch)
    {
        return ch == 0x09
            || ch == 0x0A
            || ch == 0x0D
            || (ch >= 0x20 && ch <= 0xD7FF)
            || (ch >= 0xE000 && ch <= 0xFFFD);
    }

    private static string ExpandAndValidateFormula(
        string formulaTemplate,
        Dictionary<string, string> keyToColLetter,
        int excelRow)
    {
        if (string.IsNullOrWhiteSpace(formulaTemplate) || formulaTemplate.Length > MaxFormulaLength)
        {
            throw new ReportException("invalid_formula", "Formula is blank or exceeds maximum length.");
        }
        ValidateXmlCharacters(formulaTemplate, "invalid_character");

        string trimmed = formulaTemplate.Trim();
        if (trimmed.StartsWith('='))
        {
            throw new ReportException("invalid_formula", "Formula must not start with '='.");
        }

        var sb = new StringBuilder(formulaTemplate.Length + 16);
        string rowStr = excelRow.ToString(CultureInfo.InvariantCulture);
        int i = 0;
        while (i < formulaTemplate.Length)
        {
            char ch = formulaTemplate[i];
            if (ch == '{')
            {
                int close = formulaTemplate.IndexOf('}', i + 1);
                if (close < 0)
                {
                    throw new ReportException("invalid_formula", "Unclosed '{' placeholder in formula.");
                }
                string key = formulaTemplate[(i + 1)..close];
                if (key.Length == 0)
                {
                    throw new ReportException("invalid_formula", "Empty '{}' placeholder in formula.");
                }
                if (!keyToColLetter.TryGetValue(key, out string? colLetter))
                {
                    throw new ReportException("unknown_column", $"Formula references unknown column '{key}'.");
                }
                sb.Append(colLetter).Append(rowStr);
                i = close + 1;
            }
            else if (ch == '}')
            {
                throw new ReportException("invalid_formula", "Unmatched '}' in formula.");
            }
            else
            {
                sb.Append(ch);
                i++;
            }
        }

        string expanded = sb.ToString();
        ValidateExpandedFormulaSyntax(expanded);
        return expanded;
    }

    private static void ValidateExpandedFormulaSyntax(string expanded)
    {
        if (string.IsNullOrWhiteSpace(expanded))
        {
            throw new ReportException("invalid_formula", "Expanded formula cannot be empty.");
        }

        int parenDepth = 0;
        bool hasToken = false;
        foreach (char ch in expanded)
        {
            bool allowed =
                (ch >= 'A' && ch <= 'Z') ||
                (ch >= '0' && ch <= '9') ||
                ch is ':' or '+' or '-' or '*' or '/' or '(' or ')' or '.' or ' ' or ',' or '$';

            if (!allowed)
            {
                throw new ReportException("invalid_formula", $"Formula contains disallowed character '{ch}'.");
            }

            if (ch == '(')
            {
                parenDepth++;
            }
            else if (ch == ')')
            {
                parenDepth--;
                if (parenDepth < 0)
                {
                    throw new ReportException("invalid_formula", "Formula has unbalanced parentheses.");
                }
            }

            if (ch != ' ')
            {
                hasToken = true;
            }
        }

        if (!hasToken || parenDepth != 0)
        {
            throw new ReportException("invalid_formula", "Formula has unbalanced parentheses or no tokens.");
        }
    }

    private static void WriteContentTypes(XmlWriter writer, int sheetCount)
    {
        writer.WriteStartDocument(standalone: true);
        writer.WriteStartElement("Types", ContentTypesNs);

        WriteDefaultContentType(writer, "rels", "application/vnd.openxmlformats-package.relationships+xml");
        WriteDefaultContentType(writer, "xml", "application/xml");

        WriteOverrideContentType(writer, "/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
        WriteOverrideContentType(writer, "/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml");
        WriteOverrideContentType(writer, "/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
        WriteOverrideContentType(writer, "/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
        WriteOverrideContentType(writer, "/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");

        for (int i = 1; i <= sheetCount; i++)
        {
            WriteOverrideContentType(
                writer,
                $"/xl/worksheets/sheet{i}.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
        }

        writer.WriteEndElement();
        writer.WriteEndDocument();
    }

    private static void WriteDefaultContentType(XmlWriter writer, string extension, string contentType)
    {
        writer.WriteStartElement("Default", ContentTypesNs);
        writer.WriteAttributeString("Extension", extension);
        writer.WriteAttributeString("ContentType", contentType);
        writer.WriteEndElement();
    }

    private static void WriteOverrideContentType(XmlWriter writer, string partName, string contentType)
    {
        writer.WriteStartElement("Override", ContentTypesNs);
        writer.WriteAttributeString("PartName", partName);
        writer.WriteAttributeString("ContentType", contentType);
        writer.WriteEndElement();
    }

    private static void WriteRootRels(XmlWriter writer)
    {
        writer.WriteStartDocument(standalone: true);
        writer.WriteStartElement("Relationships", PackageRelsNs);

        WriteRelationship(writer, "rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "xl/workbook.xml");
        WriteRelationship(writer, "rId2", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml");
        WriteRelationship(writer, "rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "docProps/app.xml");

        writer.WriteEndElement();
        writer.WriteEndDocument();
    }

    private static void WriteAppProps(XmlWriter writer, List<PreparedSheet> sheets)
    {
        const string EpNs = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        const string VtNs = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

        writer.WriteStartDocument(standalone: true);
        writer.WriteStartElement("Properties", EpNs);
        writer.WriteAttributeString("xmlns", "vt", null, VtNs);

        writer.WriteElementString("Application", EpNs, "Trysac.OpenxmlReportWriter");

        writer.WriteStartElement("HeadingPairs", EpNs);
        writer.WriteStartElement("vector", VtNs);
        writer.WriteAttributeString("size", "2");
        writer.WriteAttributeString("baseType", "variant");
        writer.WriteStartElement("variant", VtNs);
        writer.WriteElementString("lpstr", VtNs, "Worksheets");
        writer.WriteEndElement();
        writer.WriteStartElement("variant", VtNs);
        writer.WriteElementString("i4", VtNs, sheets.Count.ToString(CultureInfo.InvariantCulture));
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndElement();

        writer.WriteStartElement("TitlesOfParts", EpNs);
        writer.WriteStartElement("vector", VtNs);
        writer.WriteAttributeString("size", sheets.Count.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("baseType", "lpstr");
        foreach (var sheet in sheets)
        {
            writer.WriteElementString("lpstr", VtNs, sheet.Template.SheetName);
        }
        writer.WriteEndElement();
        writer.WriteEndElement();

        writer.WriteEndElement();
        writer.WriteEndDocument();
    }

    private static void WriteCoreProps(XmlWriter writer)
    {
        const string CpNs = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        const string DcNs = "http://purl.org/dc/elements/1.1/";
        const string DcTermsNs = "http://purl.org/dc/terms/";
        const string XsiNs = "http://www.w3.org/2001/XMLSchema-instance";

        writer.WriteStartDocument(standalone: true);
        writer.WriteStartElement("cp", "coreProperties", CpNs);
        writer.WriteAttributeString("xmlns", "dc", null, DcNs);
        writer.WriteAttributeString("xmlns", "dcterms", null, DcTermsNs);
        writer.WriteAttributeString("xmlns", "xsi", null, XsiNs);

        writer.WriteStartElement("dcterms", "created", DcTermsNs);
        writer.WriteAttributeString("xsi", "type", XsiNs, "dcterms:W3CDTF");
        writer.WriteString(DeterministicIsoDate);
        writer.WriteEndElement();

        writer.WriteStartElement("dcterms", "modified", DcTermsNs);
        writer.WriteAttributeString("xsi", "type", XsiNs, "dcterms:W3CDTF");
        writer.WriteString(DeterministicIsoDate);
        writer.WriteEndElement();

        writer.WriteEndElement();
        writer.WriteEndDocument();
    }

    private static void WriteWorkbook(XmlWriter writer, List<PreparedSheet> sheets)
    {
        writer.WriteStartDocument(standalone: true);
        writer.WriteStartElement("workbook", SpreadsheetNs);
        writer.WriteAttributeString("xmlns", "r", null, RelationshipsNs);

        writer.WriteStartElement("sheets", SpreadsheetNs);
        for (int i = 0; i < sheets.Count; i++)
        {
            int id = i + 1;
            writer.WriteStartElement("sheet", SpreadsheetNs);
            writer.WriteAttributeString("name", sheets[i].Template.SheetName);
            writer.WriteAttributeString("sheetId", id.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("r", "id", RelationshipsNs, $"rId{id}");
            writer.WriteEndElement();
        }
        writer.WriteEndElement();

        writer.WriteStartElement("calcPr", SpreadsheetNs);
        writer.WriteAttributeString("calcId", "171027");
        writer.WriteAttributeString("fullCalcOnLoad", "1");
        writer.WriteEndElement();

        writer.WriteEndElement();
        writer.WriteEndDocument();
    }

    private static void WriteWorkbookRels(XmlWriter writer, int sheetCount)
    {
        writer.WriteStartDocument(standalone: true);
        writer.WriteStartElement("Relationships", PackageRelsNs);

        for (int i = 1; i <= sheetCount; i++)
        {
            WriteRelationship(
                writer,
                $"rId{i}",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                $"worksheets/sheet{i}.xml");
        }

        int stylesId = sheetCount + 1;
        int sstId = sheetCount + 2;
        WriteRelationship(
            writer,
            $"rId{stylesId}",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
            "styles.xml");
        WriteRelationship(
            writer,
            $"rId{sstId}",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
            "sharedStrings.xml");

        writer.WriteEndElement();
        writer.WriteEndDocument();
    }

    private static void WriteRelationship(XmlWriter writer, string id, string type, string target)
    {
        writer.WriteStartElement("Relationship", PackageRelsNs);
        writer.WriteAttributeString("Id", id);
        writer.WriteAttributeString("Type", type);
        writer.WriteAttributeString("Target", target);
        writer.WriteEndElement();
    }

    private static void WriteStyles(XmlWriter writer, StylesBuilder styles)
    {
        writer.WriteStartDocument(standalone: true);
        writer.WriteStartElement("styleSheet", SpreadsheetNs);

        if (styles.CustomNumFmts.Count > 0)
        {
            writer.WriteStartElement("numFmts", SpreadsheetNs);
            writer.WriteAttributeString("count", styles.CustomNumFmts.Count.ToString(CultureInfo.InvariantCulture));
            foreach (var (id, code) in styles.CustomNumFmts)
            {
                writer.WriteStartElement("numFmt", SpreadsheetNs);
                writer.WriteAttributeString("numFmtId", id.ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("formatCode", code);
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        writer.WriteStartElement("fonts", SpreadsheetNs);
        writer.WriteAttributeString("count", "2");

        writer.WriteStartElement("font", SpreadsheetNs);
        writer.WriteStartElement("sz", SpreadsheetNs);
        writer.WriteAttributeString("val", "11");
        writer.WriteEndElement();
        writer.WriteStartElement("name", SpreadsheetNs);
        writer.WriteAttributeString("val", "Calibri");
        writer.WriteEndElement();
        writer.WriteEndElement();

        writer.WriteStartElement("font", SpreadsheetNs);
        writer.WriteStartElement("b", SpreadsheetNs);
        writer.WriteEndElement();
        writer.WriteStartElement("sz", SpreadsheetNs);
        writer.WriteAttributeString("val", "11");
        writer.WriteEndElement();
        writer.WriteStartElement("name", SpreadsheetNs);
        writer.WriteAttributeString("val", "Calibri");
        writer.WriteEndElement();
        writer.WriteEndElement();

        writer.WriteEndElement();

        writer.WriteStartElement("fills", SpreadsheetNs);
        writer.WriteAttributeString("count", "2");
        writer.WriteStartElement("fill", SpreadsheetNs);
        writer.WriteStartElement("patternFill", SpreadsheetNs);
        writer.WriteAttributeString("patternType", "none");
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteStartElement("fill", SpreadsheetNs);
        writer.WriteStartElement("patternFill", SpreadsheetNs);
        writer.WriteAttributeString("patternType", "gray125");
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndElement();

        writer.WriteStartElement("borders", SpreadsheetNs);
        writer.WriteAttributeString("count", "1");
        writer.WriteStartElement("border", SpreadsheetNs);
        writer.WriteStartElement("left", SpreadsheetNs);
        writer.WriteEndElement();
        writer.WriteStartElement("right", SpreadsheetNs);
        writer.WriteEndElement();
        writer.WriteStartElement("top", SpreadsheetNs);
        writer.WriteEndElement();
        writer.WriteStartElement("bottom", SpreadsheetNs);
        writer.WriteEndElement();
        writer.WriteStartElement("diagonal", SpreadsheetNs);
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndElement();

        writer.WriteStartElement("cellStyleXfs", SpreadsheetNs);
        writer.WriteAttributeString("count", "1");
        writer.WriteStartElement("xf", SpreadsheetNs);
        writer.WriteAttributeString("numFmtId", "0");
        writer.WriteAttributeString("fontId", "0");
        writer.WriteAttributeString("fillId", "0");
        writer.WriteAttributeString("borderId", "0");
        writer.WriteEndElement();
        writer.WriteEndElement();

        writer.WriteStartElement("cellXfs", SpreadsheetNs);
        writer.WriteAttributeString("count", styles.CellXfs.Count.ToString(CultureInfo.InvariantCulture));
        foreach (var xf in styles.CellXfs)
        {
            writer.WriteStartElement("xf", SpreadsheetNs);
            writer.WriteAttributeString("numFmtId", xf.NumFmtId.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("fontId", xf.FontId.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("fillId", "0");
            writer.WriteAttributeString("borderId", "0");
            writer.WriteAttributeString("xfId", "0");
            if (xf.NumFmtId != 0)
            {
                writer.WriteAttributeString("applyNumberFormat", "1");
            }
            if (xf.FontId != 0)
            {
                writer.WriteAttributeString("applyFont", "1");
            }
            writer.WriteEndElement();
        }
        writer.WriteEndElement();

        writer.WriteStartElement("cellStyles", SpreadsheetNs);
        writer.WriteAttributeString("count", "1");
        writer.WriteStartElement("cellStyle", SpreadsheetNs);
        writer.WriteAttributeString("name", "Normal");
        writer.WriteAttributeString("xfId", "0");
        writer.WriteAttributeString("builtinId", "0");
        writer.WriteEndElement();
        writer.WriteEndElement();

        writer.WriteEndElement();
        writer.WriteEndDocument();
    }

    private static void WriteSharedStrings(XmlWriter writer, SharedStringsBuilder sharedStrings)
    {
        writer.WriteStartDocument(standalone: true);
        writer.WriteStartElement("sst", SpreadsheetNs);
        writer.WriteAttributeString("count", sharedStrings.TotalReferenceCount.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("uniqueCount", sharedStrings.UniqueStrings.Count.ToString(CultureInfo.InvariantCulture));

        foreach (string text in sharedStrings.UniqueStrings)
        {
            writer.WriteStartElement("si", SpreadsheetNs);
            writer.WriteStartElement("t", SpreadsheetNs);
            if (NeedsXmlSpacePreserve(text))
            {
                writer.WriteAttributeString("xml", "space", null, "preserve");
            }
            writer.WriteString(text);
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        writer.WriteEndElement();
        writer.WriteEndDocument();
    }

    private static bool NeedsXmlSpacePreserve(string text)
    {
        if (text.Length == 0)
        {
            return false;
        }
        return char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[^1]) || text.Contains('\n') || text.Contains('\r') || text.Contains('\t');
    }

    private static void WriteWorksheet(XmlWriter writer, PreparedSheet sheet)
    {
        int colCount = sheet.Columns.Count;
        int dataRowCount = sheet.Rows.Count;
        int totalRowCount = 1 + dataRowCount + (sheet.HasTotalsRow ? 1 : 0);
        string lastColLetter = sheet.Columns[^1].ColumnLetter;
        string dimensionRef = totalRowCount == 1 && colCount == 1
            ? "A1"
            : $"A1:{lastColLetter}{totalRowCount.ToString(CultureInfo.InvariantCulture)}";

        writer.WriteStartDocument(standalone: true);
        writer.WriteStartElement("worksheet", SpreadsheetNs);

        writer.WriteStartElement("dimension", SpreadsheetNs);
        writer.WriteAttributeString("ref", dimensionRef);
        writer.WriteEndElement();

        if (sheet.Template.FreezeHeader)
        {
            writer.WriteStartElement("sheetViews", SpreadsheetNs);
            writer.WriteStartElement("sheetView", SpreadsheetNs);
            writer.WriteAttributeString("workbookViewId", "0");
            writer.WriteStartElement("pane", SpreadsheetNs);
            writer.WriteAttributeString("ySplit", "1");
            writer.WriteAttributeString("topLeftCell", "A2");
            writer.WriteAttributeString("activePane", "bottomLeft");
            writer.WriteAttributeString("state", "frozen");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        writer.WriteStartElement("cols", SpreadsheetNs);
        foreach (var col in sheet.Columns)
        {
            string idxStr = col.ColumnIndex.ToString(CultureInfo.InvariantCulture);
            writer.WriteStartElement("col", SpreadsheetNs);
            writer.WriteAttributeString("min", idxStr);
            writer.WriteAttributeString("max", idxStr);
            writer.WriteAttributeString("width", col.Spec.Width.ToString("G17", CultureInfo.InvariantCulture));
            writer.WriteAttributeString("customWidth", "1");
            writer.WriteEndElement();
        }
        writer.WriteEndElement();

        writer.WriteStartElement("sheetData", SpreadsheetNs);

        // Header row (row 1)
        writer.WriteStartElement("row", SpreadsheetNs);
        writer.WriteAttributeString("r", "1");
        foreach (var col in sheet.Columns)
        {
            writer.WriteStartElement("c", SpreadsheetNs);
            writer.WriteAttributeString("r", $"{col.ColumnLetter}1");
            writer.WriteAttributeString("s", StylesBuilder.HeaderStyleIndex.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("t", "s");
            writer.WriteElementString("v", SpreadsheetNs, col.HeaderSharedStringId.ToString(CultureInfo.InvariantCulture));
            writer.WriteEndElement();
        }
        writer.WriteEndElement();

        // Data rows
        for (int r = 0; r < dataRowCount; r++)
        {
            int excelRow = r + 2;
            string rowStr = excelRow.ToString(CultureInfo.InvariantCulture);
            var cells = sheet.Rows[r];

            writer.WriteStartElement("row", SpreadsheetNs);
            writer.WriteAttributeString("r", rowStr);
            for (int c = 0; c < colCount; c++)
            {
                var cell = cells[c];
                if (cell.IsEmpty)
                {
                    continue;
                }
                var col = sheet.Columns[c];
                string cellRef = string.Concat(col.ColumnLetter, rowStr);
                WriteCell(writer, cellRef, cell);
            }
            writer.WriteEndElement();
        }

        // Totals row
        if (sheet.HasTotalsRow)
        {
            int totalsExcelRow = dataRowCount + 2;
            int lastDataExcelRow = dataRowCount + 1;
            string totalsRowStr = totalsExcelRow.ToString(CultureInfo.InvariantCulture);
            string lastDataRowStr = lastDataExcelRow.ToString(CultureInfo.InvariantCulture);

            writer.WriteStartElement("row", SpreadsheetNs);
            writer.WriteAttributeString("r", totalsRowStr);
            foreach (var col in sheet.Columns)
            {
                if (col.Spec.Total == Aggregate.None)
                {
                    continue;
                }
                string funcName = col.Spec.Total switch
                {
                    Aggregate.Sum => "SUM",
                    Aggregate.Average => "AVERAGE",
                    Aggregate.Count => "COUNTA",
                    _ => throw new ReportException("invalid_value", $"Unsupported aggregate '{col.Spec.Total}'.")
                };
                string rangeFormula = $"{funcName}({col.ColumnLetter}2:{col.ColumnLetter}{lastDataRowStr})";
                string cellRef = string.Concat(col.ColumnLetter, totalsRowStr);
                writer.WriteStartElement("c", SpreadsheetNs);
                writer.WriteAttributeString("r", cellRef);
                if (col.TotalsStyleIndex != 0)
                {
                    writer.WriteAttributeString("s", col.TotalsStyleIndex.ToString(CultureInfo.InvariantCulture));
                }
                writer.WriteElementString("f", SpreadsheetNs, rangeFormula);
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        writer.WriteEndElement(); // sheetData

        if (sheet.Template.AutoFilter)
        {
            int filterEndRow = 1 + dataRowCount;
            string filterRef = $"A1:{lastColLetter}{filterEndRow.ToString(CultureInfo.InvariantCulture)}";
            writer.WriteStartElement("autoFilter", SpreadsheetNs);
            writer.WriteAttributeString("ref", filterRef);
            writer.WriteEndElement();
        }

        writer.WriteEndElement(); // worksheet
        writer.WriteEndDocument();
    }

    private static void WriteCell(XmlWriter writer, string cellRef, PreparedCell cell)
    {
        writer.WriteStartElement("c", SpreadsheetNs);
        writer.WriteAttributeString("r", cellRef);
        if (cell.StyleIndex != 0)
        {
            writer.WriteAttributeString("s", cell.StyleIndex.ToString(CultureInfo.InvariantCulture));
        }
        if (cell.CellType is not null)
        {
            writer.WriteAttributeString("t", cell.CellType);
        }
        if (cell.IsFormula)
        {
            writer.WriteElementString("f", SpreadsheetNs, cell.ValueOrFormula);
        }
        else
        {
            writer.WriteElementString("v", SpreadsheetNs, cell.ValueOrFormula);
        }
        writer.WriteEndElement();
    }

    private sealed record CellXfEntry(int NumFmtId, int FontId);

    private sealed class StylesBuilder
    {
        public const int DefaultStyleIndex = 0;
        public const int HeaderStyleIndex = 1;

        private int _nextCustomNumFmtId = 164;
        private readonly Dictionary<string, int> _customNumFmtByCode = new(StringComparer.Ordinal);
        private readonly Dictionary<CellXfEntry, int> _xfIndices = new();

        public List<(int Id, string Code)> CustomNumFmts { get; } = [];
        public List<CellXfEntry> CellXfs { get; } = [];

        public StylesBuilder()
        {
            RegisterXf(new CellXfEntry(0, 0)); // index 0: General, Regular
            RegisterXf(new CellXfEntry(0, 1)); // index 1: General, Bold (Header)
        }

        public int GetDataStyleIndex(CellKind kind, string? numberFormat)
        {
            int numFmtId = ResolveNumFmtId(kind, numberFormat);
            return RegisterXf(new CellXfEntry(numFmtId, 0));
        }

        public int GetTotalsStyleIndex(CellKind kind, string? numberFormat, Aggregate total)
        {
            if (total == Aggregate.Count)
            {
                return RegisterXf(new CellXfEntry(0, 1));
            }
            int numFmtId = ResolveNumFmtId(kind, numberFormat);
            return RegisterXf(new CellXfEntry(numFmtId, 1));
        }

        private int ResolveNumFmtId(CellKind kind, string? numberFormat)
        {
            if (!string.IsNullOrEmpty(numberFormat))
            {
                ValidateTextValue(numberFormat);
                if (!_customNumFmtByCode.TryGetValue(numberFormat, out int customId))
                {
                    customId = _nextCustomNumFmtId++;
                    _customNumFmtByCode[numberFormat] = customId;
                    CustomNumFmts.Add((customId, numberFormat));
                }
                return customId;
            }

            return kind == CellKind.Date ? 14 : 0;
        }

        private int RegisterXf(CellXfEntry entry)
        {
            if (!_xfIndices.TryGetValue(entry, out int index))
            {
                index = CellXfs.Count;
                CellXfs.Add(entry);
                _xfIndices[entry] = index;
            }
            return index;
        }
    }

    private sealed class SharedStringsBuilder
    {
        private readonly Dictionary<string, int> _indexByString = new(StringComparer.Ordinal);

        public List<string> UniqueStrings { get; } = [];
        public int TotalReferenceCount { get; private set; }

        public int Intern(string value)
        {
            TotalReferenceCount++;
            if (!_indexByString.TryGetValue(value, out int id))
            {
                id = UniqueStrings.Count;
                UniqueStrings.Add(value);
                _indexByString[value] = id;
            }
            return id;
        }
    }

    private sealed record PreparedColumn(
        ColumnSpec Spec,
        int ColumnIndex,
        string ColumnLetter,
        int HeaderSharedStringId,
        int DataStyleIndex,
        int TotalsStyleIndex,
        string? SampleExpandedFormula);

    private readonly record struct PreparedCell(
        bool IsEmpty,
        bool IsFormula,
        string? CellType,
        int StyleIndex,
        string ValueOrFormula)
    {
        public static PreparedCell Empty => new(true, false, null, 0, string.Empty);

        public static PreparedCell ForSharedString(int sstId, int styleIndex) =>
            new(false, false, "s", styleIndex, sstId.ToString(CultureInfo.InvariantCulture));

        public static PreparedCell ForNumber(string formattedNumber, int styleIndex) =>
            new(false, false, null, styleIndex, formattedNumber);

        public static PreparedCell ForBoolean(bool value, int styleIndex) =>
            new(false, false, "b", styleIndex, value ? "1" : "0");

        public static PreparedCell ForFormula(string formula, int styleIndex) =>
            new(false, true, null, styleIndex, formula);
    }

    private sealed record PreparedSheet(
        ReportTemplate Template,
        List<PreparedColumn> Columns,
        List<PreparedCell[]> Rows,
        bool HasTotalsRow);
}
