namespace Trysac.OpenxmlReportWriter;

/// <summary>Specifies the kind of data held by a column in a report template.</summary>
public enum CellKind
{
    /// <summary>Shared string text cell.</summary>
    Text,

    /// <summary>Numeric cell (int, long, double, decimal).</summary>
    Number,

    /// <summary>Date or date-time cell stored as an Excel serial number.</summary>
    Date,

    /// <summary>Boolean cell stored as 1 or 0 with type "b".</summary>
    Boolean,

    /// <summary>Formula cell evaluated per row from a template expression.</summary>
    Formula
}

/// <summary>Specifies the summary function placed in the totals row for a column.</summary>
public enum Aggregate
{
    /// <summary>No aggregate formula in the totals row.</summary>
    None,

    /// <summary>SUM over the data rows of the column.</summary>
    Sum,

    /// <summary>AVERAGE over the data rows of the column.</summary>
    Average,

    /// <summary>COUNTA over the data rows of the column.</summary>
    Count
}

/// <summary>Defines a single column in a report worksheet.</summary>
/// <param name="Header">The column header text written in row 1.</param>
/// <param name="Key">The dictionary key used to look up values in data rows and formula placeholders.</param>
/// <param name="Kind">The cell data kind for this column.</param>
/// <param name="NumberFormat">Optional Excel number format code (for example "#,##0.00" or "yyyy-mm-dd").</param>
/// <param name="Formula">Formula template with <c>{Key}</c> placeholders when <paramref name="Kind"/> is <see cref="CellKind.Formula"/>.</param>
/// <param name="Width">Column width in character units (defaults to 12).</param>
/// <param name="Total">Aggregate function for the optional totals row.</param>
public sealed record ColumnSpec(
    string Header,
    string Key,
    CellKind Kind,
    string? NumberFormat = null,
    string? Formula = null,
    double Width = 12,
    Aggregate Total = Aggregate.None);

/// <summary>Defines the layout and options of a worksheet report.</summary>
/// <param name="SheetName">Worksheet tab name (1 to 31 valid characters).</param>
/// <param name="Columns">Ordered list of column specifications.</param>
/// <param name="TotalsRow">Whether to append a totals row after the data rows when at least one data row exists.</param>
/// <param name="FreezeHeader">Whether to freeze the top header row (ySplit = 1).</param>
/// <param name="AutoFilter">Whether to add an autoFilter covering the header and data rows.</param>
public sealed record ReportTemplate(
    string SheetName,
    IReadOnlyList<ColumnSpec> Columns,
    bool TotalsRow = false,
    bool FreezeHeader = true,
    bool AutoFilter = true);

/// <summary>Pairs a <see cref="ReportTemplate"/> with its data rows.</summary>
/// <param name="Template">The report worksheet template.</param>
/// <param name="Rows">The sequence of row dictionaries mapping column keys to cell values.</param>
public sealed record SheetData(
    ReportTemplate Template,
    IReadOnlyList<IReadOnlyDictionary<string, object?>> Rows);
