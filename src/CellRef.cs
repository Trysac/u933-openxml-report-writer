using System.Globalization;

namespace Trysac.OpenxmlReportWriter;

/// <summary>Converts between 1-based row/column coordinates and Excel A1 notation.</summary>
public static class CellRef
{
    /// <summary>Maximum row number in SpreadsheetML (1,048,576).</summary>
    public const int MaxRow = 1_048_576;

    /// <summary>Maximum column number in SpreadsheetML (16,384, column XFD).</summary>
    public const int MaxColumn = 16_384;

    /// <summary>Converts a 1-based row and column index into an A1 cell reference.</summary>
    /// <param name="row">1-based row index (1 to 1,048,576).</param>
    /// <param name="column">1-based column index (1 to 16,384).</param>
    /// <returns>The uppercase A1 cell reference (for example, "A1" or "XFD1048576").</returns>
    /// <exception cref="ReportException">Thrown with code <c>invalid_reference</c> when row or column is out of bounds.</exception>
    public static string ToA1(int row, int column)
    {
        if (row < 1 || row > MaxRow)
        {
            throw new ReportException("invalid_reference", $"Row {row} is outside the valid Excel range 1..{MaxRow}.");
        }
        if (column < 1 || column > MaxColumn)
        {
            throw new ReportException("invalid_reference", $"Column {column} is outside the valid Excel range 1..{MaxColumn}.");
        }

        return string.Concat(ColumnToLetters(column), row.ToString(CultureInfo.InvariantCulture));
    }

    /// <summary>Parses an A1 cell reference into its 1-based row and column coordinates.</summary>
    /// <param name="a1">The A1 reference string (for example, "AB12").</param>
    /// <returns>A tuple of 1-based <c>(Row, Column)</c>.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="a1"/> is null.</exception>
    /// <exception cref="ReportException">Thrown with code <c>invalid_reference</c> if the reference is malformed or out of bounds.</exception>
    public static (int Row, int Column) Parse(string a1)
    {
        ArgumentNullException.ThrowIfNull(a1);
        if (a1.Length < 2 || a1.Length > 10)
        {
            throw new ReportException("invalid_reference", "Cell reference must contain column letters followed by row digits.");
        }

        int i = 0;
        int col = 0;
        while (i < a1.Length && a1[i] is >= 'A' and <= 'Z')
        {
            if (i >= 3)
            {
                throw new ReportException("invalid_reference", "Column letters in cell reference exceed maximum length.");
            }
            col = (col * 26) + (a1[i] - 'A' + 1);
            i++;
        }

        if (i == 0 || i == a1.Length || col < 1 || col > MaxColumn)
        {
            throw new ReportException("invalid_reference", "Cell reference has an invalid column component.");
        }

        if (a1[i] == '0')
        {
            throw new ReportException("invalid_reference", "Row number in cell reference cannot start with zero.");
        }

        long row = 0;
        while (i < a1.Length)
        {
            char ch = a1[i];
            if (ch is < '0' or > '9')
            {
                throw new ReportException("invalid_reference", "Cell reference contains invalid characters in row component.");
            }
            row = (row * 10) + (ch - '0');
            if (row > MaxRow)
            {
                throw new ReportException("invalid_reference", "Row number in cell reference exceeds 1,048,576.");
            }
            i++;
        }

        return ((int)row, col);
    }

    internal static string ColumnToLetters(int column)
    {
        Span<char> buffer = stackalloc char[3];
        int pos = 3;
        int c = column;
        while (c > 0)
        {
            c--;
            buffer[--pos] = (char)('A' + (c % 26));
            c /= 26;
        }
        return new string(buffer[pos..]);
    }
}
