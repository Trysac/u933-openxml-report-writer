using FsCheck;
using FsCheck.Fluent;
using FsCheck.Xunit;

namespace Trysac.OpenxmlReportWriter.Tests;

/// <summary>Property-based tests for CellRef round-trips and Text normalization.</summary>
public class TextPropertiesTests
{
    [Property]
    public Property CellRefRoundTripsValidCoordinates()
    {
        var gen = from row in Gen.Choose(1, CellRef.MaxRow)
                  from col in Gen.Choose(1, CellRef.MaxColumn)
                  select (row, col);

        return Prop.ForAll(Arb.From(gen), pair =>
        {
            string a1 = CellRef.ToA1(pair.row, pair.col);
            var parsed = CellRef.Parse(a1);
            return parsed.Row == pair.row && parsed.Column == pair.col;
        });
    }

    [Property]
    public Property NormalizeIsIdempotent(NonNull<string> raw)
    {
        string trimmed = raw.Get.Trim();
        if (trimmed.Length == 0)
        {
            return true.ToProperty();
        }
        string once = Text.Normalize(raw.Get);
        string twice = Text.Normalize(once);
        return (once == twice).ToProperty();
    }
}
