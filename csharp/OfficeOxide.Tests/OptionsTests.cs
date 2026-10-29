using Xunit;

namespace OfficeOxide.Tests;

/// <summary>Per-document options: limits, and refusing an incomplete result. Self-contained.</summary>
public class OptionsTests
{
    // 200 cells of 200 characters: 40,000 characters of text from one small sheet.
    private static byte[] BigSheet()
    {
        using var w = new XlsxWriter();
        w.AddSheet("S");
        for (uint i = 0; i < 200; i++) w.SetCell(0, i, 0, new string('x', 200));
        return w.ToBytes();
    }

    [Fact]
    public void ALimitAppliesToThatDocumentOnly()
    {
        var data = BigSheet();
        using var options = new ParseOptions().Set("text_chars", 5000UL);
        using var small = Document.FromBytes(data, "xlsx", options);
        using var plain = Document.FromBytes(data, "xlsx");
        Assert.True(small.PlainText().Length < plain.PlainText().Length);
    }

    [Fact]
    public void DamageListsACutAndStrictOptionsRefuseIt()
    {
        var data = BigSheet();
        using var options = new ParseOptions().Set("text_chars", 5000UL);
        using (var doc = Document.FromBytes(data, "xlsx", options))
            Assert.Contains(doc.Damage(), d => d.Kind == "truncated" && d.IsLoss);

        options.OnIncomplete = true;
        var ex = Assert.Throws<OfficeOxideException>(() => Document.FromBytes(data, "xlsx", options));
        Assert.True(ex.IsIncomplete);
        Assert.Equal(OfficeOxideErrorCode.Incomplete, ex.Code);
    }

    [Fact]
    public void StrictOptionsAcceptASoundDocumentAndBadNamesAreRefused()
    {
        using var w = new XlsxWriter();
        w.AddSheet("S");
        w.SetCell(0, 0, 0, "hello");
        using var strict = new ParseOptions(strict: true);
        using var doc = Document.FromBytes(w.ToBytes(), "xlsx", strict);
        Assert.Contains("hello", doc.PlainText());
        Assert.DoesNotContain(doc.Damage(), d => d.IsLoss);
        Assert.Throws<ArgumentException>(() => strict.Set("no_such_limit", 1UL));
        Assert.Contains("part_bytes", ParseOptions.Names());
    }
}
