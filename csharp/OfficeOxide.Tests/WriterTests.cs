using Xunit;

namespace OfficeOxide.Tests;

/// <summary>Self-contained writer and in-memory editing tests: no fixture file needed.</summary>
public class WriterTests
{
    private static readonly byte[] Png =
    {
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
        0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
        0x00, 0x03, 0x01, 0x01, 0x00, 0x18, 0xdd, 0x8d, 0xb0, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
        0x44, 0xae, 0x42, 0x60, 0x82,
    };

    // The PPTX entry points were declared void although the FFI returns a
    // status, so a write to a missing slide reported success.
    [Fact]
    public void PptxWriter_MissingSlide_Throws()
    {
        using var w = new PptxWriter();
        w.AddSlide();
        Assert.Throws<InvalidOperationException>(() => w.SetSlideTitle(99, "x"));
        Assert.Throws<InvalidOperationException>(() => w.AddSlideText(99, "x"));
        Assert.Throws<InvalidOperationException>(() => w.AddSlideImage(99, Png, "png", 0, 0, 1, 1));
        Assert.Throws<InvalidOperationException>(() => w.AddSlideImage(0, Png, "bmp", 0, 0, 1, 1));
        w.SetSlideTitle(0, "ok");
        w.AddSlideImage(0, Png, "png", 0, 0, 914400, 914400);
        using var doc = Document.FromBytes(w.ToBytes(), "pptx");
        Assert.Contains("[image-base64:", doc.ToMarkdownWithImages());
        Assert.DoesNotContain("[image-base64:", doc.ToMarkdown());
    }

    [Fact]
    public void XlsxWriter_FormulasBooleansAndExactIntegers()
    {
        using var w = new XlsxWriter();
        w.AddSheet("S");
        w.SetCell(0, 0, 0, true);
        w.SetCell(0, 1, 0, 2);
        w.SetFormula(0, 2, 0, "=SUM(A2:A2)");
        w.SetCell(0, 3, 0, new Formula("A2*2"));
        w.SetCell(0, 4, 0, "=literal");
        w.SetCell(0, 5, 0, 1L << 53);
        Assert.Throws<ArgumentOutOfRangeException>(() => w.SetCell(0, 6, 0, (1L << 53) + 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => w.SetCell(0, 6, 0, ulong.MaxValue));
        Assert.Throws<ArgumentOutOfRangeException>(() => w.SetCellStyled(0, 6, 0, long.MinValue, true));
        Assert.Throws<InvalidOperationException>(() => w.SetCell(5, 0, 0, "x"));

        using var doc = Document.FromBytes(w.ToBytes(), "xlsx");
        var ir = doc.ToIrJson();
        Assert.Contains("SUM(A2:A2)", ir);
        Assert.Contains("A2*2", ir);
        Assert.Contains("=literal", ir);
        Assert.Contains("TRUE", doc.PlainText());
    }

    // Editing an in-memory document forced a temporary-file round trip.
    [Fact]
    public void EditableDocument_FromBytes_RoundTrip()
    {
        using var w = new XlsxWriter();
        w.AddSheet("S");
        w.SetCell(0, 0, 0, "old");
        using var ed = EditableDocument.FromBytes(w.ToBytes(), "xlsx");
        ed.SetCell(0, "A1", "new");
        var ex = Assert.Throws<OfficeOxideException>(() => ed.ReplaceText("", "x"));
        Assert.Equal(OfficeOxideErrorCode.InvalidArg, ex.Code);
        using var doc = Document.FromBytes(ed.SaveToBytes(), "xlsx");
        Assert.Contains("new", doc.PlainText());
    }
}
