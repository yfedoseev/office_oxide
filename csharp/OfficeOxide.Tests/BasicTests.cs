using Xunit;

namespace OfficeOxide.Tests;

public class BasicTests
{
    private const string Fixture = "/tmp/ffi_smoke.docx";

    [Fact]
    public void Version_IsNonEmpty()
    {
        Assert.False(string.IsNullOrEmpty(Document.Version));
    }

    [Fact]
    public void DetectFormat_ReturnsExpected()
    {
        Assert.Equal("docx", Document.DetectFormat("a.docx"));
        Assert.Null(Document.DetectFormat("a.unknown"));
    }

    [Fact]
    public void Document_OpenAndExtract()
    {
        if (!File.Exists(Fixture)) return; // skipped when fixture missing
        using var doc = Document.Open(Fixture);
        Assert.Equal("docx", doc.Format);
        Assert.Contains("Hello", doc.PlainText());
        Assert.Contains("# ", doc.ToMarkdown());
    }

    [Fact]
    public void Editable_ReplaceText_Roundtrip()
    {
        if (!File.Exists(Fixture)) return;
        var outPath = "/tmp/ffi_smoke_csharp_edit.docx";
        using (var ed = EditableDocument.Open(Fixture))
        {
            var n = ed.ReplaceText("Hello", "G'day");
            Assert.True(n >= 1);
            ed.Save(outPath);
        }
        using var reopened = Document.Open(outPath);
        Assert.Contains("G'day", reopened.PlainText());
    }
}
