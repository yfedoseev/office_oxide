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

    // Every option, as every binding names it.
    private static readonly string[] OptionNames =
    {
        "include_hidden_text", "include_deleted_text", "on_incomplete", "lazy_pictures",
        "package_entries", "package_bytes", "part_bytes", "inflate_ratio",
        "embedded_depth", "embedded_packages", "text_chars", "memory_bytes",
    };

    [Fact]
    public void TheOptionNamesListEveryOptionAndSetTakesEach()
    {
        Assert.Equal(OptionNames.OrderBy(n => n, StringComparer.Ordinal),
            ParseOptions.Names().OrderBy(n => n, StringComparer.Ordinal));
        foreach (var name in OptionNames)
        {
            using var options = new ParseOptions().Set(name, 1UL);
        }
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

    [Fact]
    public void OptionsReportTheirLimitsAndTheBoundsDerivedFromThem()
    {
        using var options = new ParseOptions().Set("memory_bytes", 1000UL).Set("part_bytes", ParseOptions.NoLimit);
        var (limits, derived) = options.Limits();
        Assert.Equal(1000UL, limits["memory_bytes"]);
        Assert.Equal(ParseOptions.NoLimit, limits["part_bytes"]);
        Assert.Equal(500UL, derived["write_bytes"]);
        Assert.Equal(8, limits.Count);
    }

    [Fact]
    public void OutputsAsUtf8BytesStreamsAndFilesEqualTheStrings()
    {
        using var w = new XlsxWriter();
        var s = w.AddSheet("S");
        w.SetCell(s, 0, 0, "omega \U0001D74E and \u00e9");
        using var doc = Document.FromBytes(w.ToBytes(), "xlsx");
        var forms = new Dictionary<string, string>
        {
            ["text"] = doc.PlainText(),
            ["markdown"] = doc.ToMarkdown(),
            ["markdown_images"] = doc.ToMarkdownWithImages(),
            ["html"] = doc.ToHtml(),
            ["ir_json"] = doc.ToIrJson(),
        };
        var dir = Directory.CreateTempSubdirectory("oo-out-");
        try
        {
            foreach (var (kind, text) in forms)
            {
                Assert.Contains("\U0001D74E", text);
                var expected = System.Text.Encoding.UTF8.GetBytes(text);
                Assert.Equal(expected, doc.ToUtf8Bytes(kind));
                using var stream = new MemoryStream();
                doc.WriteTo(stream, kind);
                Assert.Equal(expected, stream.ToArray());
                var path = Path.Combine(dir.FullName, kind);
                doc.WriteTo(path, kind);
                Assert.Equal(expected, File.ReadAllBytes(path));
            }
            Assert.Throws<OfficeOxideException>(() => doc.ToUtf8Bytes("pdf"));
        }
        finally
        {
            dir.Delete(recursive: true);
        }
    }
}
