// 02_create_from_markdown: turn one Markdown source into a DOCX, an XLSX and a PPTX.
//
// Run: dotnet run --project examples/csharp/02_create_from_markdown/02_create_from_markdown.csproj
using OfficeOxide;

const string Markdown = """
    # Quarterly Report

    Revenue grew in every region.

    | Region | Q1  | Q2  |
    |--------|-----|-----|
    | North  | 120 | 135 |
    | South  | 95  | 110 |

    """;

static void Check(bool condition, string message)
{
    if (condition) return;
    Console.Error.WriteLine($"FAIL: {message}");
    Environment.Exit(1);
}

static void Has(string text, string expected) =>
    Check(text.Contains(expected), $"\"{expected}\" is missing from the text");

// CreateFromMarkdown writes to a path, so the file lives in a private temp
// file just long enough to read its bytes back.
static byte[] MarkdownToBytes(string markdown, string format)
{
    var path = Path.Combine(Path.GetTempPath(), $"oo_example_{Guid.NewGuid():N}.{format}");
    try
    {
        Document.CreateFromMarkdown(markdown, format, path);
        return File.ReadAllBytes(path);
    }
    finally
    {
        File.Delete(path);
    }
}

// Opens bytes as the given format and returns the plain text.
static string TextOf(byte[] data, string format)
{
    using var doc = Document.FromBytes(data, format);
    return doc.PlainText();
}

Console.WriteLine("step 1: write the Markdown source");

Console.WriteLine("step 2: create a DOCX, an XLSX and a PPTX from it");
string[] formats = ["docx", "xlsx", "pptx"];
var created = formats.ToDictionary(f => f, f => MarkdownToBytes(Markdown, f));

Console.WriteLine("step 3: open each and check the text");
foreach (var format in formats)
{
    var text = TextOf(created[format], format);
    Console.WriteLine($"  format: {format}");
    Console.WriteLine($"  chars: {text.Length}");
    Has(text, "North");
}

Console.WriteLine("OK 02_create_from_markdown");
